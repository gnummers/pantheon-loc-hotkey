import os
import re
import time
import json
import socket
import subprocess
from pathlib import Path
from typing import Optional, Tuple

import requests
import websocket  # websocket-client
import keyboard
import pyautogui
import pyperclip
import pygetwindow as gw
import ctypes

# =========================
# Config
# =========================
MAP_URL = "https://shalazam.info/maps/1"

# Choose any free port; we reuse one port for the chosen browser
DEBUG_PORT = 9222

# Per-browser profiles (kept separate so sessions don’t collide)
EDGE_PROFILE_DIR   = r"C:\msedgeprofile_panth"
CHROME_PROFILE_DIR = r"C:\chromeprofile_panth"

# Allow all origins for simplicity/compat
ALLOW_ORIGINS = "*"

PANTHEON_WINDOW_TITLES = ["Pantheon", "Pantheon: Rise of the Fallen"]

# Timings
TYPING_DELAY    = 0.05
CHAT_WAKE_DELAY = 0.30
AFTER_LOC_WAIT  = 1.10
DEVTOOLS_TIMEOUT = 45.0

# Hotkeys
HOTKEY_TRIGGER  = "ctrl+l"
HOTKEY_QUIT     = "ctrl+q"
HOTKEY_DEBOUNCE = 0.35


# =========================
# Small utils
# =========================
def is_port_open(host: str, port: int, timeout: float = 0.3) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(timeout)
        try:
            s.connect((host, port))
            return True
        except Exception:
            return False

def wait_for_devtools(timeout: float) -> bool:
    end = time.time() + timeout
    while time.time() < end:
        if is_port_open("127.0.0.1", DEBUG_PORT):
            try:
                requests.get(f"http://127.0.0.1:{DEBUG_PORT}/json/version", timeout=0.6)
                return True
            except Exception:
                pass
        time.sleep(0.3)
    return False

def ask_yn(msg: str) -> bool:
    try:
        return input(f"{msg} [y/N]: ").strip().lower() in ("y", "yes")
    except EOFError:
        return False


# =========================
# Browser control (Edge/Chrome)
# =========================
def find_edge_exe() -> str:
    for p in [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]:
        if os.path.exists(p):
            return p
    return "msedge.exe"

def find_chrome_exe() -> str:
    for p in [
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]:
        if os.path.exists(p):
            return p
    return "chrome.exe"

def kill_edge():
    for cmd in (r'taskkill /IM msedge.exe /F',):
        subprocess.run(cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

def kill_chrome():
    for cmd in (r'taskkill /IM chrome.exe /F',):
        subprocess.run(cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

def launch_browser_with_devtools(browser: str, url: str):
    """browser in {'edge','chrome'}"""
    if browser == "edge":
        exe = find_edge_exe()
        prof = EDGE_PROFILE_DIR
    else:
        exe = find_chrome_exe()
        prof = CHROME_PROFILE_DIR

    Path(prof).mkdir(parents=True, exist_ok=True)
    args = [
        exe,
        f"--remote-debugging-port={DEBUG_PORT}",
        f"--remote-allow-origins={ALLOW_ORIGINS}",
        f"--user-data-dir={prof}",
        "--no-first-run",
        "--no-default-browser-check",
        url,
    ]
    subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

def ensure_browser_ready(browser: str):
    """Attach to existing DevTools socket if open; otherwise (optionally) relaunch browser."""
    if not is_port_open("127.0.0.1", DEBUG_PORT):
        print(f"Launching {browser.title()} with DevTools on port {DEBUG_PORT}…")
        if browser == "edge":
            kill_edge(); time.sleep(0.8)
        else:
            kill_chrome(); time.sleep(0.8)
        launch_browser_with_devtools(browser, MAP_URL)
        if not wait_for_devtools(DEVTOOLS_TIMEOUT):
            print("[!] DevTools port did not open. Check firewall or change DEBUG_PORT.")
            return False
    else:
        print(f"[info] Detected existing DevTools socket on {DEBUG_PORT}; will try to attach…")
    return True


# =========================
# DevTools (CDP) helpers
# =========================
def list_targets() -> list:
    r = requests.get(f"http://127.0.0.1:{DEBUG_PORT}/json", timeout=2.0)
    return r.json()

def find_shalazam_target() -> Optional[dict]:
    try:
        for t in list_targets():
            if t.get("type") == "page" and "shalazam.info/maps/1" in (t.get("url") or ""):
                return t
        for t in list_targets():
            if t.get("type") == "page":
                return t
    except Exception:
        pass
    return None

class CDPClient:
    def __init__(self, ws_url: str):
        # Send exactly one Origin header via origin= to satisfy modern Chromium
        self.ws = websocket.create_connection(
            ws_url,
            enable_multithread=True,
            origin=f"http://127.0.0.1:{DEBUG_PORT}",
        )
        self.msg_id = 0  # initialize request counter

    def send(self, method: str, params: dict | None = None):
        self.msg_id += 1
        payload = {"id": self.msg_id, "method": method}
        if params:
            payload["params"] = params
        self.ws.send(json.dumps(payload))
        while True:
            resp = json.loads(self.ws.recv())
            if resp.get("id") == self.msg_id:
                return resp

    def eval(self, expression: str):
        return self.send("Runtime.evaluate", {
            "expression": expression,
            "awaitPromise": True,
            "returnByValue": True
        })

    def navigate(self, url: str):
        return self.send("Page.navigate", {"url": url})

    def enable(self):
        self.send("Runtime.enable", {})
        self.send("Page.enable", {})

    def close(self):
        try:
            self.ws.close()
        except Exception:
            pass

def connect_to_shalazam_cdp(allow_relaunch=True, browser="edge") -> Optional[CDPClient]:
    # Try to attach to whichever page we find
    try:
        t = find_shalazam_target()
        if t:
            cdp = CDPClient(t["webSocketDebuggerUrl"])
            cdp.enable()
            if "shalazam.info/maps/1" not in (t.get("url") or ""):
                cdp.navigate(MAP_URL); time.sleep(1.2)
            return cdp
    except websocket._exceptions.WebSocketBadStatusException as e:
        # 403 means browser wasn’t launched with --remote-allow-origins
        if "403" in str(e) and allow_relaunch:
            print("[warn] CDP 403 Forbidden. Relaunching with --remote-allow-origins…")
            if ask_yn(f"Close {browser.title()} and relaunch with the correct flags now?"):
                if browser == "edge":
                    kill_edge()
                else:
                    kill_chrome()
                time.sleep(0.8)
                launch_browser_with_devtools(browser, MAP_URL)
                if not wait_for_devtools(DEVTOOLS_TIMEOUT):
                    print("[!] DevTools port did not open after relaunch.")
                    return None
                return connect_to_shalazam_cdp(allow_relaunch=False, browser=browser)
            return None
        else:
            print(f"[!] CDP connect failed: {e}")
            return None
    except Exception as e:
        print(f"[!] CDP attach error: {e}")

    # If no targets and we’re allowed, start the browser ourselves
    if allow_relaunch:
        print(f"[info] Launching {browser.title()} with DevTools…")
        if browser == "edge":
            kill_edge(); time.sleep(0.8)
        else:
            kill_chrome(); time.sleep(0.8)
        launch_browser_with_devtools(browser, MAP_URL)
        if not wait_for_devtools(DEVTOOLS_TIMEOUT):
            print("[!] DevTools port did not open.")
            return None
        return connect_to_shalazam_cdp(allow_relaunch=False, browser=browser)

    return None


# =========================
# Pantheon helpers (robust Win32 focus)
# =========================
def focus_pantheon() -> bool:
    target = None
    for needle in PANTHEON_WINDOW_TITLES:
        for t in gw.getAllTitles():
            if needle.lower() in t.lower():
                matches = gw.getWindowsWithTitle(t)
                if matches:
                    target = matches[0]
                    break
        if target:
            break
    if not target:
        return False

    try:
        hwnd = int(target._hWnd)
    except Exception:
        return False

    user32 = ctypes.windll.user32
    IsIconic = user32.IsIconic
    ShowWindow = user32.ShowWindow
    GetForegroundWindow = user32.GetForegroundWindow
    GetWindowThreadProcessId = user32.GetWindowThreadProcessId
    AttachThreadInput = user32.AttachThreadInput
    SetForegroundWindow = user32.SetForegroundWindow
    BringWindowToTop = user32.BringWindowToTop
    SetFocus = user32.SetFocus
    SetWindowPos = user32.SetWindowPos

    SW_RESTORE = 9
    HWND_TOPMOST = -1
    HWND_NOTOPMOST = -2
    SWP_NOMOVE = 0x0002
    SWP_NOSIZE = 0x0001
    SWP_SHOWWINDOW = 0x0040

    try:
        if IsIconic(hwnd):
            ShowWindow(hwnd, SW_RESTORE)
            time.sleep(0.05)

        fg = GetForegroundWindow()
        tid_fg = GetWindowThreadProcessId(fg, None)
        tid_hwnd = GetWindowThreadProcessId(hwnd, None)
        AttachThreadInput(tid_fg, tid_hwnd, True)
        try:
            BringWindowToTop(hwnd)
            SetForegroundWindow(hwnd)
            SetFocus(hwnd)
        finally:
            AttachThreadInput(tid_fg, tid_hwnd, False)

        SetWindowPos(hwnd, HWND_TOPMOST,   0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)
        SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW)

        if GetForegroundWindow() != hwnd:
            try:
                x = target.left + 40
                y = target.top + 10
                pyautogui.moveTo(x, y, duration=0.05)
                pyautogui.click()
            except Exception:
                pass

        time.sleep(0.05)
        return True
    except Exception:
        try:
            if target.isMinimized:
                target.restore()
            x = target.left + 40
            y = target.top + 10
            pyautogui.moveTo(x, y, duration=0.05)
            pyautogui.click()
            time.sleep(0.05)
            return True
        except Exception:
            return False

def send_loc_and_copy():
    pyautogui.press("enter")
    time.sleep(CHAT_WAKE_DELAY)
    pyautogui.typewrite("/loc", interval=TYPING_DELAY)
    pyautogui.press("enter")
    time.sleep(AFTER_LOC_WAIT)

def get_clipboard_text() -> str:
    try:
        return pyperclip.paste()
    except Exception:
        return ""

def parse_jumploc(raw: str) -> Tuple[Optional[str], Optional[str]]:
    m = re.match(
        r"^\s*\/?jumploc\s+(-?\d+(\.\d+)?)\s+(-?\d+(\.\d+)?)\s+(-?\d+(\.\d+)?)\s+(-?\d+(\.\d+)?)",
        (raw or "").strip(), flags=re.IGNORECASE
    )
    if not m:
        return None, None
    return m.group(1), m.group(5)  # X (1st), Y (3rd)


# =========================
# JS to drop the pin on Shalazam
# =========================
JS_DROP_PIN = r"""
(() => {
  const qs = s => document.querySelector(s);
  let xInput = qs("input[placeholder='X']") || qs("input[name='x']") || qs("#x");
  let yInput = qs("input[placeholder='Y']") || qs("input[name='y']") || qs("#y");
  if (!xInput || !yInput) return { ok:false, reason:"inputs-not-found" };
  const setVal = (el, v) => {
    const d = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value');
    d.set.call(el, v);
    el.dispatchEvent(new Event('input', {bubbles:true}));
    el.dispatchEvent(new Event('change', {bubbles:true}));
  };
  setVal(xInput, "%(X)s");
  setVal(yInput, "%(Y)s");
  let dropBtn = Array.from(document.querySelectorAll("button,input[type='button']"))
    .find(b => (b.innerText || b.value || "").toLowerCase().includes("drop"));
  if (!dropBtn) {
    const cand = (xInput.closest("form,div,section") || document).querySelectorAll("button");
    dropBtn = Array.from(cand).find(b => (b.innerText || "").toLowerCase().includes("drop"));
  }
  if (!dropBtn) return { ok:false, reason:"drop-button-not-found" };
  dropBtn.click();
  return { ok:true };
})();
"""

def cdp_drop_pin(cdp: "CDPClient", x: str, y: str) -> bool:
    code = JS_DROP_PIN.replace("%(X)s", x).replace("%(Y)s", y)
    try:
        res = cdp.eval(code)
        val = ((res or {}).get("result") or {}).get("result", {}).get("value")
        return bool(isinstance(val, dict) and val.get("ok"))
    except Exception as e:
        print(f"[!] JS eval error: {e}")
        return False


# =========================
# Main + hotkeys
# =========================
def main():
    # Choose browser
    choice = input("Use (E)dge or (C)hrome? [E/C]: ").strip().lower()
    browser = "edge" if choice in ("", "e", "edge") else "chrome"
    print(f"[info] Using {browser.title()}")

    # Ensure DevTools is up (launch if needed)
    if not ensure_browser_ready(browser):
        return

    # Attach to Shalazam tab
    cdp = connect_to_shalazam_cdp(allow_relaunch=True, browser=browser)
    if not cdp:
        print("[!] Could not attach to a Shalazam tab.")
        return
    print("[OK] Attached to Shalazam via CDP.")
    print(f"Hotkeys:\n  {HOTKEY_TRIGGER} → grab /loc and drop pin\n  {HOTKEY_QUIT}   → quit")

    last_fire = 0.0

    def on_trigger():
        nonlocal last_fire, cdp
        if time.time() - last_fire < HOTKEY_DEBOUNCE:
            return
        last_fire = time.time()

        print("\n[*] Capturing /loc…")
        if not focus_pantheon():
            print("[!] Pantheon window not found or couldn’t be focused.")
            return
        send_loc_and_copy()
        raw = get_clipboard_text()
        if not raw:
            print("[!] Clipboard empty after /loc.")
            return
        x, y = parse_jumploc(raw)
        if not (x and y):
            print(f"[!] Parse failed: {raw!r}")
            return
        print(f"[INFO] Parsed X={x} Y={y}. Dropping pin…")

        try:
            ok = cdp_drop_pin(cdp, x, y)
        except Exception:
            try:
                cdp.close()
            except Exception:
                pass
            # Reattach once without relaunch
            cdp = connect_to_shalazam_cdp(allow_relaunch=False, browser=browser)
            ok = cdp_drop_pin(cdp, x, y) if cdp else False

        print("[OK] Pin dropped." if ok else "[!] Failed to drop pin.")

    keyboard.add_hotkey(HOTKEY_TRIGGER, on_trigger)
    print("\nReady! (Run as Administrator for reliable global hotkeys.)")
    print(f"Press {HOTKEY_QUIT} to exit.")
    keyboard.wait(HOTKEY_QUIT)

    try:
        cdp.close()
    except Exception:
        pass
    print("\nExiting… bye!")


if __name__ == "__main__":
    main()
