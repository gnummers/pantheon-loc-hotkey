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


# =========================
# Config
# =========================
MAP_URL = "https://shalazam.info/maps/1"
DEBUG_PORT = 9222
EDGE_PROFILE_DIR = r"C:\msedgeprofile_panth"

# IMPORTANT: allow our Origin on the CDP socket to prevent 403
ALLOW_ORIGINS = "*"

PANTHEON_WINDOW_TITLES = ["Pantheon", "Pantheon: Rise of the Fallen"]

# Timings
TYPING_DELAY = 0.03
CHAT_WAKE_DELAY = 0.20
AFTER_LOC_WAIT = 1.10
EDGE_START_TIMEOUT = 45.0

# Hotkeys
HOTKEY_TRIGGER = "ctrl+l"
HOTKEY_QUIT = "ctrl+q"
HOTKEY_DEBOUNCE = 0.35


# =========================
# Utils
# =========================
def is_port_open(host: str, port: int, timeout: float = 0.3) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(timeout)
        try:
            s.connect((host, port))
            return True
        except Exception:
            return False

def kill_edge():
    for cmd in (r'taskkill /IM msedge.exe /F', r'taskkill /IM msedgedriver.exe /F'):
        subprocess.run(cmd, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

def ask_yn(msg: str) -> bool:
    try:
        return input(f"{msg} [y/N]: ").strip().lower() in ("y", "yes")
    except EOFError:
        return False


# =========================
# Edge DevTools (CDP)
# =========================
def launch_edge_with_devtools(url: str) -> None:
    Path(EDGE_PROFILE_DIR).mkdir(parents=True, exist_ok=True)
    edge_candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        "msedge.exe",
    ]
    edge = next((p for p in edge_candidates if os.path.exists(p) or p == "msedge.exe"), "msedge.exe")
    args = [
        edge,
        f"--remote-debugging-port={DEBUG_PORT}",
        f"--remote-allow-origins={ALLOW_ORIGINS}",
        f"--user-data-dir={EDGE_PROFILE_DIR}",
        "--no-first-run",
        "--no-default-browser-check",
        MAP_URL if url is None else url,
    ]
    subprocess.Popen(args, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

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
        # Ensure we send the allowed Origin; Edge enforces it when that flag is set
        self.ws = websocket.create_connection(
            ws_url, enable_multithread=True, header=[f"Origin: http://127.0.0.1:{DEBUG_PORT}"]
        )
        self.msg_id = 0

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

def connect_to_shalazam_cdp(allow_relaunch=True) -> Optional[CDPClient]:
    # Try attach to existing first
    try:
        t = find_shalazam_target()
        if t:
            cdp = CDPClient(t["webSocketDebuggerUrl"])
            cdp.enable()
            if "shalazam.info/maps/1" not in (t.get("url") or ""):
                cdp.navigate(MAP_URL); time.sleep(1.2)
            return cdp
    except websocket._exceptions.WebSocketBadStatusException as e:
        # 403 = missing allow-origins; relaunch Edge properly
        if "403" in str(e) and allow_relaunch:
            print("[warn] CDP 403 Forbidden. Edge must be started with --remote-allow-origins.")
            if ask_yn("Close Edge and relaunch with the correct flags now?"):
                kill_edge(); time.sleep(0.8)
                launch_edge_with_devtools(MAP_URL)
                if not wait_for_devtools(EDGE_START_TIMEOUT):
                    print("[!] DevTools port did not open after relaunch.")
                    return None
                return connect_to_shalazam_cdp(allow_relaunch=False)
            return None
        else:
            print(f"[!] CDP connect failed: {e}")
            return None
    except Exception as e:
        print(f"[!] CDP attach error: {e}")

    # If no target and allowed to relaunch, start Edge ourselves
    if allow_relaunch:
        print("[info] Launching Edge with DevTools + allow-origins…")
        kill_edge(); time.sleep(0.8)
        launch_edge_with_devtools(MAP_URL)
        if not wait_for_devtools(EDGE_START_TIMEOUT):
            print("[!] DevTools port did not open.")
            return None
        return connect_to_shalazam_cdp(allow_relaunch=False)
    return None


# =========================
# Pantheon helpers
# =========================
def focus_pantheon() -> bool:
    for needle in PANTHEON_WINDOW_TITLES:
        for t in gw.getAllTitles():
            if needle.lower() in t.lower():
                win = gw.getWindowsWithTitle(t)[0]
                if win.isMinimized:
                    win.restore()
                win.activate()
                return True
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
# JS injection to drop the pin
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
    if not is_port_open("127.0.0.1", DEBUG_PORT):
        print(f"Launching Edge with DevTools on port {DEBUG_PORT}…")
        launch_edge_with_devtools(MAP_URL)
        if not wait_for_devtools(EDGE_START_TIMEOUT):
            print("[!] DevTools port did not open. Check firewall or change DEBUG_PORT.")
            return
    else:
        print(f"[info] Edge already listening on {DEBUG_PORT}. Attempting to attach…")

    cdp = connect_to_shalazam_cdp(allow_relaunch=True)
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
            print("[!] Pantheon window not found.")
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
            # Session might have died; reattach once
            try:
                cdp.close()
            except Exception:
                pass
            cdp = connect_to_shalazam_cdp(allow_relaunch=False)
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
