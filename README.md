# Pantheon‑Loc‑Hotkey

Lulu wants a map fragment!  
Automate dropping your `/loc` in Pantheon onto Shalazam maps via **Edge or Chrome**—with a hotkey.

---

##  Features

- Choose either **Microsoft Edge** or **Google Chrome** at launch.
- Sends `/loc` in Pantheon → translates `/jumploc X Z Y H` → instantly drops a pin on [Shalazam Map #1](https://shalazam.info/maps/1).
- Uses **Chrome DevTools Protocol (CDP)** to talk directly to your browser tab.
- Robust focus to Pantheon—even when minimized or hidden.
- Global hotkeys:
  - **Ctrl + L**: Capture location & drop them on the map.
  - **Ctrl + Q**: Quit cleanly.
- No internet API required after install—everything runs locally.

---

##  Repo Contents

```
├── dist/
│   ├── pantheon_loc_hotkey_chrome_or_edge.exe
│   └── pantheon_loc_hotkey_edge.exe
├── pantheon_loc_hotkey_chrome_or_edge.py
├── pantheon_loc_hotkey_edge.py
├── pantheon-loc-hotkey.spec
├── pantheon-loc-hotkey-chrome-or-edge.spec
├── requirements.txt
└── LICENSE (GPL‑3.0)
```

- **Built executables** are already in `dist/`. Use them directly—no Python install needed.
- `.py` scripts for full transparency or customization.
- `.spec` files for building new versions with PyInstaller.

---

##  Requirements (if using `.py`)

- Windows 10 or later
- Python 3.10+
- Dependencies (install via PowerShell/Terminal):
  ```powershell
  py -m pip install -r requirements.txt
  ```
- **Run as Administrator** (required for global hotkey support via `keyboard`).

---

##  Using the EXE (Quick Start)

1. Double-click  
   - `pantheon_loc_hotkey_chrome_or_edge.exe` — asks which browser to use  
   - or `pantheon_loc_hotkey_edge.exe` — defaults to Microsoft Edge
2. Choose your browser (if prompted).
3. Script will ensure the browser is launched with:
   - `--remote-debugging-port=9222`
   - `--remote-allow-origins=*`
4. A console window will appear; you’ll see "Attached via CDP".
5. In-game:
   - Press **Ctrl + L** to send `/loc`, pin your location.
   - Press **Ctrl + Q** to exit.

---

##  Building Your Own EXE

You can rebuild or customize using PyInstaller:

```powershell
py -m pip install pyinstaller
py -m PyInstaller --onefile --uac-admin --console pantheon_loc_hotkey_chrome_or_edge.py
```

The `.spec` files included can also be used if you need special assets or icons.

---

##  Troubleshooting

| Issue                          | Fix                                                                 |
|-------------------------------|----------------------------------------------------------------------|
| Clipboard empty after `/loc`  | Increase `AFTER_LOC_WAIT` delay in script (1.10 s)                   |
| CDP 403 Forbidden             | Browser must launch with `--remote-allow-origins=*`                 |
| Hotkeys not working           | Run the script or exe **as Administrator**                          |

---

##  License

Licensed under **GPL‑3.0** — see [LICENSE](LICENSE) for details.
