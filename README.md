# telegram-code-runner

Telegram-controlled GitHub Actions Windows VM — now with **full remote desktop** like TeamViewer, directly from Telegram + browser.

## Quick Start

1. Add secret `TELEGRAM_BOT_TOKEN` in repo Settings → Secrets
2. Run workflow **Telegram Runner** via Actions → Run workflow (optionally set Allowed Chat ID)
3. In Telegram, send `/start` to your bot → get full help

## Features

### 🖥️ Full VM Control (NEW — write anywhere, like local)
- **Type anywhere:** `type Hello World` — types into active window (Notepad, browser, etc).  
  Aliases: `paste <text>`, `typepaste <text>` for long/unicode via clipboard (reliable).
- **Keys & hotkeys:** `press enter`, `press ctrl+c`, `press alt+f4`, `press win+r`, `hotkey ctrl+shift+t`
- **Mouse:** `click 500 300` (coords), `click Save` (UI control name), `rclick 100 200`, `doubleclick 400 500`, `move 800 500`, `drag 100 100 500 500`, `scroll up 500`, `pos`
- **Live Remote Desktop:** `livestream` / `live` → get `https://*.trycloudflare.com` link.  
  Open on phone/PC: **tap to click, drag to select, scroll, type** in the text box, keyboard shortkeys (Ctrl+C/V, Alt+F4, Win+R…). Works like VNC with ~5-8 FPS, auto-reconnects.

### 📁 File Manager (Telegram)
- `pwd`, `ls [path]`, `cat <file>`, `get <file>` (sends file), `mkdir`, `rm`, `mv <src> <dst>`, `cp <src> <dst>`, `write <file> <text>`, `append <file> <text>`, `find <pattern>`, `tree`, `du`, `zip`, `unzip`, `wget <url>`, `cd <path>`
- **Upload:** just send any file/photo to Telegram → saved to VM (`/home/runner/...`)

### 🪟 Window & Apps
- `apps` (list start-menu & desktop), `windows` / `opened apps`, `open notepad` / `open chrome`, `browser https://...`, `focus <window>`, `close <window>` (Alt+F4), `minimize`/`maximize`, `buttons` (list controls in active window), `click <control name>`

### ⚙️ System
- `ps` (top processes), `kill <pid|name>`, `sysinfo` (CPU/RAM/Disk), `env [var]`, `uptime`, `ip`
- Clipboard: `clip get`, `clip set <text>`, `copyclip <text>`
- Interactive stdin: `input <text>` sends to running process (for prompts)
- `input`, `hold <key>` / `release <key>`

### 💻 Shell & Code
- `/<command>` → shell (e.g. `/dir`, `/pip list`), `cmd <command>` / `exec <command>` explicit shell
- Any other text → Python (runs as script)
- `terminate` kills running task

## Examples

```
screen
livestream
type Hello from my phone!
press win+r
type notepad
press enter
type This was typed remotely!
press ctrl+s
click 500 400
pos
ls
get myfile.zip
write notes.txt Hello\nWorld
browser https://google.com
sysinfo
```

## How it works

- GitHub Actions `windows-latest` runner checks out code, installs Python deps (`mss`, `opencv`, `flask`, `pyautogui`, `uiautomation`, `psutil`, `pyperclip`), installs `cloudflared`, then runs `.github/runner.py`.
- Runner polls Telegram `getUpdates`, executes commands, replies, serves Flask on `0.0.0.0:5000` for screen streaming, exposes via `cloudflared` tunnel.
- Mouse/keyboard injected via `pyautogui` + `uiautomation`; typing uses clipboard fallback for reliability on Unicode/long text.

## Workflow

`.github/workflows/telegram-runner.yml` installs all deps, starts `runner.py`. Timeout 360 min (6h) per run; re-run workflow to extend.

---

> Tip: Use `livestream` for the most **local-like** experience. Open the link, tap where you want to click, type in the toolbar box → instantly appears in VM. Combine with `type` / `press` via Telegram for quick actions without opening the browser.
