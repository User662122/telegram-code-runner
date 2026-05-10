import requests
import subprocess
import tempfile
import time
import sys
import os
import glob
import threading
import signal

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
ALLOWED_CHAT_ID = os.environ.get("ALLOWED_CHAT_ID")
if ALLOWED_CHAT_ID:
    ALLOWED_CHAT_ID = int(ALLOWED_CHAT_ID)

BASE  = f"https://api.telegram.org/bot{TOKEN}"
TICK  = chr(96) * 3

current_process = None

def send_message(chat_id, text, parse_mode='Markdown'):
    try:
        if len(text) > 4000: text = text[:3900] + "\n...(truncated)"
        return requests.post(f"{BASE}/sendMessage", json={"chat_id": chat_id, "text": text, "parse_mode": parse_mode}, timeout=10).json()
    except Exception as e: print(f"[send error] {e}", flush=True)

def send_document(chat_id, file_path):
    try:
        if not os.path.exists(file_path):
            send_message(chat_id, f"Error: `{file_path}` not found.")
            return
        with open(file_path, 'rb') as f:
            return requests.post(f"{BASE}/sendDocument", data={"chat_id": chat_id}, files={"document": f}, timeout=30).json()
    except Exception as e: print(f"[send_doc error] {e}", flush=True)

def send_photo(chat_id, file_path):
    try:
        if not os.path.exists(file_path): return
        with open(file_path, 'rb') as f:
            requests.post(f"{BASE}/sendPhoto", data={"chat_id": chat_id}, files={"photo": f}, timeout=30)
    except Exception as e: print(f"[send_photo error] {e}", flush=True)

def download_file(chat_id, file_id, file_name):
    try:
        file_name = os.path.basename(file_name)
        file_info = requests.get(f"{BASE}/getFile", params={"file_id": file_id}).json()
        if not file_info.get("ok"): return
        url = f"https://api.telegram.org/file/bot{TOKEN}/{file_info['result']['file_path']}"
        with open(file_name, 'wb') as f: f.write(requests.get(url).content)
        send_message(chat_id, f"Saved `{file_name}`")
    except Exception as e: send_message(chat_id, f"Error: {e}")

def run_command(chat_id, cmd, is_python=False):
    global current_process
    def target():
        global current_process
        try:
            if is_python:
                with tempfile.NamedTemporaryFile(mode="w", suffix=".py", delete=False, encoding="utf-8") as f:
                    f.write(cmd); fname = f.name
                p_cmd = [sys.executable, fname]
            else:
                p_cmd = cmd; fname = None
            current_process = subprocess.Popen(p_cmd, shell=not is_python, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, cwd=os.getcwd())
            out, _ = current_process.communicate()
            current_process = None
            if fname:
                try: os.unlink(fname)
                except: pass
            res = out.strip() or "(No output)"
            if len(res) > 3800:
                with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
                    f.write(res); t_name = f.name
                send_document(chat_id, t_name); os.unlink(t_name)
            else: send_message(chat_id, f"{TICK}\n{res}\n{TICK}")
        except Exception as e: current_process = None; send_message(chat_id, f"Error: {e}")
    threading.Thread(target=target).start()

def take_screenshot(chat_id):
    filename = "screenshot.png"
    try:
        if os.name == 'nt':
            ps_cmd = "powershell -command \"[Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')|Out-Null;[Reflection.Assembly]::LoadWithPartialName('System.Drawing')|Out-Null;$S=[System.Windows.Forms.Screen]::PrimaryScreen;$B=New-Object System.Drawing.Bitmap($S.Bounds.Width,$S.Bounds.Height);$G=[System.Drawing.Graphics]::FromImage($B);$G.CopyFromScreen(0,0,0,0,$B.Size);$B.Save('screenshot.png',[System.Drawing.Imaging.ImageFormat]::Png);$G.Dispose();$B.Dispose();\""
            subprocess.run(ps_cmd, shell=True, check=True)
        else:
            from PIL import ImageGrab
            ImageGrab.grab().save(filename)
        send_photo(chat_id, filename)
        if os.path.exists(filename): os.remove(filename)
    except Exception as e: send_message(chat_id, f"Error: {e}")

def ui_automation(chat_id, action, params=None):
    if os.name != 'nt':
        send_message(chat_id, "UI Automation only supported on Windows."); return
    try:
        import uiautomation as auto
        import pyautogui
        root = auto.GetRootControl()
        if action == "list_windows":
            wins = [w.Name for w in root.GetChildren() if w.Name]
            send_message(chat_id, "Opened Apps:\n" + "\n".join(wins))
        elif action == "list_buttons":
            # Logic to find buttons in current window
            curr_win = auto.GetForegroundControl()
            while curr_win and curr_win.ControlTypeName != "WindowControl":
                curr_win = curr_win.GetParentControl()
            if not curr_win: send_message(chat_id, "No active window found."); return
            btns = [b.Name for b in curr_win.GetChildren() if b.ControlTypeName == "ButtonControl" and b.Name]
            send_message(chat_id, f"Buttons in `{curr_win.Name}`:\n" + "\n".join(btns))
        elif action == "click":
            # Recursive search for button name
            def find_and_click(ctrl, target):
                if target.lower() in (ctrl.Name or "").lower():
                    ctrl.Click(); return True
                for child in ctrl.GetChildren():
                    if find_and_click(child, target): return True
                return False
            if find_and_click(root, params): send_message(chat_id, f"Clicked `{params}`")
            else: send_message(chat_id, f"Could not find `{params}`")
        elif action == "press":
            keys = params.split("+")
            pyautogui.hotkey(*keys)
            send_message(chat_id, f"Pressed `{params}`")
        elif action == "type":
            pyautogui.write(params)
            send_message(chat_id, f"Typed `{params}`")
    except Exception as e: send_message(chat_id, f"UI Error: {e}")

WELCOME = "*GitHub VM Bot*\n- `screen`: Screenshot\n- `terminate`: Kill task\n- `apps`: List opened apps\n- `buttons`: List buttons in active window\n- `click <name>`: Click something\n- `press <keys>`: Hotkeys (e.g., press Ctrl+A)\n- `type <text>`: Type text"

if not TOKEN: sys.exit(1)
try:
    r = requests.get(f"{BASE}/getUpdates", params={"offset": -1}, timeout=10).json()
    offset = r["result"][0]["update_id"] + 1 if r.get("result") else 0
except: offset = 0

while True:
    try:
        updates = requests.get(f"{BASE}/getUpdates", params={"offset": offset, "timeout": 30}, timeout=40).json().get("result", [])
        for upd in updates:
            offset = upd["update_id"] + 1
            msg = upd.get("message", {})
            chat_id = msg.get("chat", {}).get("id")
            if not chat_id or (ALLOWED_CHAT_ID and chat_id != ALLOWED_CHAT_ID): continue

            if "document" in msg: download_file(chat_id, msg["document"]["file_id"], msg["document"].get("file_name", "file")); continue
            elif "photo" in msg: download_file(chat_id, msg["photo"][-1]["file_id"], f"photo_{int(time.time())}.jpg"); continue

            text = (msg.get("text") or "").strip()
            if not text: continue
            if text in ("/start", "/help"): send_message(chat_id, WELCOME); continue
            if text == "/stop": sys.exit(0)
            if text.lower() == "screen": take_screenshot(chat_id); continue
            if text.lower() == "terminate":
                if current_process: current_process.terminate(); send_message(chat_id, "Terminated.")
                else: send_message(chat_id, "No task.")
                continue
            if text.lower() == "apps": ui_automation(chat_id, "list_windows"); continue
            if text.lower() == "buttons": ui_automation(chat_id, "list_buttons"); continue
            if text.lower().startswith("click "): ui_automation(chat_id, "click", text[6:].strip()); continue
            if text.lower().startswith("press "): ui_automation(chat_id, "press", text[6:].strip()); continue
            if text.lower().startswith("type "): ui_automation(chat_id, "type", text[5:].strip()); continue

            is_shell = text.startswith("/") or any(text.startswith(x) for x in ["pip ", "npm ", "git ", "python "])
            run_command(chat_id, text[1:].strip() if text.startswith("/") else text, is_python=not is_shell)
    except Exception as e: print(f"Error: {e}"); time.sleep(5)
