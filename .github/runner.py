import requests
import subprocess
import tempfile
import time
import sys
import os
import glob
import threading
import signal
import re

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

        if action == "opened_apps":
            wins = [w.Name for w in root.GetChildren() if w.Name]
            send_message(chat_id, "Opened Apps:\n" + "\n".join(wins))

        elif action == "available_apps":
            apps = {}
            try:
                for path in [os.path.join(os.environ['USERPROFILE'], 'Desktop'), r'C:\Users\Public\Desktop']:
                    if os.path.exists(path):
                        for f in os.listdir(path):
                            if f.endswith('.lnk'): apps[f.replace('.lnk', '')] = os.path.join(path, f)
                for path in [os.path.join(os.environ['AppData'], r'Microsoft\Windows\Start Menu\Programs'),
                             r'C:\ProgramData\Microsoft\Windows\Start Menu\Programs']:
                    if os.path.exists(path):
                        for root_dir, dirs, files in os.walk(path):
                            for f in files:
                                if f.endswith('.lnk'): apps[f.replace('.lnk', '')] = os.path.join(root_dir, f)
            except: pass

            if params == "list":
                res = sorted(list(apps.keys()))
                send_message(chat_id, "Available Apps:\n" + ("\n".join(res[:100]) or "None found"))
            else:
                target = params.lower()
                match = next((path for name, path in apps.items() if name.lower() == target), None)
                if not match:
                    common = {"paint": "mspaint", "notepad": "notepad", "calc": "calc", "cmd": "cmd", "explorer": "explorer"}
                    match = common.get(target)
                if match:
                    os.startfile(match) if os.path.exists(match) else subprocess.Popen(match)
                    send_message(chat_id, f"Opening `{params}`...")
                else:
                    send_message(chat_id, f"Could not find app `{params}`.")

        elif action == "list_buttons":
            curr_win = auto.GetForegroundControl()
            while curr_win and curr_win.ControlTypeName != "WindowControl":
                curr_win = curr_win.GetParentControl()
            if not curr_win: send_message(chat_id, "No active window found."); return
            controls = []
            def find_controls(ctrl):
                interactive_types = ["ButtonControl", "MenuItemControl", "ListItemControl", "TreeItemControl", "TabItemControl", "HyperlinkControl", "SplitButtonControl", "CheckBoxControl", "RadioButtonControl"]
                if ctrl.ControlTypeName in interactive_types:
                    if ctrl.Name: controls.append(f"{ctrl.ControlTypeName[:-7]}: {ctrl.Name}")
                if len(controls) > 100: return
                for child in ctrl.GetChildren(): find_controls(child)
            find_controls(curr_win)
            res = "\n".join(sorted(list(set(controls)))[:60])
            send_message(chat_id, f"Controls in `{curr_win.Name}`:\n" + (res or "No controls found"))

        elif action == "click" or action == "double_click":
            def perform_action(ctrl, target):
                if target.lower() in (ctrl.Name or "").lower():
                    if action == "click": ctrl.Click()
                    else: ctrl.DoubleClick()
                    return True
                for child in ctrl.GetChildren():
                    if perform_action(child, target): return True
                return False
            if perform_action(root, params): send_message(chat_id, f"{'Clicked' if action == 'click' else 'Double-clicked'} `{params}`")
            else: send_message(chat_id, f"Could not find `{params}`")

        elif action == "press":
            pyautogui.hotkey(*params.split("+"))
            send_message(chat_id, f"Pressed `{params}`")

        elif action == "type":
            pyautogui.write(params)
            send_message(chat_id, f"Typed `{params}`")

    except Exception as e: send_message(chat_id, f"UI Error: {e}")

def livestream(chat_id):
    def run_server():
        try:
            from flask import Flask, Response
            from PIL import ImageGrab
            import numpy as np
            import cv2

            app = Flask(__name__)
            last_frame = None

            def gen():
                nonlocal last_frame
                while True:
                    img = ImageGrab.grab()
                    frame = cv2.cvtColor(np.array(img), cv2.COLOR_RGB2BGR)
                    if last_frame is not None:
                        diff = cv2.absdiff(frame, last_frame)
                        if np.mean(diff) < 0.2:
                            time.sleep(0.2); continue
                    last_frame = frame.copy()
                    _, buffer = cv2.imencode(".jpg", frame, [int(cv2.IMWRITE_JPEG_QUALITY), 30])
                    yield (b'--frame\r\nContent-Type: image/jpeg\r\n\r\n' + buffer.tobytes() + b'\r\n')
                    time.sleep(0.1)

            @app.route("/")
            def index():
                return """
                <html>
                  <head><title>GitHub VM Live</title></head>
                  <body style="margin:0; background: #000; display:flex; align-items:center; justify-content:center;">
                    <img src="/stream" style="max-width:100%; max-height:100vh;">
                  </body>
                </html>
                """

            @app.route("/stream")
            def stream(): return Response(gen(), mimetype='multipart/x-mixed-replace; boundary=frame')

            send_message(chat_id, "Starting Cloudflare tunnel...")
            # Windows: Run cloudflared.exe from its installation path
            cf_cmd = ["cloudflared", "tunnel", "--url", "http://127.0.0.1:5000"]
            cf_proc = subprocess.Popen(cf_cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, bufsize=1)

            url = None
            start_time = time.time()
            while time.time() - start_time < 90:
                line = cf_proc.stdout.readline()
                if not line: break
                print(f"[cf] {line.strip()}", flush=True)
                match = re.search(r'https://[a-zA-Z0-9-]+\.trycloudflare\.com', line)
                if match:
                    url = match.group(0); break

            if url: send_message(chat_id, f"LiveStream online: {url}")
            else: send_message(chat_id, "Failed to capture Cloudflare URL.")

            app.run(host="0.0.0.0", port=5000, threaded=True) # Bind to all interfaces
        except Exception as e: send_message(chat_id, f"LiveStream Error: {e}")

    threading.Thread(target=run_server, daemon=True).start()

WELCOME = "*GitHub VM Bot*\n- `screen`: Screenshot\n- `livestream`: Live screen view\n- `terminate`: Kill task\n- `apps`: Available apps\n- `opened apps`: Running apps\n- `buttons`: List controls\n- `click <name>`: Click\n- `open <app>`: Launch app"

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

            orig_text = (msg.get("text") or "").strip()
            text = orig_text.lower()
            if not text: continue

            if text in ("/start", "/help"): send_message(chat_id, WELCOME); continue
            if text == "/stop": sys.exit(0)
            if text == "screen": take_screenshot(chat_id); continue
            if text == "livestream": livestream(chat_id); continue
            if text == "terminate":
                if current_process: current_process.terminate(); send_message(chat_id, "Terminated.")
                else: send_message(chat_id, "No task.")
                continue

            if text == "opened apps": ui_automation(chat_id, "opened_apps"); continue
            if text == "apps": ui_automation(chat_id, "available_apps", "list"); continue
            if text.startswith("open "): ui_automation(chat_id, "available_apps", orig_text[5:].strip()); continue
            if text == "buttons": ui_automation(chat_id, "list_buttons"); continue
            if text.startswith("click "): ui_automation(chat_id, "click", orig_text[6:].strip()); continue
            if text.startswith("double click "): ui_automation(chat_id, "double_click", orig_text[13:].strip()); continue
            if text.startswith("press "): ui_automation(chat_id, "press", orig_text[6:].strip()); continue
            if text.startswith("type "): ui_automation(chat_id, "type", orig_text[5:].strip()); continue

            is_shell = text.startswith("/") or any(text.startswith(x) for x in ["pip ", "npm ", "git ", "python "])
            run_command(chat_id, orig_text[1:].strip() if text.startswith("/") else orig_text, is_python=not is_shell)
    except Exception as e: print(f"Error: {e}"); time.sleep(5)
