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
# Optional: restrict to a specific chat ID for security
ALLOWED_CHAT_ID = os.environ.get("ALLOWED_CHAT_ID")
if ALLOWED_CHAT_ID:
    ALLOWED_CHAT_ID = int(ALLOWED_CHAT_ID)

BASE  = f"https://api.telegram.org/bot{TOKEN}"
TICK  = chr(96) * 3

# Global variable to track the current running process
current_process = None

def send_message(chat_id, text, parse_mode='Markdown'):
    try:
        if len(text) > 4000:
            text = text[:3900] + "\n...(truncated)"

        r = requests.post(
            f"{BASE}/sendMessage",
            json={"chat_id": chat_id, "text": text, "parse_mode": parse_mode},
            timeout=10
        )
        return r.json()
    except Exception as e:
        print(f"[send error] {e}", flush=True)

def send_document(chat_id, file_path):
    try:
        if not os.path.exists(file_path):
            send_message(chat_id, f"Error: File `{file_path}` not found.")
            return

        with open(file_path, 'rb') as f:
            r = requests.post(
                f"{BASE}/sendDocument",
                data={"chat_id": chat_id},
                files={"document": f},
                timeout=30
            )
        return r.json()
    except Exception as e:
        print(f"[send_document error] {e}", flush=True)
        send_message(chat_id, f"Error sending document: {e}")

def send_photo(chat_id, file_path):
    try:
        if not os.path.exists(file_path):
            return
        with open(file_path, 'rb') as f:
            requests.post(f"{BASE}/sendPhoto", data={"chat_id": chat_id}, files={"photo": f}, timeout=30)
    except Exception as e:
        print(f"[send_photo error] {e}", flush=True)

def download_file(chat_id, file_id, file_name):
    try:
        file_name = os.path.basename(file_name)
        r = requests.get(f"{BASE}/getFile", params={"file_id": file_id})
        file_info = r.json()
        if not file_info.get("ok"):
            send_message(chat_id, f"Failed to get file info: {file_info.get('description')}")
            return

        file_path = file_info["result"]["file_path"]
        download_url = f"https://api.telegram.org/file/bot{TOKEN}/{file_path}"

        r = requests.get(download_url)
        with open(file_name, 'wb') as f:
            f.write(r.content)

        send_message(chat_id, f"Saved `{file_name}`")
    except Exception as e:
        print(f"[download error] {e}", flush=True)
        send_message(chat_id, f"Error downloading file: {e}")

def run_command(chat_id, cmd, is_python=False):
    global current_process

    def target():
        global current_process
        try:
            if is_python:
                with tempfile.NamedTemporaryFile(mode="w", suffix=".py", delete=False, encoding="utf-8") as f:
                    f.write(cmd)
                    fname = f.name
                process_cmd = [sys.executable, fname]
            else:
                process_cmd = cmd
                fname = None

            current_process = subprocess.Popen(
                process_cmd, shell=not is_python,
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, cwd=os.getcwd()
            )

            output, _ = current_process.communicate()
            current_process = None

            if fname:
                try: os.unlink(fname)
                except: pass

            out = output.strip() or "(No output)"
            if len(out) > 3800:
                with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
                    f.write(out)
                    temp_name = f.name
                send_document(chat_id, temp_name)
                os.unlink(temp_name)
            else:
                send_message(chat_id, f"{TICK}\n{out}\n{TICK}")

        except Exception as e:
            current_process = None
            send_message(chat_id, f"Error: {e}")

    thread = threading.Thread(target=target)
    thread.start()

def take_screenshot(chat_id):
    filename = "screenshot.png"
    try:
        if os.name == 'nt':
            # Windows: Use PowerShell method which is more reliable for GH Actions
            ps_cmd = (
                "powershell -command \"[Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null; "
                "[Reflection.Assembly]::LoadWithPartialName('System.Drawing') | Out-Null; "
                "$Screen = [System.Windows.Forms.Screen]::PrimaryScreen; "
                "$Bitmap = New-Object System.Drawing.Bitmap($Screen.Bounds.Width, $Screen.Bounds.Height); "
                "$Graphics = [System.Drawing.Graphics]::FromImage($Bitmap); "
                "$Graphics.CopyFromScreen(0, 0, 0, 0, $Bitmap.Size); "
                "$Bitmap.Save('screenshot.png', [System.Drawing.Imaging.ImageFormat]::Png); "
                "$Graphics.Dispose(); $Bitmap.Dispose();\""
            )
            subprocess.run(ps_cmd, shell=True, check=True)
        else:
            # Linux/Other: Try Pillow
            from PIL import ImageGrab
            screenshot = ImageGrab.grab()
            screenshot.save(filename)

        send_photo(chat_id, filename)
        if os.path.exists(filename):
            os.remove(filename)
    except Exception as e:
        # Final fallback attempt with Pillow
        try:
            from PIL import ImageGrab
            screenshot = ImageGrab.grab()
            screenshot.save(filename)
            send_photo(chat_id, filename)
            if os.path.exists(filename): os.remove(filename)
        except Exception as e2:
            send_message(chat_id, f"Screenshot error: {e}\nFallback error: {e2}\nNote: If on Linux, ensure 'xvfb' is running or 'scrot' is installed.")

WELCOME = (
    "*GitHub VM Bot*\n\n"
    "- *Python code* -> send directly\n"
    "- *Shell command* -> start with `/` or `pip`/`npm`/`git`/`python`\n"
    "- *Files* -> I will save them\n"
    "- `screen` -> take screenshot\n"
    "- `terminate` -> kill current task"
)

if not TOKEN:
    print("Error: TELEGRAM_BOT_TOKEN environment variable not set.")
    sys.exit(1)

print("GitHub VM Bot started.", flush=True)

try:
    r = requests.get(f"{BASE}/getUpdates", params={"offset": -1}, timeout=10)
    updates = r.json().get("result", [])
    offset = updates[0]["update_id"] + 1 if updates else 0
except:
    offset = 0

while True:
    try:
        r = requests.get(f"{BASE}/getUpdates", params={"offset": offset, "timeout": 30}, timeout=40)
        updates = r.json().get("result", [])

        for upd in updates:
            offset = upd["update_id"] + 1
            msg = upd.get("message", {})
            chat_id = msg.get("chat", {}).get("id")
            if not chat_id or (ALLOWED_CHAT_ID and chat_id != ALLOWED_CHAT_ID):
                continue

            # Handle Media
            file_id = None
            file_name = None
            if "document" in msg:
                file_id, file_name = msg["document"]["file_id"], msg["document"].get("file_name", "file")
            elif "photo" in msg:
                file_id, file_name = msg["photo"][-1]["file_id"], f"photo_{int(time.time())}.jpg"
            elif "video" in msg:
                file_id, file_name = msg["video"]["file_id"], msg["video"].get("file_name", "video.mp4")

            if file_id:
                send_message(chat_id, f"Downloading `{file_name}`...")
                download_file(chat_id, file_id, file_name)
                continue

            text = (msg.get("text") or "").strip()
            if not text: continue

            if text in ("/start", "/help"):
                send_message(chat_id, WELCOME); continue

            if text == "/stop":
                send_message(chat_id, "Stopping..."); sys.exit(0)

            if text.lower() == "screen":
                take_screenshot(chat_id); continue

            if text.lower() == "terminate":
                if current_process:
                    current_process.terminate()
                    send_message(chat_id, "Task terminated.")
                else:
                    send_message(chat_id, "No task running.")
                continue

            if text == "/ls" or text.startswith("/ls "):
                path = text.split(maxsplit=1)[1] if " " in text else "."
                try:
                    files = os.listdir(path)
                    send_message(chat_id, f"{TICK}\n" + ("\n".join(files) or "Empty") + f"\n{TICK}")
                except Exception as e: send_message(chat_id, f"Error: {e}")
                continue

            if text.startswith("/cd "):
                try:
                    os.chdir(text[4:].strip())
                    send_message(chat_id, f"Dir: `{os.getcwd()}`")
                except Exception as e: send_message(chat_id, f"Error: {e}")
                continue

            if text.startswith("/get "):
                for f in glob.glob(text[5:].strip()):
                    if os.path.isfile(f): send_document(chat_id, f)
                continue

            is_shell = text.startswith("/") or any(text.startswith(x) for x in ["pip ", "npm ", "git ", "python "])
            cmd = text[1:].strip() if text.startswith("/") else text

            send_message(chat_id, "Running...")
            run_command(chat_id, cmd, is_python=not is_shell)

    except Exception as e:
        print(f"Error: {e}"); time.sleep(5)
