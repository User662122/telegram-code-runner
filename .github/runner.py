import requests
import subprocess
import tempfile
import time
import sys
import os
import glob

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
# Optional: restrict to a specific chat ID for security
ALLOWED_CHAT_ID = os.environ.get("ALLOWED_CHAT_ID")
if ALLOWED_CHAT_ID:
    ALLOWED_CHAT_ID = int(ALLOWED_CHAT_ID)

BASE  = f"https://api.telegram.org/bot{TOKEN}"
TICK  = chr(96) * 3

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

def download_file(chat_id, file_id, file_name):
    try:
        # Sanitize filename to prevent path traversal
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

def run_python(code):
    with tempfile.NamedTemporaryFile(mode="w", suffix=".py",
                                     delete=False, encoding="utf-8") as f:
        f.write(code)
        fname = f.name
    try:
        r = subprocess.run(
            [sys.executable, fname],
            capture_output=True, text=True, timeout=60,
            cwd=os.getcwd()
        )
        return (r.stdout + r.stderr).strip() or "(No output)"
    except subprocess.TimeoutExpired:
        return "Error: Timed out (60s limit)"
    except Exception as e:
        return f"Error: {e}"
    finally:
        try:
            os.unlink(fname)
        except:
            pass

def run_shell(command):
    try:
        r = subprocess.run(
            command, shell=True, capture_output=True, text=True, timeout=60,
            cwd=os.getcwd()
        )
        return (r.stdout + r.stderr).strip() or "(No output)"
    except subprocess.TimeoutExpired:
        return "Error: Timed out (60s limit)"
    except Exception as e:
        return f"Error: {e}"

WELCOME = (
    "*GitHub VM Bot*\n\n"
    "Send me:\n"
    "- *Python code* -> send directly\n"
    "- *Shell command* -> start with `/` or `pip` or `npm` etc.\n"
    "- *Files/Images/Videos* -> I will save them to the current directory\n\n"
    "*Commands:*\n"
    "- `/ls` : List files\n"
    "- `/get <filename>` : Download a file\n"
    "- `/cd <dir>` : Change directory\n"
    "- `/stop` : Stop the runner"
)

if not TOKEN:
    print("Error: TELEGRAM_BOT_TOKEN environment variable not set.")
    sys.exit(1)

print("GitHub VM Bot started.", flush=True)

offset = 0
while True:
    try:
        r = requests.get(
            f"{BASE}/getUpdates",
            params={"offset": offset, "timeout": 30},
            timeout=40
        )
        updates = r.json().get("result", [])

        for upd in updates:
            offset = upd["update_id"] + 1
            msg     = upd.get("message", {})
            chat_id = msg.get("chat", {}).get("id")

            if not chat_id:
                continue

            # Security: check if chat_id is allowed
            if ALLOWED_CHAT_ID and chat_id != ALLOWED_CHAT_ID:
                print(f"Unauthorized access attempt from chat_id: {chat_id}")
                continue

            file_id = None
            file_name = None

            # Handle Documents
            if "document" in msg:
                doc = msg["document"]
                file_id = doc["file_id"]
                file_name = doc.get("file_name", "uploaded_file")

            # Handle Photos
            elif "photo" in msg:
                photo = msg["photo"][-1]
                file_id = photo["file_id"]
                file_name = f"photo_{int(time.time())}.jpg"

            # Handle Video
            elif "video" in msg:
                vid = msg["video"]
                file_id = vid["file_id"]
                file_name = vid.get("file_name", f"video_{int(time.time())}.mp4")

            # Handle Audio
            elif "audio" in msg:
                aud = msg["audio"]
                file_id = aud["file_id"]
                file_name = aud.get("file_name", f"audio_{int(time.time())}.mp3")

            if file_id:
                send_message(chat_id, f"Downloading `{file_name}`...")
                download_file(chat_id, file_id, file_name)
                continue

            text = (msg.get("text") or "").strip()
            if not text:
                continue

            print(f"[chat {chat_id}] {text[:60]}", flush=True)

            if text in ("/start", "/help"):
                send_message(chat_id, WELCOME)
                continue

            if text == "/stop":
                send_message(chat_id, "Stopping runner...")
                sys.exit(0)

            if text == "/ls" or text.startswith("/ls "):
                args = text.split(maxsplit=1)
                path = args[1] if len(args) > 1 else "."
                try:
                    files = os.listdir(path)
                    if not files:
                        send_message(chat_id, f"Directory `{path}` is empty.")
                    else:
                        out = "\n".join(files)
                        send_message(chat_id, f"{TICK}\n{out}\n{TICK}")
                except Exception as e:
                    send_message(chat_id, f"Error: {e}")
                continue

            if text.startswith("/cd "):
                new_dir = text[4:].strip()
                try:
                    os.chdir(new_dir)
                    send_message(chat_id, f"Changed directory to `{os.getcwd()}`")
                except Exception as e:
                    send_message(chat_id, f"Error: {e}")
                continue

            if text.startswith("/get "):
                file_pattern = text[5:].strip()
                files = glob.glob(file_pattern)
                if not files:
                    send_message(chat_id, f"No files found matching `{file_pattern}`")
                else:
                    for f in files:
                        if os.path.isfile(f):
                            send_document(chat_id, f)
                        else:
                            send_message(chat_id, f"`{f}` is not a file.")
                continue

            is_shell = text.startswith("/") or \
                       text.startswith("pip ") or text == "pip" or \
                       text.startswith("npm ") or text == "npm" or \
                       text.startswith("git ") or text == "git" or \
                       text.startswith("python ") or text == "python"

            if is_shell:
                command = text[1:].strip() if text.startswith("/") else text
                send_message(chat_id, f"Running shell command...")
                out = run_shell(command)
                if len(out) > 3800:
                    with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
                        f.write(out)
                        temp_name = f.name
                    send_document(chat_id, temp_name)
                    os.unlink(temp_name)
                else:
                    send_message(chat_id, f"{TICK}\n{out}\n{TICK}")
            else:
                send_message(chat_id, 'Running Python code...')
                out = run_python(text)
                if len(out) > 3800:
                    with tempfile.NamedTemporaryFile(mode="w", suffix=".txt", delete=False) as f:
                        f.write(out)
                        temp_name = f.name
                    send_document(chat_id, temp_name)
                    os.unlink(temp_name)
                else:
                    send_message(chat_id, f"{TICK}\n{out}\n{TICK}")

    except requests.exceptions.Timeout:
        pass
    except KeyboardInterrupt:
        sys.exit(0)
    except Exception as e:
        print(f"[loop error] {e}", flush=True)
        time.sleep(5)
