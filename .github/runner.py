import requests
import subprocess
import tempfile
import time
import sys
import os

TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
BASE  = f"https://api.telegram.org/bot{TOKEN}"
TICK  = chr(96) * 3

def send(chat_id, text, parse_mode='Markdown'):
    try:
        requests.post(
            f"{BASE}/sendMessage",
            json={"chat_id": chat_id, "text": text, "parse_mode": parse_mode},
            timeout=10
        )
    except Exception as e:
        print(f"[send error] {e}", flush=True)

def run_python(code):
    with tempfile.NamedTemporaryFile(mode="w", suffix=".py",
                                     delete=False, encoding="utf-8") as f:
        f.write(code)
        fname = f.name
    try:
        r = subprocess.run(
            [sys.executable, fname],
            capture_output=True, text=True, timeout=60
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
    if command.startswith('pip ') or command == 'pip':
        command = f'{sys.executable} -m {command}'
    elif command.startswith('pip3 ') or command == 'pip3':
        command = f'{sys.executable} -m pip{command[4:]}'
    try:
        r = subprocess.run(
            command, shell=True, capture_output=True, text=True, timeout=60
        )
        return (r.stdout + r.stderr).strip() or "(No output)"
    except subprocess.TimeoutExpired:
        return "Error: Timed out (60s limit)"
    except Exception as e:
        return f"Error: {e}"

WELCOME = (
    "*GitHub Windows VM Bot*\n\n"
    "Running on Windows Server! Send me:\n\n"
    "- *Python code* -> send it directly\n"
    "- *Shell/CMD command* -> start with /\n\n"
    "*Examples:*\n"
    "print('Hello from Windows!')\n"
    "/pip install requests\n"
    "/python --version\n"
    "/dir"
)

print("GitHub Windows VM Bot started.", flush=True)
print(f"Python: {sys.version}", flush=True)

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
            text    = (msg.get("text") or "").strip()

            if not text or not chat_id:
                continue

            print(f"[chat {chat_id}] {text[:60]}", flush=True)

            if text in ("/start", "/help"):
                send(chat_id, WELCOME)
                continue

            if text.startswith("/"):
                command = text[1:].strip()
                if not command:
                    send(chat_id, 'Provide a command after /.')
                    continue
                send(chat_id, 'Running on Windows VM...')
                out = run_shell(command)
                if len(out) > 3800:
                    out = out[:3800] + '\n...(truncated)'
                send(chat_id, f"{TICK}\n{out}\n{TICK}")
            else:
                send(chat_id, 'Running Python on Windows VM...')
                out = run_python(text)
                if len(out) > 3800:
                    out = out[:3800] + '\n...(truncated)'
                send(chat_id, f"{TICK}\n{out}\n{TICK}")

    except requests.exceptions.Timeout:
        pass
    except Exception as e:
        print(f"[loop error] {e}", flush=True)
        time.sleep(5)