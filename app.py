from flask import Flask, render_template, request, jsonify, Response, stream_with_context
from google import genai
from google.genai import types
from dotenv import load_dotenv
import webbrowser
import os
import datetime
import threading
import subprocess
import urllib.parse
import json

try:
    import speech_recognition as sr
except ImportError:
    sr = None

try:
    import pyttsx3
except ImportError:
    pyttsx3 = None

load_dotenv()

app = Flask(__name__)

# ── Configuration (loaded from .env) ─────────────────────
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash-lite")
FLASK_PORT = int(os.environ.get("FLASK_PORT", 5000))
FLASK_DEBUG = os.environ.get("FLASK_DEBUG", "false").lower() == "true"

gemini_client = None
gemini_chat = None

def init_gemini():
    global gemini_client, gemini_chat
    key = GEMINI_API_KEY.strip() if GEMINI_API_KEY else ""
    if not key or key == "YOUR_GEMINI_API_KEY_HERE":
        return None
    try:
        gemini_client = genai.Client(api_key=key)
        model_name = GEMINI_MODEL
        gemini_chat = gemini_client.chats.create(
            model=model_name,
            config=types.GenerateContentConfig(
                system_instruction=(
                    "You are Atlas, a smart, fast, and helpful personal AI assistant. "
                    "Keep your replies concise, helpful, and natural. "
                    "Format code snippets using markdown code blocks. "
                    "If the user asks you to open apps or search something, "
                    "confirm what action is being taken in a friendly way."
                )
            )
        )
        return gemini_chat
    except Exception as e:
        print(f"[Gemini Init Error] {e}")
        try:
            # Fallback to gemini-3.5-flash
            gemini_chat = gemini_client.chats.create(
                model="gemini-3.5-flash",
                config=types.GenerateContentConfig(
                    system_instruction="You are Atlas, a fast and helpful personal AI assistant."
                )
            )
            return gemini_chat
        except Exception:
            return None


# ── pyttsx3 (lazy init for wake word only) ────────────────
engine = None

def get_engine():
    global engine
    if pyttsx3 is None:
        return None
    if engine is None:
        try:
            engine = pyttsx3.init()
            engine.setProperty('rate', 175)
            engine.setProperty('volume', 1.0)
            voices = engine.getProperty('voices')
            for voice in voices:
                if "zira" in voice.name.lower() or "female" in voice.name.lower():
                    engine.setProperty('voice', voice.id)
                    break
        except Exception as e:
            print(f"[TTS Init Warning] {e}")
    return engine

def speak(text):
    """Used only for wake word audio feedback."""
    try:
        eng = get_engine()
        if eng:
            eng.say(text)
            eng.runAndWait()
    except Exception as e:
        print(f"[Voice Error] {e}")

# ── Memory ────────────────────────────────────────────────
MEMORY_FILE = "memory.txt"

def save_memory(user_msg, ai_reply):
    try:
        with open(MEMORY_FILE, "a", encoding="utf-8") as f:
            f.write(f"You: {user_msg}\nAtlas: {ai_reply}\n\n")
    except Exception as e:
        print(f"[Memory Save Info] {e}")

def ask_ai(prompt):
    """Send message to Google Gemini and get a reply."""
    global gemini_chat, gemini_client

    key = GEMINI_API_KEY.strip() if GEMINI_API_KEY else ""
    if not key or key == "YOUR_GEMINI_API_KEY_HERE":
        return (
            "⚠️ **Google Gemini API Key Required**\n\n"
            "Please paste your free Gemini API key into the `.env` file:\n"
            "```env\nGEMINI_API_KEY=your_key_here\n```\n"
            "🔗 You can generate a free key in seconds at: https://aistudio.google.com/"
        )

    if gemini_chat is None:
        init_gemini()

    if gemini_chat is None:
        return "⚠️ Error: Unable to initialize Google Gemini client. Please verify your API key."

    try:
        response = gemini_chat.send_message(prompt)
        reply = response.text.strip()
        save_memory(prompt, reply)
        return reply
    except Exception as e:
        err_str = str(e)
        # Handle model availability fallbacks
        if "404" in err_str or "not found" in err_str.lower() or "503" in err_str:
            for alt_model in ["gemini-3.5-flash", "gemini-3.8-flash"]:
                try:
                    gemini_chat = gemini_client.chats.create(
                        model=alt_model,
                        config=types.GenerateContentConfig(
                            system_instruction="You are Atlas, a fast and helpful personal AI assistant."
                        )
                    )
                    response = gemini_chat.send_message(prompt)
                    reply = response.text.strip()
                    save_memory(prompt, reply)
                    return reply
                except Exception:
                    continue
        return f"⚠️ Google Gemini Error: {err_str}"

# ── Multi-strategy Desktop Application & URL Launcher ────
def open_desktop_target(target, args=""):
    """Launch application or URL on the interactive user desktop with robust fallbacks."""
    if os.name != "nt":
        return False
    # Strategy 1: Windows Shell COM object (dispatched to the active user's desktop explorer session)
    try:
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        shell = win32com.client.Dispatch("Shell.Application")
        shell.ShellExecute(target, args, "", "open", 1)
        pythoncom.CoUninitialize()
        return True
    except Exception as e:
        print(f"[Launcher COM Info] {e}")

    # Strategy 2: Standard os.startfile
    try:
        if args:
            os.startfile(f"{target} {args}")
        else:
            os.startfile(target)
        return True
    except Exception as e:
        print(f"[Launcher startfile Info] {e}")

    # Strategy 3: Subprocess with Windows shell 'start'
    try:
        cmd_str = f'start "" "{target}" {args}'.strip()
        subprocess.Popen(cmd_str, shell=True)
        return True
    except Exception as e:
        print(f"[Launcher subprocess Info] {e}")

    return False

# ── Handle commands ───────────────────────────────────────
def handle_command(command):
    cmd = command.lower().strip()

    # YouTube search
    if "youtube" in cmd and ("search" in cmd or "play" in cmd or "find" in cmd or "listen" in cmd):
        query = cmd
        for word in ["search youtube for", "search youtube", "play on youtube", "listen to on youtube",
                     "find on youtube", "youtube search", "youtube play", "youtube"]:
            query = query.replace(word, "").strip()
        if not query:
            query = "music"
        url = f"https://www.youtube.com/results?search_query={urllib.parse.quote_plus(query)}"
        open_desktop_target(url)
        webbrowser.open(url)
        return {
            "response": f"🎬 Searching YouTube for **{query}**...\n\n[Open YouTube Results ↗]({url})",
            "url": url
        }

    # Open YouTube
    elif "youtube" in cmd:
        url = "https://www.youtube.com"
        open_desktop_target(url)
        webbrowser.open(url)
        return {
            "response": f"▶️ Opening YouTube!\n\n[Open YouTube ↗]({url})",
            "url": url
        }

    # Google search
    elif "google" in cmd or "search for" in cmd or cmd.startswith("search "):
        query = cmd
        for word in ["search google for", "google search for", "search for", "google", "search"]:
            query = query.replace(word, "").strip()
        if not query:
            query = "news"
        url = f"https://www.google.com/search?q={urllib.parse.quote_plus(query)}"
        open_desktop_target(url)
        webbrowser.open(url)
        return {
            "response": f"🔍 Googling **{query}**...\n\n[Open Google Search ↗]({url})",
            "url": url
        }

    # Open Notepad
    elif "notepad" in cmd or "text editor" in cmd:
        success = open_desktop_target("notepad.exe")
        return {"response": "📝 Opening Notepad!" if success else "⚠️ Could not launch Notepad."}

    # Open Calculator
    elif "calculator" in cmd or "calc" in cmd:
        success = open_desktop_target("calc.exe")
        return {"response": "🔢 Opening Calculator!" if success else "⚠️ Could not launch Calculator."}

    # Open File Explorer
    elif "file explorer" in cmd or "explorer" in cmd or "open files" in cmd or "my computer" in cmd:
        user_dir = os.path.expanduser("~")
        success = open_desktop_target("explorer.exe", user_dir)
        return {"response": "📁 Opening File Explorer!" if success else "⚠️ Could not launch File Explorer."}

    # Open Task Manager
    elif "task manager" in cmd or "taskmgr" in cmd or "task mgr" in cmd:
        success = open_desktop_target("taskmgr.exe")
        return {"response": "📊 Opening Task Manager!" if success else "⚠️ Could not launch Task Manager."}

    # Open Settings
    elif "settings" in cmd or "control panel" in cmd:
        success = open_desktop_target("ms-settings:")
        return {"response": "⚙️ Opening Windows Settings!" if success else "⚠️ Could not launch Settings."}

    # Open Paint
    elif "paint" in cmd or "mspaint" in cmd:
        success = open_desktop_target("mspaint.exe")
        return {"response": "🎨 Opening Paint!" if success else "⚠️ Could not launch Paint."}

    # Open Camera
    elif "camera" in cmd or "webcam" in cmd:
        success = open_desktop_target("microsoft.windows.camera:")
        return {"response": "📷 Opening Camera!" if success else "⚠️ Could not launch Camera."}

    # Current time
    elif "time" in cmd and ("what" in cmd or "current" in cmd or "now" in cmd):
        now = datetime.datetime.now().strftime("%I:%M %p")
        return {"response": f"🕐 It's currently {now}"}

    # Current date
    elif "date" in cmd and ("what" in cmd or "today" in cmd or "current" in cmd):
        today = datetime.datetime.now().strftime("%A, %B %d, %Y")
        return {"response": f"📅 Today is {today}"}

    # Open WhatsApp Web
    elif "whatsapp" in cmd:
        url = "https://web.whatsapp.com"
        open_desktop_target(url)
        webbrowser.open(url)
        return {
            "response": f"💬 Opening WhatsApp Web!\n\n[Open WhatsApp Web ↗]({url})",
            "url": url
        }

    # Open Gmail
    elif "gmail" in cmd or "email" in cmd or "mail" in cmd:
        url = "https://mail.google.com"
        open_desktop_target(url)
        webbrowser.open(url)
        return {
            "response": f"📧 Opening Gmail!\n\n[Open Gmail ↗]({url})",
            "url": url
        }

    # Open GitHub
    elif "github" in cmd:
        url = "https://github.com"
        open_desktop_target(url)
        webbrowser.open(url)
        return {
            "response": f"🐙 Opening GitHub!\n\n[Open GitHub ↗]({url})",
            "url": url
        }

    # Open ChatGPT
    elif "chatgpt" in cmd or "openai" in cmd:
        url = "https://chatgpt.com"
        open_desktop_target(url)
        webbrowser.open(url)
        return {
            "response": f"🤖 Opening ChatGPT!\n\n[Open ChatGPT ↗]({url})",
            "url": url
        }

    # Open Spotify
    elif "spotify" in cmd:
        url = "https://open.spotify.com"
        open_desktop_target(url)
        webbrowser.open(url)
        return {
            "response": f"🎵 Opening Spotify!\n\n[Open Spotify ↗]({url})",
            "url": url
        }

    # Open Netflix
    elif "netflix" in cmd:
        url = "https://www.netflix.com"
        open_desktop_target(url)
        webbrowser.open(url)
        return {
            "response": f"🎬 Opening Netflix!\n\n[Open Netflix ↗]({url})",
            "url": url
        }

    # Open Twitter/X
    elif "twitter" in cmd or "open x" in cmd or cmd == "x":
        url = "https://x.com"
        open_desktop_target(url)
        webbrowser.open(url)
        return {
            "response": f"🐦 Opening X / Twitter!\n\n[Open X ↗]({url})",
            "url": url
        }

    # Open Instagram
    elif "instagram" in cmd or "insta" in cmd:
        url = "https://www.instagram.com"
        open_desktop_target(url)
        webbrowser.open(url)
        return {
            "response": f"📸 Opening Instagram!\n\n[Open Instagram ↗]({url})",
            "url": url
        }

    # Shutdown PC
    elif "shutdown" in cmd or "shut down" in cmd:
        os.system("shutdown /s /t 10")
        return {"response": "⚠️ Shutting down your PC in 10 seconds!"}

    # Restart PC
    elif "restart" in cmd:
        os.system("shutdown /r /t 10")
        return {"response": "🔄 Restarting your PC in 10 seconds!"}

    # Lock PC
    elif "lock" in cmd and "pc" in cmd:
        os.system("rundll32.exe user32.dll,LockWorkStation")
        return {"response": "🔒 Locking your PC!"}

    # Everything else → Google Gemini AI
    else:
        ai_reply = ask_ai(command)
        return {"response": ai_reply}


# ── Wake Word Listener ────────────────────────────────────
wake_word_active = False
wake_word_status = {"listening": False, "triggered": False, "last_command": ""}

def wake_word_listener():
    if sr is None:
        print("[Wake Word] speech_recognition is not available in this environment.")
        wake_word_status["listening"] = False
        return
    recognizer = sr.Recognizer()
    recognizer.energy_threshold = 3000
    recognizer.dynamic_energy_threshold = True
    print("[Wake Word] Listening for 'Hey Atlas'...")

    while wake_word_active:
        try:
            with sr.Microphone() as source:
                recognizer.adjust_for_ambient_noise(source, duration=0.5)
                wake_word_status["listening"] = True
                audio = recognizer.listen(source, timeout=5, phrase_time_limit=4)
                wake_word_status["listening"] = False

            text = recognizer.recognize_google(audio).lower()
            print(f"[Wake Word] Heard: {text}")

            if "hey atlas" in text or "atlas" in text:
                print("[Wake Word] Triggered!")
                wake_word_status["triggered"] = True
                speak("Yes? I'm listening.")

                with sr.Microphone() as source:
                    recognizer.adjust_for_ambient_noise(source, duration=0.3)
                    print("[Wake Word] Waiting for command...")
                    cmd_audio = recognizer.listen(source, timeout=6, phrase_time_limit=8)

                command = recognizer.recognize_google(cmd_audio)
                print(f"[Wake Word] Command: {command}")
                wake_word_status["last_command"] = command
                wake_word_status["triggered"] = False

                # Wake word replies are spoken via pyttsx3 since
                # they come from outside the browser
                result = handle_command(command)
                reply_text = result.get("response", "") if isinstance(result, dict) else str(result)
                clean = ''.join(c for c in reply_text if ord(c) < 10000).strip()
                threading.Thread(target=speak, args=(clean,), daemon=True).start()

        except sr.WaitTimeoutError:
            wake_word_status["listening"] = False
        except sr.UnknownValueError:
            wake_word_status["listening"] = False
            wake_word_status["triggered"] = False
        except Exception as e:
            print(f"[Wake Word Error] {e}")
            wake_word_status["listening"] = False
            wake_word_status["error"] = str(e)
            break

wake_thread = None

# ── Flask routes ──────────────────────────────────────────
@app.route("/")
def home():
    return render_template("index.html")

@app.route("/chat", methods=["POST"])
def chat():
    data = request.json or {}
    user_input = data.get("message", "").strip()
    if not user_input:
        return jsonify({"response": "Please say something!"})
    result = handle_command(user_input)
    if isinstance(result, dict):
        return jsonify(result)
    return jsonify({"response": result})

@app.route("/chat/stream", methods=["POST"])
def chat_stream():
    data = request.json or {}
    user_input = data.get("message", "").strip()
    if not user_input:
        return jsonify({"response": "Please say something!"})

    cmd_lower = user_input.lower().strip()
    command_keywords = [
        "youtube", "google", "search for", "calculator", "calc", "settings",
        "time", "date", "whatsapp", "gmail", "github", "chatgpt", "spotify",
        "netflix", "twitter", "open x", "lock", "shutdown", "restart", "paint",
        "camera", "notepad", "file explorer", "task manager"
    ]
    if any(k in cmd_lower for k in command_keywords):
        result = handle_command(user_input)
        if isinstance(result, dict):
            return jsonify(result)
        return jsonify({"response": result})

    def generate():
        global gemini_chat, gemini_client
        key = GEMINI_API_KEY.strip() if GEMINI_API_KEY else ""
        if not key or key == "YOUR_GEMINI_API_KEY_HERE":
            yield f"data: {json.dumps({'chunk': '⚠️ Google Gemini API key required.'})}\n\n"
            yield "data: [DONE]\n\n"
            return
        if gemini_chat is None:
            init_gemini()
        if gemini_chat is None:
            yield f"data: {json.dumps({'chunk': '⚠️ Unable to initialize Gemini client.'})}\n\n"
            yield "data: [DONE]\n\n"
            return
        
        full_reply = []
        try:
            stream = gemini_chat.send_message_stream(user_input)
            for chunk in stream:
                if chunk.text:
                    full_reply.append(chunk.text)
                    yield f"data: {json.dumps({'chunk': chunk.text})}\n\n"
            save_memory(user_input, "".join(full_reply))
        except Exception as e:
            yield f"data: {json.dumps({'chunk': f'⚠️ Gemini Error: {str(e)}'})}\n\n"
        yield "data: [DONE]\n\n"

    return Response(stream_with_context(generate()), mimetype="text/event-stream")

@app.route("/wake/start", methods=["POST"])
def start_wake():
    global wake_word_active, wake_thread
    if not wake_word_active:
        wake_word_active = True
        wake_thread = threading.Thread(target=wake_word_listener, daemon=True)
        wake_thread.start()
        speak("Hey Atlas wake word is now active.")
        return jsonify({"status": "started"})
    return jsonify({"status": "already running"})

@app.route("/wake/stop", methods=["POST"])
def stop_wake():
    global wake_word_active
    wake_word_active = False
    speak("Wake word disabled.")
    return jsonify({"status": "stopped"})

@app.route("/wake/status")
def wake_status():
    return jsonify(wake_word_status)

@app.route("/voice/test", methods=["POST"])
def voice_test():
    return jsonify({
        "status": "ok",
        "message": "Hello! I am Atlas, your personal AI assistant. Voice is working perfectly!"
    })

@app.route("/clear", methods=["POST"])
def clear_conversation():
    global gemini_chat
    gemini_chat = None
    init_gemini()
    return jsonify({"status": "cleared"})

# ── Run ───────────────────────────────────────────────────
if __name__ == "__main__":
    print("=" * 55)
    print("  ATLAS AI — Personal Desktop Assistant")
    print(f"  Open your browser: http://127.0.0.1:{FLASK_PORT}")
    print("=" * 55)
    app.run(port=FLASK_PORT, debug=FLASK_DEBUG)
