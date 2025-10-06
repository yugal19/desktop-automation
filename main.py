import os
import time
import threading
import asyncio
from dotenv import load_dotenv

load_dotenv()

from deepgram import DeepgramClient, LiveTranscriptionEvents, LiveOptions
import pyaudio
import subprocess
import webbrowser

# Tray icon
import pystray
from PIL import Image, ImageDraw

# interpreter and actions
from interpreter import parse_command
import actions

# Optional TTS
try:
    import pyttsx3

    tts_engine = pyttsx3.init()
except Exception:
    tts_engine = None

DEEPGRAM_KEY = os.getenv("DEEPGRAM_API_KEY")
if not DEEPGRAM_KEY:
    raise ValueError("Set DEEPGRAM_API_KEY in .env")

RATE = 16000
CHUNK = 1024
FORMAT = pyaudio.paInt16
CHANNELS = 1

_listening_flag = threading.Event()
_listening_flag.set()  # start listening by default
_should_stop = threading.Event()

# dictation state: when active, transcripts are appended to the target (notepad|word)
dictation_state = {"active": False, "target": None}

# short-term fragment buffer: store last unknown transcript to try combining
_fragment_buf = {"text": None, "time": 0.0}


def speak_feedback(text: str):
    if tts_engine:
        try:
            tts_engine.say(text)
            tts_engine.runAndWait()
        except Exception:
            pass


def _create_image(w=64, h=64, color=(40, 140, 240)):
    img = Image.new("RGB", (w, h), color)
    d = ImageDraw.Draw(img)
    d.ellipse((14, 10, 50, 46), fill=(255, 255, 255))
    d.rectangle((28, 36, 36, 54), fill=(255, 255, 255))
    return img


def _execute_command_safe(cmd: dict, raw_transcript: str):
    """
    Execute parsed command; runs in its own thread to avoid blocking Deepgram callback.
    """
    intent = cmd.get("intent")
    try:
        if intent == "start_dictation":
            target = cmd.get("target", "notepad")
            dictation_state["active"] = True
            dictation_state["target"] = target
            msg = f"Dictation started for {target}."
            print(msg)
            speak_feedback(msg)
            return

        if intent == "stop_dictation":
            dictation_state["active"] = False
            dictation_state["target"] = None
            msg = "Dictation stopped."
            print(msg)
            speak_feedback(msg)
            return

        if intent == "next_line":
            target = dictation_state.get("target", "notepad")
            if target == "word":
                # Word: press Enter
                if actions._word_app:
                    actions._word_app.Selection.TypeParagraph()
            else:
                # Notepad: press Enter
                actions.write_in_notepad("", newline=True)
            msg = "➡️ Moved to next line."
            print(msg)
            speak_feedback(msg)
            return

        if intent == "save":
            target = cmd.get("target") or dictation_state.get("target")
            if target == "word":
                res = actions.save_in_word()
            elif target == "notepad":
                # use timestamped filename if none provided
                filename = f"note_{int(time.time())}.txt"
                res = actions.save_in_notepad(filename)
            else:
                # fallback: try both
                res = actions.save_in_notepad(None)
            print(res)
            speak_feedback(res)
            return

        if intent == "close":
            target = (cmd.get("target") or "").lower()
            if target == "word":
                res = actions.close_word()
            elif target == "notepad":
                res = actions.close_notepad()
            elif target in ["file explorer", "explorer", "windows explorer"]:
                try:
                    subprocess.Popen("taskkill /im explorer.exe /f", shell=True)
                    # Restart explorer for system stability
                    time.sleep(1)
                    subprocess.Popen("start explorer.exe", shell=True)
                    res = "✅ Closed all File Explorer windows."
                except Exception as e:
                    res = f"❌ Could not close File Explorer: {e}"
            else:
                exe_map = {
                    "brave": "brave.exe",
                    "chrome": "chrome.exe",
                    "vscode": "Code.exe",
                    "visual studio code": "Code.exe",
                    "calculator": "calc.exe",
                }
                exe_name = exe_map.get(target, f"{target}.exe")
                try:
                    subprocess.Popen(f"taskkill /im {exe_name} /f", shell=True)
                    res = f"✅ Closed {target}"
                except Exception as e:
                    res = f"❌ Could not close {target}: {e}"

            print(res)
            speak_feedback(res)
            return

        if intent == "write":
            target = cmd.get("target", "notepad")
            content = cmd.get("content") or raw_transcript
            if not content.strip():
                return

            if target == "word":
                res = actions.write_in_word(content)
            else:
                # type continuously, no newline
                res = actions.write_in_notepad(content, newline=False)

            print(f"📝 Writing (same line): {content}")
            return

        # open + write compound
        if intent == "open_and_write":
            app = cmd.get("app")
            content = cmd.get("content", "")
            # open app first
            open_res = actions.open_app(app)
            print(open_res)
            time.sleep(0.6)
            # then write content
            if "word" in app:
                res = actions.write_in_word(content)
            else:
                res = actions.write_in_notepad(content)
            print(res)
            speak_feedback(res)
            return

        # open app
        if intent == "open":
            app = cmd.get("app") or raw_transcript
            res = actions.open_app(app)
            print(res)
            speak_feedback(res)
            return

        # search web (google/wikipedia/youtube)
        if intent == "search_web":
            engine = cmd.get("engine", "google")
            q = cmd.get("query") or raw_transcript
            browser = cmd.get("browser")  # may be 'brave' or 'chrome'
            if not q:
                return
            if engine == "wikipedia":
                url = "https://en.wikipedia.org/wiki/" + q.replace(" ", "_")
            elif engine == "youtube":
                url = "https://www.youtube.com/results?search_query=" + q.replace(
                    " ", "+"
                )
            else:  # google
                url = "https://www.google.com/search?q=" + q.replace(" ", "+")
            try:
                if browser:
                    # try to open specific browser with url
                    subprocess.Popen(f'start {browser} "{url}"', shell=True)
                else:
                    webbrowser.open(url)
                print(f"✅ Opened web search for: {q}")
                speak_feedback("Opened web search")
            except Exception as e:
                print("❌ Error opening web search:", e)
            return

        # file/folder search (Explorer)
        if intent == "search":
            name = cmd.get("name") or raw_transcript
            res = actions.search_in_explorer(name)
            print(res)
            speak_feedback(res)
            return

        if intent == "stop":
            msg = "🛑 Stopping assistant by voice command."
            print(msg)
            speak_feedback("Assistant stopped")
            quit_app()  # shuts tray & exits
            return
        # unknown - don't stop listening; if dictation is ON, this should have been handled earlier
        if intent == "unknown":
            print("⚠️ Unknown command:", raw_transcript)
            return

        print("⚠️ Unhandled intent:", intent)
    except Exception as e:
        print("❌ Error executing command:", e)


async def _deepgram_runner():
    """
    Connect to Deepgram and stream from microphone. The on_transcript callback
    does light parsing + buffering, and dispatches execution to a thread.
    """
    client = DeepgramClient(DEEPGRAM_KEY)
    dg_conn = client.listen.websocket.v("1")

    # keep fragment buffer and dictation state accessible
    # _fragment_buf = {"text": None, "time": 0.0} defined at module scope

    def on_transcript(self, result, **kwargs):
        if not _listening_flag.is_set():
            return

        try:
            transcript = result.channel.alternatives[0].transcript.strip()
        except Exception:
            return
        if not transcript:
            return

        # quick cleanup: remove leading/trailing punctuation
        clean_transcript = transcript.strip()
        print("📝 Transcript:", clean_transcript)

        # If dictation is active, handle dictation first (append mode)
        if dictation_state["active"]:
            # check for dictation-control phrases in the new transcript
            ctrl = parse_command(clean_transcript)
            if ctrl.get("intent") in ("stop_dictation", "save", "close"):
                # execute control actions
                threading.Thread(
                    target=_execute_command_safe,
                    args=(ctrl, clean_transcript),
                    daemon=True,
                ).start()
                return
            # else append content directly to target
            target = dictation_state.get("target", "notepad")
            if target == "word":
                threading.Thread(
                    target=_execute_command_safe,
                    args=(
                        {
                            "intent": "write",
                            "target": "word",
                            "content": clean_transcript,
                        },
                        clean_transcript,
                    ),
                    daemon=True,
                ).start()
            else:
                threading.Thread(
                    target=_execute_command_safe,
                    args=(
                        {
                            "intent": "write",
                            "target": "notepad",
                            "content": clean_transcript,
                        },
                        clean_transcript,
                    ),
                    daemon=True,
                ).start()
            return

        # Not dictation: try parse
        cmd = parse_command(clean_transcript)

        # If parse returned actionable intent, execute and clear fragment buffer
        if cmd.get("intent") and cmd.get("intent") != "unknown":
            threading.Thread(
                target=_execute_command_safe, args=(cmd, clean_transcript), daemon=True
            ).start()
            _fragment_buf["text"] = None
            _fragment_buf["time"] = 0.0
            return

        # If unknown: attempt to combine with previous unknown fragment if recent
        now = time.time()
        prev_text = _fragment_buf.get("text")
        prev_time = _fragment_buf.get("time", 0)
        if prev_text and (now - prev_time) <= 2.0:
            combined = prev_text + " " + clean_transcript
            combined_cmd = parse_command(combined)
            if combined_cmd.get("intent") and combined_cmd.get("intent") != "unknown":
                threading.Thread(
                    target=_execute_command_safe,
                    args=(combined_cmd, combined),
                    daemon=True,
                ).start()
                _fragment_buf["text"] = None
                _fragment_buf["time"] = 0.0
                return

        # otherwise buffer this transcript as possible fragment
        _fragment_buf["text"] = clean_transcript
        _fragment_buf["time"] = now
        # Also log unknowns but do not stop listening
        print("⚠️ Unknown command (buffered):", clean_transcript)
        return

    dg_conn.on(LiveTranscriptionEvents.Transcript, on_transcript)

    opts = LiveOptions(
        model="nova-2",
        punctuate=True,
        interim_results=False,
        encoding="linear16",
        sample_rate=RATE,
    )
    started = dg_conn.start(opts)
    if not started:
        raise RuntimeError("❌ Failed to start Deepgram connection")

    p = pyaudio.PyAudio()
    try:
        stream = p.open(
            format=FORMAT,
            channels=CHANNELS,
            rate=RATE,
            input=True,
            frames_per_buffer=CHUNK,
        )
    except Exception as e:
        print("❌ PyAudio error:", e)
        dg_conn.finish()
        p.terminate()
        return

    print("🎙 Assistant listening (Deepgram). Say 'stop listening' or use tray to stop.")
    while not _should_stop.is_set():
        try:
            data = stream.read(CHUNK, exception_on_overflow=False)
            dg_conn.send(data)
        except Exception as e:
            print("⚠️ Audio read/send error:", e)
            time.sleep(0.1)
            continue

    # end cleanup
    try:
        dg_conn.finish()
    except Exception:
        pass
    try:
        stream.stop_stream()
        stream.close()
    except Exception:
        pass
    p.terminate()
    print("🛑 Deepgram loop terminated.")


# Thread runner helper
def _run_loop(loop):
    asyncio.set_event_loop(loop)
    loop.run_until_complete(_deepgram_runner())
    loop.close()


# Tray integration & control
_loop_thread = None
_async_loop = None
_tray_icon = None


def start_listening(_=None):
    global _loop_thread, _async_loop
    if _loop_thread and _loop_thread.is_alive():
        _listening_flag.set()
        print("Listening resumed.")
        return
    _async_loop = asyncio.new_event_loop()
    _loop_thread = threading.Thread(target=_run_loop, args=(_async_loop,), daemon=True)
    _listening_flag.set()
    _should_stop.clear()
    _loop_thread.start()
    print("Started listening thread.")


def stop_listening(_=None):
    _listening_flag.clear()
    print("Listening paused.")


def toggle_listening(icon, item):
    if _listening_flag.is_set():
        stop_listening()
    else:
        if not (_loop_thread and _loop_thread.is_alive()):
            start_listening()
        else:
            _listening_flag.set()
    # icon.update_menu may exist on some pystray builds; ignore if not


def quit_app(icon=None, item=None):
    print("Quitting assistant...")
    _should_stop.set()
    _listening_flag.clear()
    try:
        icon.stop()
    except Exception:
        pass
    time.sleep(0.4)
    os._exit(0)


def run_tray():
    global _tray_icon
    image = _create_image()
    menu = pystray.Menu(
        pystray.MenuItem(
            lambda item: (
                "Pause Listening" if _listening_flag.is_set() else "Resume Listening"
            ),
            toggle_listening,
        ),
        pystray.MenuItem("Quit", quit_app),
    )
    _tray_icon = pystray.Icon("DesktopAssistant", image, "Desktop Assistant", menu)
    start_listening()
    _tray_icon.run()


if __name__ == "__main__":
    run_tray()
