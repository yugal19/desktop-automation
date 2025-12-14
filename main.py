# import os
# import time
# import threading
# import asyncio
# from dotenv import load_dotenv
# import re

# load_dotenv()

# from deepgram import DeepgramClient, LiveTranscriptionEvents, LiveOptions
# import pyaudio
# import subprocess
# import webbrowser

# # Tray icon
# import pystray
# from PIL import Image, ImageDraw

# # interpreter and actions
# from interpreter import parse_command
# import actions

# # Optional TTS
# try:
#     import pyttsx3

#     tts_engine = pyttsx3.init()
# except Exception:
#     tts_engine = None

# DEEPGRAM_KEY = os.getenv("DEEPGRAM_API_KEY")
# if not DEEPGRAM_KEY:
#     raise ValueError("Set DEEPGRAM_API_KEY in .env")

# RATE = 16000
# CHUNK = 1024
# FORMAT = pyaudio.paInt16
# CHANNELS = 1

# _listening_flag = threading.Event()
# _listening_flag.set()  # start listening by default
# _should_stop = threading.Event()

# # dictation state: when active, transcripts are appended to the target (notepad|word|excel)
# dictation_state = {"active": False, "target": None}

# # short-term fragment buffer: store last unknown transcript to try combining
# _fragment_buf = {"text": None, "time": 0.0}


# def speak_feedback(text: str):
#     if tts_engine:
#         try:
#             tts_engine.say(text)
#             tts_engine.runAndWait()
#         except Exception:
#             pass
#     # always print as well for visibility
#     if text:
#         print(text)


# def _create_image(w=64, h=64, color=(40, 140, 240)):
#     img = Image.new("RGB", (w, h), color)
#     d = ImageDraw.Draw(img)
#     d.ellipse((14, 10, 50, 46), fill=(255, 255, 255))
#     d.rectangle((28, 36, 36, 54), fill=(255, 255, 255))
#     return img


# def _run_excel_write(text: str):
#     """
#     Detect cell references from spoken text and write into Excel properly.
#     Uses actions.continue_writing_excel() for actual writing.
#     """
#     try:
#         # Try to detect any cell reference like 'in cell A1' or 'in B2'
#         match = re.search(
#             r"\bcell\s*([A-Za-z]+\d+)\b|\bin\s*([A-Za-z]+\d+)\b|\binto\s*([A-Za-z]+\d+)\b",
#             text,
#             re.IGNORECASE,
#         )

#         # Extract detected cell (if any)
#         cell = None
#         if match:
#             cell = (match.group(1) or match.group(2) or match.group(3)).upper()

#         # Clean the spoken text to only the value to write
#         value = re.sub(
#             r"\b(write|right|in\s*cell\s*[A-Za-z]+\d+|in\s*[A-Za-z]+\d+|into\s*[A-Za-z]+\d+)\b",
#             "",
#             text,
#             flags=re.IGNORECASE,
#         ).strip()

#         # If no value found, skip
#         if not value:
#             res = "⚠️ Nothing to write."
#             print(res)
#             speak_feedback(res)
#             return

#         # Call Excel writer (cell-aware)
#         if cell:
#             # Explicit cell target
#             clean_text = f"Write {value} in cell {cell}"
#             res = actions.continue_writing_excel(clean_text)
#         else:
#             # No cell mentioned — default to append
#             res = actions.continue_writing_excel(value)

#     except Exception as e:
#         res = f"❌ Error writing to Excel: {e}"

#     print(res)
#     speak_feedback(res)


# def _execute_command_safe(cmd: dict, raw_transcript: str):
#     """
#     Execute parsed command; runs in its own thread to avoid blocking callbacks.
#     """
#     # defensive: ensure cmd is a dict
#     if not cmd or not isinstance(cmd, dict):
#         print("⚠️ Empty or invalid command; ignoring.")
#         return

#     intent = cmd.get("intent")
#     try:
#         # --- Dictation control ---
#         if intent == "start_dictation":
#             target = cmd.get("target", "notepad")
#             dictation_state["active"] = True
#             dictation_state["target"] = target
#             msg = f"Dictation started for {target}."
#             print(msg)
#             speak_feedback(msg)
#             return

#         if intent == "stop_dictation":
#             dictation_state["active"] = False
#             dictation_state["target"] = None
#             msg = "Dictation stopped."
#             print(msg)
#             speak_feedback(msg)
#             return

#         if intent == "next_line":
#             target = dictation_state.get("target", "notepad")
#             if target == "word":
#                 # Word: press Enter paragraph
#                 try:
#                     if actions._word_app:
#                         actions._word_app.Selection.TypeParagraph()
#                 except Exception:
#                     pass
#             else:
#                 # Notepad: press Enter (via actions helper)
#                 actions.write_in_notepad("", newline=True)
#             msg = "➡️ Moved to next line."
#             print(msg)
#             speak_feedback(msg)
#             return

#         # --- Save / Close / Close-other apps ---
#         if intent == "save":
#             target = cmd.get("target") or dictation_state.get("target")
#             if target == "word":
#                 res = actions.save_in_word()
#             elif target == "notepad":
#                 filename = f"note_{int(time.time())}.txt"
#                 res = actions.save_in_notepad(filename)
#             else:
#                 # fallback to notepad save
#                 res = actions.save_in_notepad(None)
#             print(res)
#             speak_feedback(res)
#             return

#         if intent == "close":
#             target = (cmd.get("target") or "").lower()
#             if target == "word":
#                 res = actions.close_word()
#             elif target == "notepad":
#                 res = actions.close_notepad()
#             elif target in ["file explorer", "explorer", "windows explorer"]:
#                 try:
#                     subprocess.Popen("taskkill /im explorer.exe /f", shell=True)
#                     time.sleep(1)
#                     subprocess.Popen("start explorer.exe", shell=True)
#                     res = "✅ Closed all File Explorer windows."
#                 except Exception as e:
#                     res = f"❌ Could not close File Explorer: {e}"
#             else:
#                 exe_map = {
#                     "brave": "brave.exe",
#                     "chrome": "chrome.exe",
#                     "vscode": "Code.exe",
#                     "visual studio code": "Code.exe",
#                     "calculator": "calc.exe",
#                     "excel": "excel.exe",
#                 }
#                 exe_name = exe_map.get(target, f"{target}.exe")
#                 try:
#                     subprocess.Popen(f"taskkill /im {exe_name} /f", shell=True)
#                     res = f"✅ Closed {target}"
#                 except Exception as e:
#                     res = f"❌ Could not close {target}: {e}"

#             print(res)
#             speak_feedback(res)
#             return

#         # ---------- Open Excel explicitly ----------
#         if intent == "open":
#             app = (cmd.get("app") or raw_transcript or "").strip()
#             if not app:
#                 msg = "⚠️ No app specified to open."
#                 print(msg)
#                 speak_feedback(msg)
#                 return

#             # If user asked to open Excel, use the COM open helper so we attach to same workbook.
#             if "excel" in app.lower():
#                 res = actions.open_excel()
#             else:
#                 res = actions.open_app(app)
#             print(res)
#             speak_feedback(res)
#             return

#         # generic `write` intent (may be notepad/word/excel depending on parser)
#         if intent == "write":
#             target = cmd.get("target", "notepad")
#             content = cmd.get("content") or raw_transcript or ""
#             if not content.strip():
#                 return

#             if target == "word":
#                 res = actions.write_in_word(content)
#             else:
#                 res = actions.write_in_notepad(content, newline=False)

#             print(f"📝 Writing (same line): {content}")
#             speak_feedback(res)
#             return

#         # explicit excel write command
#         if intent == "write_excel":
#             cell = cmd.get("cell")
#             content = cmd.get("content") or raw_transcript or ""
#             if not cell:
#                 # Try parsing cell name directly from voice
#                 match = re.search(r"cell\s*([a-zA-Z]+\d+)", raw_transcript.lower())
#                 if match:
#                     cell = match.group(1).upper()

#             if not cell:
#                 msg = "⚠️ Could not detect target Excel cell."
#                 print(msg)
#                 speak_feedback(msg)
#                 return

#             res = actions.write_in_excel(content, cell)
#             print(res)
#             speak_feedback(res)
#             return

#         # ---------- Open and write Excel (compound) ----------
#         if intent == "open_and_write_excel":
#             app = cmd.get("app", "excel")
#             content = cmd.get("content") or raw_transcript or ""
#             cell = cmd.get("cell")

#             # open via COM helper
#             open_res = actions.open_excel()
#             print(open_res)
#             speak_feedback(open_res)
#             time.sleep(0.4)

#             # write (cell or append)
#             if cell:
#                 res = actions.write_in_excel(content, cell)
#             else:
#                 res = actions.continue_writing_excel(content)
#             print(res)
#             speak_feedback(res)
#             return

#         # web search
#         if intent == "search_web":
#             engine = cmd.get("engine", "google")
#             q = cmd.get("query") or raw_transcript or ""
#             browser = cmd.get("browser")
#             if not q:
#                 return
#             if engine == "wikipedia":
#                 url = "https://en.wikipedia.org/wiki/" + q.replace(" ", "_")
#             elif engine == "youtube":
#                 url = "https://www.youtube.com/results?search_query=" + q.replace(
#                     " ", "+"
#                 )
#             else:
#                 url = "https://www.google.com/search?q=" + q.replace(" ", "+")
#             try:
#                 if browser:
#                     subprocess.Popen(f'start {browser} "{url}"', shell=True)
#                 else:
#                     webbrowser.open(url)
#                 print(f"✅ Opened web search for: {q}")
#                 speak_feedback("Opened web search")
#             except Exception as e:
#                 print("❌ Error opening web search:", e)
#             return

#         # explorer search
#         if intent == "search":
#             name = cmd.get("name") or raw_transcript or ""
#             res = actions.search_in_explorer(name)
#             print(res)
#             speak_feedback(res)
#             return

#         # stop assistant
#         if intent == "stop":
#             msg = "🛑 Stopping assistant by voice command."
#             print(msg)
#             speak_feedback("Assistant stopped")
#             quit_app()
#             return

#         # unknown
#         if intent == "unknown":
#             print("⚠️ Unknown command:", raw_transcript)
#             return

#         print("⚠️ Unhandled intent:", intent)

#     except Exception as e:
#         print("❌ Error executing command:", e)


# async def _deepgram_runner():
#     """
#     Connect to Deepgram and stream from microphone. The on_transcript callback
#     does light parsing + buffering, and dispatches execution to a thread.
#     """
#     client = DeepgramClient(DEEPGRAM_KEY)
#     dg_conn = client.listen.websocket.v("1")

#     def on_transcript(self, result, **kwargs):
#         if not _listening_flag.is_set():
#             return

#         try:
#             transcript = result.channel.alternatives[0].transcript.strip()
#         except Exception:
#             return
#         if not transcript:
#             return

#         # quick cleanup: remove leading/trailing punctuation
#         clean_transcript = transcript.strip()
#         print("📝 Transcript:", clean_transcript)

#         # If dictation is active, handle dictation first (append mode)
#         if dictation_state["active"]:
#             # check for dictation-control phrases in the new transcript
#             ctrl = parse_command(clean_transcript)
#             if ctrl and ctrl.get("intent") in ("stop_dictation", "save", "close"):
#                 # execute control actions
#                 threading.Thread(
#                     target=_execute_command_safe,
#                     args=(ctrl, clean_transcript),
#                     daemon=True,
#                 ).start()
#                 return

#             # else append content directly to target
#             target = dictation_state.get("target", "notepad")
#             if target == "word":
#                 threading.Thread(
#                     target=_execute_command_safe,
#                     args=(
#                         {
#                             "intent": "write",
#                             "target": "word",
#                             "content": clean_transcript,
#                         },
#                         clean_transcript,
#                     ),
#                     daemon=True,
#                 ).start()
#             elif target == "excel":
#                 # For excel dictation, use COM-based continue_writing so focus isn't required.
#                 threading.Thread(
#                     target=_run_excel_write, args=(clean_transcript,), daemon=True
#                 ).start()
#             else:
#                 threading.Thread(
#                     target=_execute_command_safe,
#                     args=(
#                         {
#                             "intent": "write",
#                             "target": "notepad",
#                             "content": clean_transcript,
#                         },
#                         clean_transcript,
#                     ),
#                     daemon=True,
#                 ).start()
#             return

#         # Not dictation: try parse
#         cmd = parse_command(clean_transcript)
#         if not cmd or not isinstance(cmd, dict):
#             print("No valid Command (parse_command returned None or invalid).")
#             return

#         if cmd.get("intent") and cmd.get("intent") != "unknown":
#             threading.Thread(
#                 target=_execute_command_safe, args=(cmd, clean_transcript), daemon=True
#             ).start()
#             _fragment_buf["text"] = None
#             _fragment_buf["time"] = 0.0
#             return

#         # If unknown: attempt to combine with previous unknown fragment if recent
#         now = time.time()
#         prev_text = _fragment_buf.get("text")
#         prev_time = _fragment_buf.get("time", 0)
#         if prev_text and (now - prev_time) <= 2.0:
#             combined = prev_text + " " + clean_transcript
#             combined_cmd = parse_command(combined)
#             if combined_cmd.get("intent") and combined_cmd.get("intent") != "unknown":
#                 threading.Thread(
#                     target=_execute_command_safe,
#                     args=(combined_cmd, combined),
#                     daemon=True,
#                 ).start()
#                 _fragment_buf["text"] = None
#                 _fragment_buf["time"] = 0.0
#                 return

#         # otherwise buffer this transcript as possible fragment
#         _fragment_buf["text"] = clean_transcript
#         _fragment_buf["time"] = now
#         # Also log unknowns but do not stop listening
#         print("⚠️ Unknown command (buffered):", clean_transcript)
#         return

#     dg_conn.on(LiveTranscriptionEvents.Transcript, on_transcript)

#     opts = LiveOptions(
#         model="nova-2",
#         punctuate=True,
#         interim_results=False,
#         encoding="linear16",
#         sample_rate=RATE,
#     )
#     started = dg_conn.start(opts)
#     if not started:
#         raise RuntimeError("❌ Failed to start Deepgram connection")

#     p = pyaudio.PyAudio()
#     try:
#         stream = p.open(
#             format=FORMAT,
#             channels=CHANNELS,
#             rate=RATE,
#             input=True,
#             frames_per_buffer=CHUNK,
#         )
#     except Exception as e:
#         print("❌ PyAudio error:", e)
#         dg_conn.finish()
#         p.terminate()
#         return

#     print("🎙 Assistant listening (Deepgram). Say 'stop listening' or use tray to stop.")
#     while not _should_stop.is_set():
#         try:
#             data = stream.read(CHUNK, exception_on_overflow=False)
#             dg_conn.send(data)
#         except Exception as e:
#             print("⚠️ Audio read/send error:", e)
#             time.sleep(0.1)
#             continue

#     # end cleanup
#     try:
#         dg_conn.finish()
#     except Exception:
#         pass
#     try:
#         stream.stop_stream()
#         stream.close()
#     except Exception:
#         pass
#     p.terminate()
#     print("🛑 Deepgram loop terminated.")


# # Thread runner helper
# def _run_loop(loop):
#     asyncio.set_event_loop(loop)
#     loop.run_until_complete(_deepgram_runner())
#     loop.close()


# # Tray integration & control
# _loop_thread = None
# _async_loop = None
# _tray_icon = None


# def start_listening(_=None):
#     global _loop_thread, _async_loop
#     if _loop_thread and _loop_thread.is_alive():
#         _listening_flag.set()
#         print("Listening resumed.")
#         return
#     _async_loop = asyncio.new_event_loop()
#     _loop_thread = threading.Thread(target=_run_loop, args=(_async_loop,), daemon=True)
#     _listening_flag.set()
#     _should_stop.clear()
#     _loop_thread.start()
#     print("Started listening thread.")


# def stop_listening(_=None):
#     _listening_flag.clear()
#     print("Listening paused.")


# def toggle_listening(icon, item):
#     if _listening_flag.is_set():
#         stop_listening()
#     else:
#         if not (_loop_thread and _loop_thread.is_alive()):
#             start_listening()
#         else:
#             _listening_flag.set()
#     # icon.update_menu may exist on some pystray builds; ignore if not


# def quit_app(icon=None, item=None):
#     print("Quitting assistant...")
#     _should_stop.set()
#     _listening_flag.clear()
#     try:
#         icon.stop()
#     except Exception:
#         pass
#     time.sleep(0.4)
#     os._exit(0)


# def run_tray():
#     global _tray_icon
#     image = _create_image()
#     menu = pystray.Menu(
#         pystray.MenuItem(
#             lambda item: (
#                 "Pause Listening" if _listening_flag.is_set() else "Resume Listening"
#             ),
#             toggle_listening,
#         ),
#         pystray.MenuItem("Quit", quit_app),
#     )
#     _tray_icon = pystray.Icon("DesktopAssistant", image, "Desktop Assistant", menu)
#     start_listening()
#     _tray_icon.run()


# if __name__ == "__main__":
#     run_tray()


# main.py
# main.py (updated)
# voice_form_dictation_with_overwrite.py
# Updated assistant script with robust "explicit field correction & replacement" flow
# (Integrates Flow 2: mentioning a field name resets / overwrites that field)
# voice_form_dictation_with_overwrite.py
# Updated assistant script with robust "explicit field correction & replacement" flow
# (Integrates Flow 2: mentioning a field name resets / overwrites that field)

import os
import time
import threading
import asyncio
import re
import sys
import subprocess
import webbrowser
import json
from datetime import datetime

from dotenv import load_dotenv

load_dotenv()

from deepgram import DeepgramClient, LiveTranscriptionEvents, LiveOptions
import pyaudio

import pystray
from PIL import Image, ImageDraw

from interpreter import parse_command
import actions
import claude_writer

import web_socket_server
import web_form_controller

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


FORM_OVERWRITE_PENDING = {"active": False, "timer": None}

dictation_state = {"active": False, "target": None}

_fragment_buf = {"text": None, "time": 0.0}


FORM_SUBMISSION_BUFFER: dict = {}
_form_buffer_lock = threading.Lock()
SUBMISSIONS_DIR = "submissions"
os.makedirs(SUBMISSIONS_DIR, exist_ok=True)

CURRENT_FORM_FIELD = {"name": None}

OVERWRITE_PENDING_TIMEOUT = 4.0


def speak_feedback(text: str):
    if tts_engine:
        try:
            tts_engine.say(text)
            tts_engine.runAndWait()
        except Exception:
            pass
    if text:
        print(text)


def _create_image(w=64, h=64, color=(40, 140, 240)):
    img = Image.new("RGB", (w, h), color)
    d = ImageDraw.Draw(img)
    d.ellipse((14, 10, 50, 46), fill=(255, 255, 255))
    d.rectangle((28, 36, 36, 54), fill=(255, 255, 255))
    return img


def _focus_claude_window():
    try:
        import pygetwindow as gw

        for title in gw.getAllTitles():
            if not title:
                continue
            lower = title.lower()
            if "claude" in lower or "anthropic" in lower:
                wins = gw.getWindowsWithTitle(title)
                if wins:
                    w = wins[0]
                    try:
                        w.restore()
                    except Exception:
                        pass
                    try:
                        w.activate()
                    except Exception:
                        try:
                            w.minimize()
                            w.maximize()
                        except Exception:
                            pass
                    time.sleep(0.15)
                    return True
    except Exception:
        pass
    return False


def activate_claude_excel_mode():
    focused = _focus_claude_window()
    if focused:
        print("✅ Focused Claude Desktop for Excel operations.")
    else:
        print(
            "⚠️ Could not focus Claude Desktop. Please open and focus Claude manually."
        )
    dictation_state["target"] = "claude_excel"


def _run_excel_write(text: str):
    try:
        clean = text.strip()
        if not clean:
            msg = "⚠️ Nothing to send to Claude for Excel."
            print(msg)
            speak_feedback(msg)
            return
        claude_writer.push_text(clean)
    except Exception as e:
        res = f"❌ Error pushing Excel text to Claude: {e}"
        print(res)
        speak_feedback(res)


def _save_form_buffer_to_json(buffer: dict) -> str:
    if not buffer:
        data_to_save = {}
    else:
        data_to_save = dict(buffer)

    ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    filename = f"form_{ts}.json"
    path = os.path.join(SUBMISSIONS_DIR, filename)
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(
                {"timestamp": ts, "fields": data_to_save},
                f,
                ensure_ascii=False,
                indent=4,
            )
        return path
    except Exception as e:
        print("❌ Error saving form JSON:", e)
        return ""


def _clear_overwrite_pending_locked():
    """Must be called with _form_buffer_lock held."""
    FORM_OVERWRITE_PENDING["active"] = False
    timer = FORM_OVERWRITE_PENDING.get("timer")
    if timer:
        try:
            timer.cancel()
        except Exception:
            pass
    FORM_OVERWRITE_PENDING["timer"] = None


def _start_overwrite_pending_locked(
    field_name: str, timeout: float = OVERWRITE_PENDING_TIMEOUT
):
    """Set overwrite-pending and start/reset a timer that clears it after `timeout` seconds.
    Must be called with _form_buffer_lock held."""

    existing = FORM_OVERWRITE_PENDING.get("timer")
    if existing:
        try:
            existing.cancel()
        except Exception:
            pass

    FORM_OVERWRITE_PENDING["active"] = True

    def _timeout_clear():
        with _form_buffer_lock:
            FORM_OVERWRITE_PENDING["active"] = False
            FORM_OVERWRITE_PENDING["timer"] = None

    timer = threading.Timer(timeout, _timeout_clear)
    FORM_OVERWRITE_PENDING["timer"] = timer
    timer.daemon = True
    timer.start()


def _execute_command_safe(cmd: dict, raw_transcript: str):
    if not cmd or not isinstance(cmd, dict):
        print("⚠️ Empty or invalid command; ignoring.")
        return

    intent = cmd.get("intent")
    try:
        if intent == "start_dictation":
            target = cmd.get("target", "notepad")

            if target == "form":
                dictation_state["active"] = True
                dictation_state["target"] = "form"
                msg = (
                    "Form filling mode started. "
                    "Speak things like 'first name Yugal', 'surname Chandak', or 'submit form'."
                )
                print(msg)
                speak_feedback(msg)
                return

            if target == "excel":
                activate_claude_excel_mode()
                claude_writer.start_dictation()
                dictation_state["active"] = True
                dictation_state["target"] = "claude_excel"
                msg = "Dictation started for Excel (via Claude)."
                print(msg)
                speak_feedback(msg)
                return

            dictation_state["active"] = True
            dictation_state["target"] = target
            msg = f"Dictation started for {target}."
            print(msg)
            speak_feedback(msg)
            return

        if intent == "stop_dictation":
            if dictation_state.get("target") == "claude_excel":
                claude_writer.stop_dictation()
            with _form_buffer_lock:
                _clear_overwrite_pending_locked()
            dictation_state["active"] = False
            dictation_state["target"] = None
            msg = "Dictation stopped."
            print(msg)
            speak_feedback(msg)
            return

        if intent == "next_line":
            target = dictation_state.get("target", "notepad")
            if target == "word":
                try:
                    if actions._word_app:
                        actions._word_app.Selection.TypeParagraph()
                except Exception:
                    pass
            else:
                actions.write_in_notepad("", newline=True)
            msg = "➡️ Moved to next line."
            print(msg)
            speak_feedback(msg)
            return

        if intent == "fill_form":
            field = cmd.get("field")
            value = (cmd.get("value") or "").strip()

            if not field:
                speak_feedback("⚠️ Could not detect field name.")
                return

            with _form_buffer_lock:
                CURRENT_FORM_FIELD["name"] = field
                _start_overwrite_pending_locked(field)

                if value:
                    FORM_SUBMISSION_BUFFER[field] = value
                    _clear_overwrite_pending_locked()
                else:
                    FORM_SUBMISSION_BUFFER[field] = ""

                current_value = FORM_SUBMISSION_BUFFER.get(field, "")

            ok = web_form_controller.fill_field(field, current_value)
            if ok:
                if value:
                    speak_feedback(f"Filled {field} with {value}.")
                else:
                    speak_feedback(f"{field} selected.")
            else:
                speak_feedback("⚠️ Browser not connected. Please open the form first.")
            return

        if intent == "submit_form":
            with _form_buffer_lock:
                buffer_copy = dict(FORM_SUBMISSION_BUFFER)

            saved_path = ""
            try:
                saved_path = _save_form_buffer_to_json(buffer_copy)
                if saved_path:
                    print(f"📝 Form saved to {saved_path}")
                else:
                    print("⚠️ Form not saved (error).")
            except Exception as e:
                print("❌ Error while saving form:", e)

            ok = web_form_controller.submit_form()
            if ok:
                with _form_buffer_lock:
                    FORM_SUBMISSION_BUFFER.clear()
                    CURRENT_FORM_FIELD["name"] = None
                    _clear_overwrite_pending_locked()
                if saved_path:
                    speak_feedback(f"Form submitted and saved to {saved_path}.")
                else:
                    speak_feedback("Form submitted (but saving failed).")
            else:
                speak_feedback("⚠️ Browser not connected. Cannot submit the form.")

            if dictation_state.get("target") == "form":
                dictation_state["active"] = False
                dictation_state["target"] = None
            return

        if intent == "save":
            target = cmd.get("target") or dictation_state.get("target")
            if target == "word":
                res = actions.save_in_word()
            elif target == "notepad":
                filename = f"note_{int(time.time())}.txt"
                res = actions.save_in_notepad(filename)
            else:
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
                    "excel": "excel.exe",
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

        if intent == "open":
            app = (cmd.get("app") or raw_transcript or "").strip()
            if not app:
                msg = "⚠️ No app specified to open."
                print(msg)
                speak_feedback(msg)
                return

            if app.lower() == "form":
                web_socket_server.start_in_thread()
                res = actions.open_form()
                print(res)
                speak_feedback(res)
                return

            if "excel" in app.lower():
                focused = actions.focus_claude_window()
                if focused:
                    res = "✅ Focused Claude Desktop for Excel operations."
                else:
                    res = (
                        "⚠️ Could not focus Claude Desktop. Please open Claude manually."
                    )
                dictation_state["target"] = "claude_excel"
            else:
                res = actions.open_app(app)
            print(res)
            speak_feedback(res)
            return

        if intent == "write":
            target = cmd.get("target", "notepad")
            content = cmd.get("content") or raw_transcript or ""
            if not content.strip():
                return

            if target == "word":
                res = actions.write_in_word(content)
            else:
                res = actions.write_in_notepad(content, newline=False)

            print(f"📝 Writing (same line): {content}")
            speak_feedback(res)
            return

        if intent == "write_excel":
            cell = cmd.get("cell")
            content = cmd.get("content") or raw_transcript or ""
            if not cell:
                match = re.search(r"cell\s*([a-zA-Z]+\d+)", raw_transcript.lower())
                if match:
                    cell = match.group(1).upper()

            if not cell:
                msg = "⚠️ Could not detect target Excel cell."
                print(msg)
                speak_feedback(msg)
                return

            phrase = f"Write {content} in cell {cell}"
            claude_writer.send_immediate(phrase)
            res = f"✅ Sent to Claude (immediate): {phrase}"
            print(res)
            speak_feedback(res)
            return

        if intent == "open_and_write_excel":
            content = cmd.get("content") or raw_transcript or ""
            cell = cmd.get("cell")

            focused = actions.focus_claude_window()
            if not focused:
                msg = "⚠️ Could not focus Claude Desktop. Please open Claude manually."
                print(msg)
                speak_feedback(msg)
                return

            if cell:
                phrase = f"Open Excel and write {content} in cell {cell}"
            else:
                phrase = f"Open Excel and write {content}"

            claude_writer.send_immediate(phrase)
            res = f"✅ Sent to Claude (immediate): {phrase}"
            print(res)
            speak_feedback(res)
            return

        if intent == "search_web":
            engine = cmd.get("engine", "google")
            q = cmd.get("query") or raw_transcript or ""
            browser = cmd.get("browser")
            if not q:
                return
            if engine == "wikipedia":
                url = "https://en.wikipedia.org/wiki/" + q.replace(" ", "_")
            elif engine == "youtube":
                url = "https://www.youtube.com/results?search_query=" + q.replace(
                    " ", "+"
                )
            else:
                url = "https://www.google.com/search?q=" + q.replace(" ", "+")
            try:
                if browser:
                    subprocess.Popen(f'start {browser} "{url}"', shell=True)
                else:
                    webbrowser.open(url)
                print(f"✅ Opened web search for: {q}")
                speak_feedback("Opened web search")
            except Exception as e:
                print("❌ Error opening web search:", e)
            return

        if intent == "search":
            name = cmd.get("name") or raw_transcript or ""
            res = actions.search_in_explorer(name)
            print(res)
            speak_feedback(res)
            return

        if intent == "stop":
            msg = "🛑 Stopping assistant by voice command."
            print(msg)
            speak_feedback("Assistant stopped")
            quit_app()
            return

        if intent == "unknown":
            print("⚠️ Unknown command:", raw_transcript)
            return

        print("⚠️ Unhandled intent:", intent)

    except Exception as e:
        print("❌ Error executing command:", e)


async def _deepgram_runner():
    client = DeepgramClient(DEEPGRAM_KEY)
    dg_conn = client.listen.websocket.v("1")

    def on_transcript(self, result, **kwargs):
        if not _listening_flag.is_set():
            return

        try:
            transcript = result.channel.alternatives[0].transcript.strip()
        except Exception:
            return
        if not transcript:
            return

        clean_transcript = transcript.strip()
        print("📝 Transcript:", clean_transcript)

        if dictation_state["active"]:
            target = dictation_state.get("target", "notepad")

            if target == "form":
                cmd = parse_command(clean_transcript)

                if not (
                    cmd and isinstance(cmd, dict) and cmd.get("intent") == "fill_form"
                ):
                    m = re.match(
                        r"^\s*(first name|firstname|given name|last name|lastname|surname|family name|name|email|phone|mobile|address)\s*[,;:\-]?\s*(.*)$",
                        clean_transcript,
                        flags=re.I,
                    )
                    if m:
                        field_raw = m.group(1).lower()
                        value = (m.group(2) or "").strip()

                        field_map = {
                            "first name": "first_name",
                            "firstname": "first_name",
                            "given name": "first_name",
                            "last name": "last_name",
                            "lastname": "last_name",
                            "surname": "surname",
                            "family name": "last_name",
                            "name": "name",
                            "email": "email",
                            "phone": "phone",
                            "mobile": "phone",
                            "address": "address",
                        }
                        field = field_map.get(field_raw, field_raw.replace(" ", "_"))

                        inferred_cmd = {
                            "intent": "fill_form",
                            "field": field,
                            "value": value,
                        }
                        print("🔍 Regex fallback detected fill_form:", inferred_cmd)
                        _execute_command_safe(inferred_cmd, clean_transcript)
                        return

                if cmd and cmd.get("intent") == "fill_form":
                    _execute_command_safe(cmd, clean_transcript)
                    return

                if cmd and cmd.get("intent") in ("submit_form", "stop_dictation"):
                    threading.Thread(
                        target=_execute_command_safe,
                        args=(cmd, clean_transcript),
                        daemon=True,
                    ).start()
                    return

                with _form_buffer_lock:
                    active_field = CURRENT_FORM_FIELD.get("name")
                    overwrite_pending = FORM_OVERWRITE_PENDING.get("active", False)

                if active_field:
                    value = clean_transcript.strip()

                    with _form_buffer_lock:
                        if overwrite_pending:
                            FORM_SUBMISSION_BUFFER[active_field] = value
                            _clear_overwrite_pending_locked()
                        else:
                            prev = FORM_SUBMISSION_BUFFER.get(active_field, "")
                            FORM_SUBMISSION_BUFFER[active_field] = (
                                prev + " " + value if prev else value
                            )
                        final_value = FORM_SUBMISSION_BUFFER[active_field]

                    web_form_controller.fill_field(active_field, final_value)
                    speak_feedback(f"{active_field} updated.")
                else:
                    speak_feedback("⚠️ Please say the field name first.")
                return

            ctrl = parse_command(clean_transcript)
            if ctrl and ctrl.get("intent") == "stop_dictation":
                threading.Thread(
                    target=_execute_command_safe,
                    args=(ctrl, clean_transcript),
                    daemon=True,
                ).start()
                return

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
            elif target == "claude_excel":
                threading.Thread(
                    target=_run_excel_write, args=(clean_transcript,), daemon=True
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

        cmd = parse_command(clean_transcript)
        if not cmd or not isinstance(cmd, dict):
            print("No valid Command (parse_command returned None or invalid).")
            return

        if cmd.get("intent") and cmd.get("intent") != "unknown":
            threading.Thread(
                target=_execute_command_safe, args=(cmd, clean_transcript), daemon=True
            ).start()
            _fragment_buf["text"] = None
            _fragment_buf["time"] = 0.0
            return

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

        _fragment_buf["text"] = clean_transcript
        _fragment_buf["time"] = now
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


def _run_loop(loop):
    asyncio.set_event_loop(loop)
    loop.run_until_complete(_deepgram_runner())
    loop.close()


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
    web_socket_server.start_in_thread()

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
