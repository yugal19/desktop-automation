"""
Claude dictation handler with SILENCE-BASED sending.

Public API:
  - start_dictation()    # begin silence-based accumulation
  - push_text(text)      # append a final transcript chunk to the buffer
  - stop_dictation()     # stop watcher and flush leftover buffer immediately
  - send_immediate(text) # type text and press Enter right away

Configuration:
  - CLAUDE_SILENCE_SECONDS (env) can override the default 4.0 seconds
"""

import os
import time
import threading
from typing import Optional
from pynput.keyboard import Controller, Key

keyboard = Controller()

# config via env (defaults to 4.0 seconds)
try:
    SILENCE_SECONDS = float(os.getenv("CLAUDE_SILENCE_SECONDS", "4.0"))
except Exception:
    SILENCE_SECONDS = 4.0

CHECK_INTERVAL = 0.15  # how often watcher checks for silence (seconds)

# internal state
_buffer = ""  # accumulated text
_last_speech_time = 0.0  # timestamp of last push_text call
_active = False  # dictation mode on/off
_thread: Optional[threading.Thread] = None
_lock = threading.Lock()


def _type_into_claude(msg: str):
    """Type the final message and press Enter (blocking, types character-by-character)."""
    msg = (msg or "").strip()
    if not msg:
        return

    print(f"🟢 Typing into Claude (len {len(msg)}): {msg}")

    # type slowly but reliably
    for ch in msg:
        keyboard.press(ch)
        keyboard.release(ch)
        time.sleep(0.006)

    keyboard.press(Key.enter)
    keyboard.release(Key.enter)


def send_immediate(message: str):
    """Public: type message immediately and press Enter."""
    if not message or not message.strip():
        return
    with _lock:
        # don't interfere with active dictation buffer; send separately
        _type_into_claude(message.strip())


def push_text(text: str):
    """
    Append text (final transcript) to buffer and update last speech time.
    main.py should call this for each final transcript chunk received while dictation is active.
    """
    global _buffer, _last_speech_time

    if not _active:
        # ignore pushes when dictation not active
        return

    if not text:
        return

    text = text.strip()
    if not text:
        return

    with _lock:
        if _buffer:
            _buffer += " " + text
        else:
            _buffer = text
        _last_speech_time = time.time()

    print(f"🎤 Buffer updated ({len(_buffer)} chars): {_buffer}")


def _watcher():
    """Background thread: when silence >= threshold, send accumulated buffer and clear it."""
    global _buffer, _active

    print(f"🔎 Claude watcher started (silence threshold {SILENCE_SECONDS}s).")
    while _active:
        time.sleep(CHECK_INTERVAL)
        with _lock:
            if not _buffer:
                continue
            # if enough silence passed since last push_text, send buffer
            silence = time.time() - _last_speech_time
            if silence >= SILENCE_SECONDS:
                to_send = _buffer.strip()
                _buffer = ""  # clear before sending to avoid races
                print("🤫 Silence detected → sending accumulated text to Claude")
                try:
                    _type_into_claude(to_send)
                except Exception as e:
                    print("❌ Error while typing into Claude:", e)

    print("🔎 Claude watcher stopped.")


def start_dictation():
    """Begin silence-based dictation. Safe to call multiple times."""
    global _active, _last_speech_time, _thread
    if _active:
        return "Already running"
    _active = True
    _last_speech_time = time.time()
    _thread = threading.Thread(target=_watcher, daemon=True)
    _thread.start()
    print("🎧 Claude dictation STARTED (silence-based).")
    return "Started claude dictation"


def stop_dictation():
    """Stop dictation and flush leftover buffer immediately (press Enter)."""
    global _active, _buffer
    if not _active:
        return "Not running"
    _active = False

    # flush leftover buffer if any
    with _lock:
        leftover = _buffer.strip()
        _buffer = ""
    if leftover:
        print("🛑 Stopping dictation → sending leftover buffer to Claude")
        try:
            _type_into_claude(leftover)
        except Exception as e:
            print("❌ Error while flushing leftover buffer:", e)

    # wait briefly for watcher to exit (non-blocking)
    print("⛔ Claude dictation STOPPED.")
    return "Stopped claude dictation"
