import os
import subprocess
import time
from typing import Optional
import psutil
import pyautogui
import webbrowser
from pathlib import Path

try:
    import pygetwindow as gw
except Exception:
    gw = None

try:
    import win32com.client as win32
except Exception:
    win32 = None

try:
    import web_socket_server
    import web_form_controller
except Exception:
    web_socket_server = None
    web_form_controller = None
_FORM_HTML = r"E:\Projects\Desktop-automation\desktop-main\form.html"


def is_notepad_running() -> bool:
    for p in psutil.process_iter(["name"]):
        name = p.info.get("name") or ""
        if "notepad" in name.lower():
            return True
    return False


def _focus_window_by_title_contains(part: str, wait: float = 0.15) -> bool:
    try:
        if not gw:
            return False
        wins = gw.getWindowsWithTitle(part)
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
            time.sleep(wait)
            return True
    except Exception:
        pass
    return False


def write_in_notepad(text: str, newline: bool = False) -> str:
    if not text:
        return "⚠️ Nothing to write."

    if not is_notepad_running():
        subprocess.Popen("start notepad", shell=True)
        time.sleep(0.9)

    _focus_window_by_title_contains("notepad", wait=0.15)

    try:
        if newline:
            pyautogui.typewrite(text + "\n", interval=0.03)
        else:
            pyautogui.typewrite(text + " ", interval=0.03)
        return "✅ Appended text to Notepad."
    except Exception as e:
        return f"❌ Error writing in Notepad: {e}"


def save_in_notepad(filename: Optional[str] = None) -> str:
    try:
        _focus_window_by_title_contains("notepad", wait=0.12)
        pyautogui.hotkey("ctrl", "s")
        time.sleep(0.4)
        if filename:
            pyautogui.typewrite(filename)
            time.sleep(0.12)
            pyautogui.press("enter")
            time.sleep(0.3)
        return "✅ Notepad save attempted."
    except Exception as e:
        return f"❌ Error saving Notepad: {e}"


def close_notepad() -> str:
    try:
        _focus_window_by_title_contains("notepad", wait=0.08)
        pyautogui.hotkey("alt", "f4")
        time.sleep(0.2)
        return "✅ Close Notepad attempted."
    except Exception:
        try:
            subprocess.run("taskkill /im notepad.exe /f", shell=True, check=False)
            return "✅ Notepad process terminated."
        except Exception as e:
            return f"❌ Error closing Notepad: {e}"


_word_app = None
_word_doc = None


def write_in_word(text: str) -> str:
    global _word_app, _word_doc
    if not text:
        return "⚠️ Nothing to write."
    if win32 is None:
        return "❌ pywin32 not installed; Word automation unavailable."
    try:
        if _word_app is None:
            _word_app = win32.Dispatch("Word.Application")
            _word_app.Visible = True
            _word_doc = _word_app.Documents.Add()
            time.sleep(0.4)
        selection = _word_app.Selection
        selection.TypeText(text + "\n")
        return "✅ Appended text to Word."
    except Exception as e:
        return f"❌ Error writing in Word: {e}"


def save_in_word(filename: Optional[str] = None) -> str:
    global _word_doc
    if win32 is None:
        return "❌ pywin32 not installed; cannot save Word document."
    if _word_doc is None:
        return "⚠️ No Word document open."
    try:
        if filename:
            docpath = os.path.join(os.path.expanduser("~"), "Documents", filename)
            _word_doc.SaveAs(docpath)
            return f"✅ Word saved as {docpath}"
        else:
            _word_doc.Save()
            return "✅ Word document saved."
    except Exception as e:
        return f"❌ Error saving Word: {e}"


def close_word() -> str:
    global _word_app, _word_doc
    if win32 is None:
        return "❌ pywin32 not installed; cannot close Word."
    try:
        if _word_doc:
            _word_doc.Close(False)
            _word_doc = None
        if _word_app:
            _word_app.Quit()
            _word_app = None
        return "✅ Closed Word application."
    except Exception as e:
        return f"❌ Error closing Word: {e}"


def _ensure_ws_server_started() -> None:
    """
    Start the local WebSocket server thread if available.
    Safe to call multiple times.
    """
    if web_socket_server is None:
        return
    try:
        web_socket_server.start_in_thread()
    except Exception:
        pass


def open_form() -> str:
    """
    Open the local HTML form in the default browser and ensure WebSocket server is running.
    Returns a status string similar to other actions helpers.
    """
    _ensure_ws_server_started()

    form_path = Path(_FORM_HTML)
    if not form_path.exists():
        return f"❌ form.html not found at {_FORM_HTML}"

    try:
        try:
            os.startfile(str(form_path))
        except Exception:
            webbrowser.open(f"file://{str(form_path)}")
        time.sleep(0.6)
        return f"✅ Opened form at {str(form_path)}"
    except Exception as e:
        return f"❌ Could not open form: {e}"


def close_form() -> str:
    """
    Best-effort: try to find an open browser window containing the form title and close it.
    Non-destructive: if it fails, returns a descriptive error string.
    """
    title_substr = "Voice Controlled Form"
    try:
        if gw:
            wins = gw.getWindowsWithTitle(title_substr)
            if wins:
                for w in wins:
                    try:
                        w.close()
                    except Exception:
                        try:
                            w.minimize()
                            w.close()
                        except Exception:
                            pass
                return "✅ Closed form window(s)."
            else:
                return "⚠️ No form window found to close."
        else:
            return "⚠️ pygetwindow not available; cannot close form window programmatically."
    except Exception as e:
        return f"❌ Error closing form: {e}"


def send_form_field(field_id: str, value: str) -> str:
    """
    Convenience wrapper to send a fill command via web_form_controller if available.
    Returns status string.
    """
    if web_form_controller is None:
        return "❌ web_form_controller not available."
    try:
        ok = web_form_controller.fill_field(field_id, value)
        if ok:
            return f"✅ Sent {field_id} = {value} to browser."
        else:
            return "⚠️ Browser is not connected to WebSocket."
    except Exception as e:
        return f"❌ Error sending field: {e}"


def submit_form() -> str:
    """
    Convenience wrapper to send a submit command via web_form_controller if available.
    """
    if web_form_controller is None:
        return "❌ web_form_controller not available."
    try:
        ok = web_form_controller.submit_form()
        if ok:
            return "✅ Submitted form (sent submit command)."
        else:
            return "⚠️ Browser is not connected to WebSocket."
    except Exception as e:
        return f"❌ Error sending submit command: {e}"


def open_app(name: str) -> str:
    if not name:
        return "⚠️ No app name provided."

    name = name.strip()
    lower = name.lower()

    try:
        # URLs
        if lower.startswith("http://") or lower.startswith("https://"):
            webbrowser.open(name)
            return f"✅ Opening URL: {name}"

        # Web shortcuts
        web_apps = {
            "youtube": "https://www.youtube.com",
            "gmail": "https://mail.google.com",
            "google": "https://www.google.com",
            "whatsapp": "https://web.whatsapp.com",
            "facebook": "https://www.facebook.com",
            "instagram": "https://www.instagram.com",
            "twitter": "https://twitter.com",
            "linkedin": "https://www.linkedin.com",
            "reddit": "https://www.reddit.com",
        }

        for key, url in web_apps.items():
            if key in lower:
                webbrowser.open(url)
                return f"✅ Opening {key.capitalize()} in browser."

        if "chrome" in lower:
            subprocess.Popen("start chrome", shell=True)
            return "✅ Opening Chrome."

        if "brave" in lower:
            subprocess.Popen("start brave", shell=True)
            return "✅ Opening Brave."

        if "notepad" in lower:
            subprocess.Popen("start notepad", shell=True)
            return "✅ Opening Notepad."

        if "word" in lower or "winword" in lower:
            subprocess.Popen("start winword", shell=True)
            return "✅ Opening Word."

        if "calculator" in lower or "calc" in lower:
            subprocess.Popen("start calc", shell=True)
            return "✅ Opening Calculator."

        if (
            "visual studio code" in lower
            or "vs code" in lower
            or "vscode" in lower
            or "visual studio" in lower
        ):
            subprocess.Popen("start code", shell=True)
            return "✅ Opening VSCode."
        subprocess.Popen(f"start {name}", shell=True)
        return f"✅ Attempting to open {name}"

    except Exception as e:
        return f"❌ Error opening '{name}': {e}"


def open_in_explorer_with_search(name: str, location: Optional[str] = None) -> str:
    if not name:
        return "⚠️ No name provided."

    try:
        # Open Explorer at location (default: user home)
        if location and os.path.exists(location):
            subprocess.Popen(f'explorer "{location}"', shell=True)
        else:
            subprocess.Popen("explorer", shell=True)

        time.sleep(1.5)

        if _focus_window_by_title_contains(
            "File Explorer", wait=0.2
        ) or _focus_window_by_title_contains("explorer", wait=0.2):
            pyautogui.hotkey("ctrl", "f")
            time.sleep(0.2)

            pyautogui.typewrite(name, interval=0.05)
            time.sleep(0.5)

            pyautogui.press("enter")

            return f"✅ Searching for '{name}' in File Explorer."
        else:
            return "❌ Could not focus File Explorer window."

    except Exception as e:
        return f"❌ Error searching in File Explorer: {e}"


def search_in_explorer(name: str) -> str:
    return open_in_explorer_with_search(name, None)


def close_file_explorer() -> str:
    """
    Closes all File Explorer windows, then restarts Explorer so desktop/taskbar return.
    """
    try:
        subprocess.run(
            ["taskkill", "/f", "/im", "explorer.exe"], shell=True, check=False
        )
        time.sleep(1)
        subprocess.Popen("explorer.exe", shell=True)
        return "✅ Restarted File Explorer."
    except Exception as e:
        return f"❌ Error restarting File Explorer: {e}"


def focus_claude_window(wait: float = 0.15) -> bool:
    """
    Try to focus a Claude (Anthropic) desktop window by checking common title substrings.
    Returns True if a window was focused.
    """
    if not gw:
        return False
    candidates = ["claude", "anthropic", "claude desktop", "claude -"]
    try:
        for title in gw.getAllTitles():
            if not title:
                continue
            lower = title.lower()
            if any(c in lower for c in candidates):
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
                    time.sleep(wait)
                    return True
    except Exception:
        pass
    return False


def is_claude_running() -> bool:
    try:
        for p in psutil.process_iter(["name"]):
            name = p.info.get("name") or ""
            if "claude" in name.lower() or "anthropic" in name.lower():
                return True
    except Exception:
        pass
    return False
