import os
import subprocess
import time
from typing import Optional
import psutil
import pyautogui
import webbrowser

try:
    import pygetwindow as gw
except Exception:
    gw = None

try:
    import win32com.client as win32
except Exception:
    win32 = None


# ---------- Notepad helpers ----------
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
            # Add Enter at the end
            pyautogui.typewrite(text + "\n", interval=0.03)
        else:
            # Just type continuously, no Enter
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


# ---------- Word helpers ----------
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


# ---------- Open generic app ----------
def open_app(name: str) -> str:
    if not name:
        return "⚠️ No app name provided."

    name = name.strip()
    lower = name.lower()

    try:
        # 🌐 Handle direct URLs
        if lower.startswith("http://") or lower.startswith("https://"):
            webbrowser.open(name)
            return f"✅ Opening URL: {name}"

        # 🎯 Web shortcuts (auto open in browser)
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

        # 🖥️ Desktop apps mapping
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

        # 🖥️ Fallback: try to open by name
        subprocess.Popen(f"start {name}", shell=True)
        return f"✅ Attempting to open {name}"

    except Exception as e:
        return f"❌ Error opening '{name}': {e}"


# ---------- Explorer search/open ----------
# def open_in_explorer(name: str, location: Optional[str] = None) -> str:
#     q = (name or "").strip().lower()
#     if not q:
#         return "⚠️ No name provided."

#     # Prioritize user directories
#     user_home = os.path.expanduser("~")
#     priority_dirs = [
#         os.path.join(user_home, "Desktop"),
#         os.path.join(user_home, "Documents"),
#         os.path.join(user_home, "Downloads"),
#         os.path.join(user_home, "Pictures"),
#         os.path.join(user_home, "Videos"),
#     ]

#     roots = []
#     if location:
#         roots.append(location if location.endswith("\\") else location + "\\")
#     roots.extend(priority_dirs)
#     for drv in ["C:\\", "D:\\", "E:\\"]:
#         if drv not in roots:
#             roots.append(drv)

#     max_checked = 50000
#     checked = 0
#     for root in roots:
#         if not os.path.exists(root):
#             continue
#         for dirpath, dirnames, filenames in os.walk(root):
#             checked += 1
#             if checked > max_checked:
#                 return "⚠️ Search aborted (too many files). Please be more specific."
#             for d in dirnames:
#                 if q in d.lower():
#                     path = os.path.join(dirpath, d)
#                     try:
#                         subprocess.Popen(f'explorer "{path}"')
#                         return f"✅ Opened folder: {path}"
#                     except Exception as e:
#                         return f"❌ Error opening folder: {e}"
#             for f in filenames:
#                 if q in f.lower():
#                     path = os.path.join(dirpath, f)
#                     try:
#                         subprocess.Popen(f'explorer /select,"{path}"')
#                         return f"✅ Opened file location: {path}"
#                     except Exception as e:
#                         return f"❌ Error opening file: {e}"
#     return f"⚠️ No matching file or folder found for '{name}'"


def open_in_explorer_with_search(name: str, location: Optional[str] = None) -> str:
    if not name:
        return "⚠️ No name provided."

    try:
        # Open Explorer at location (default: user home)
        if location and os.path.exists(location):
            subprocess.Popen(f'explorer "{location}"', shell=True)
        else:
            subprocess.Popen("explorer", shell=True)

        time.sleep(1.5)  # wait for Explorer to open

        # Focus Explorer window
        if _focus_window_by_title_contains(
            "File Explorer", wait=0.2
        ) or _focus_window_by_title_contains("Explorer", wait=0.2):
            # Trigger search (Ctrl+F)
            pyautogui.hotkey("ctrl", "f")
            time.sleep(0.2)

            # Type the query
            pyautogui.typewrite(name, interval=0.05)
            time.sleep(0.5)

            # Press Enter to run the search
            pyautogui.press("enter")

            return f"✅ Searching for '{name}' in File Explorer."
        else:
            return "❌ Could not focus File Explorer window."

    except Exception as e:
        return f"❌ Error searching in File Explorer: {e}"


def search_in_explorer(name: str) -> str:
    return open_in_explorer_with_search(name, None)


# ---------- Close File Explorer ----------
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
