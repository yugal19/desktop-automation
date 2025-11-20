# import os
# import subprocess
# import time
# from typing import Optional
# import psutil
# import pyautogui
# import webbrowser
# import re

# try:
#     import pygetwindow as gw
# except Exception:
#     gw = None

# try:
#     import win32com.client as win32
# except Exception:
#     win32 = None


# # ---------- Notepad helpers ----------
# def is_notepad_running() -> bool:
#     for p in psutil.process_iter(["name"]):
#         name = p.info.get("name") or ""
#         if "notepad" in name.lower():
#             return True
#     return False


# def _focus_window_by_title_contains(part: str, wait: float = 0.15) -> bool:
#     try:
#         if not gw:
#             return False
#         wins = gw.getWindowsWithTitle(part)
#         if wins:
#             w = wins[0]
#             try:
#                 w.restore()
#             except Exception:
#                 pass
#             try:
#                 w.activate()
#             except Exception:
#                 try:
#                     w.minimize()
#                     w.maximize()
#                 except Exception:
#                     pass
#             time.sleep(wait)
#             return True
#     except Exception:
#         pass
#     return False


# def write_in_notepad(text: str, newline: bool = False) -> str:
#     if not text:
#         return "⚠️ Nothing to write."

#     if not is_notepad_running():
#         subprocess.Popen("start notepad", shell=True)
#         time.sleep(0.9)

#     _focus_window_by_title_contains("notepad", wait=0.15)

#     try:
#         if newline:
#             # Add Enter at the end
#             pyautogui.typewrite(text + "\n", interval=0.03)
#         else:
#             # Just type continuously, no Enter
#             pyautogui.typewrite(text + " ", interval=0.03)

#         return "✅ Appended text to Notepad."
#     except Exception as e:
#         return f"❌ Error writing in Notepad: {e}"


# def save_in_notepad(filename: Optional[str] = None) -> str:
#     try:
#         _focus_window_by_title_contains("notepad", wait=0.12)
#         pyautogui.hotkey("ctrl", "s")
#         time.sleep(0.4)
#         if filename:
#             pyautogui.typewrite(filename)
#             time.sleep(0.12)
#             pyautogui.press("enter")
#             time.sleep(0.3)
#         return "✅ Notepad save attempted."
#     except Exception as e:
#         return f"❌ Error saving Notepad: {e}"


# def close_notepad() -> str:
#     try:
#         _focus_window_by_title_contains("notepad", wait=0.08)
#         pyautogui.hotkey("alt", "f4")
#         time.sleep(0.2)
#         return "✅ Close Notepad attempted."
#     except Exception:
#         try:
#             subprocess.run("taskkill /im notepad.exe /f", shell=True, check=False)
#             return "✅ Notepad process terminated."
#         except Exception as e:
#             return f"❌ Error closing Notepad: {e}"


# # ---------- Word helpers ----------
# _word_app = None
# _word_doc = None


# def write_in_word(text: str) -> str:
#     global _word_app, _word_doc
#     if not text:
#         return "⚠️ Nothing to write."
#     if win32 is None:
#         return "❌ pywin32 not installed; Word automation unavailable."
#     try:
#         if _word_app is None:
#             _word_app = win32.Dispatch("Word.Application")
#             _word_app.Visible = True
#             _word_doc = _word_app.Documents.Add()
#             time.sleep(0.4)
#         selection = _word_app.Selection
#         selection.TypeText(text + "\n")
#         return "✅ Appended text to Word."
#     except Exception as e:
#         return f"❌ Error writing in Word: {e}"


# def save_in_word(filename: Optional[str] = None) -> str:
#     global _word_doc
#     if win32 is None:
#         return "❌ pywin32 not installed; cannot save Word document."
#     if _word_doc is None:
#         return "⚠️ No Word document open."
#     try:
#         if filename:
#             docpath = os.path.join(os.path.expanduser("~"), "Documents", filename)
#             _word_doc.SaveAs(docpath)
#             return f"✅ Word saved as {docpath}"
#         else:
#             _word_doc.Save()
#             return "✅ Word document saved."
#     except Exception as e:
#         return f"❌ Error saving Word: {e}"


# def close_word() -> str:
#     global _word_app, _word_doc
#     if win32 is None:
#         return "❌ pywin32 not installed; cannot close Word."
#     try:
#         if _word_doc:
#             _word_doc.Close(False)
#             _word_doc = None
#         if _word_app:
#             _word_app.Quit()
#             _word_app = None
#         return "✅ Closed Word application."
#     except Exception as e:
#         return f"❌ Error closing Word: {e}"


# # ---------- Open generic app ----------
# def open_app(name: str) -> str:
#     if not name:
#         return "⚠️ No app name provided."

#     name = name.strip()
#     lower = name.lower()

#     try:
#         # 🌐 Handle direct URLs
#         if lower.startswith("http://") or lower.startswith("https://"):
#             webbrowser.open(name)
#             return f"✅ Opening URL: {name}"

#         # 🎯 Web shortcuts (auto open in browser)
#         web_apps = {
#             "youtube": "https://www.youtube.com",
#             "gmail": "https://mail.google.com",
#             "google": "https://www.google.com",
#             "whatsapp": "https://web.whatsapp.com",
#             "facebook": "https://www.facebook.com",
#             "instagram": "https://www.instagram.com",
#             "twitter": "https://twitter.com",
#             "linkedin": "https://www.linkedin.com",
#             "reddit": "https://www.reddit.com",
#         }

#         for key, url in web_apps.items():
#             if key in lower:
#                 webbrowser.open(url)
#                 return f"✅ Opening {key.capitalize()} in browser."

#         # 🖥️ Desktop apps mapping
#         if "chrome" in lower:
#             subprocess.Popen("start chrome", shell=True)
#             return "✅ Opening Chrome."

#         if "brave" in lower:
#             subprocess.Popen("start brave", shell=True)
#             return "✅ Opening Brave."

#         if "notepad" in lower:
#             subprocess.Popen("start notepad", shell=True)
#             return "✅ Opening Notepad."

#         if "word" in lower or "winword" in lower:
#             subprocess.Popen("start winword", shell=True)
#             return "✅ Opening Word."

#         if "calculator" in lower or "calc" in lower:
#             subprocess.Popen("start calc", shell=True)
#             return "✅ Opening Calculator."

#         if (
#             "visual studio code" in lower
#             or "vs code" in lower
#             or "vscode" in lower
#             or "visual studio" in lower
#         ):
#             subprocess.Popen("start code", shell=True)
#             return "✅ Opening VSCode."

#         if "excel" in lower or "microsoft excel" in lower:
#             # Use COM to open Excel (keeps _excel_app global and workbook)s
#             try:
#                 subprocess.Popen("start excel", shell=True)
#                 return "✅ Starting Excel (fallback)."
#             except Exception:
#                 return f"❌ Error opening Excel: {e}"

#         # 🖥️ Fallback: try to open by name
#         subprocess.Popen(f"start {name}", shell=True)
#         return f"✅ Attempting to open {name}"

#     except Exception as e:
#         return f"❌ Error opening '{name}': {e}"


# def open_in_explorer_with_search(name: str, location: Optional[str] = None) -> str:
#     if not name:
#         return "⚠️ No name provided."

#     try:
#         # Open Explorer at location (default: user home)
#         if location and os.path.exists(location):
#             subprocess.Popen(f'explorer "{location}"', shell=True)
#         else:
#             subprocess.Popen("explorer", shell=True)

#         time.sleep(1.5)  # wait for Explorer to open

#         # Focus Explorer window
#         if _focus_window_by_title_contains(
#             "File Explorer", wait=0.2
#         ) or _focus_window_by_title_contains("Explorer", wait=0.2):
#             # Trigger search (Ctrl+F)
#             pyautogui.hotkey("ctrl", "f")
#             time.sleep(0.2)

#             # Type the query
#             pyautogui.typewrite(name, interval=0.05)
#             time.sleep(0.5)

#             # Press Enter to run the search
#             pyautogui.press("enter")

#             return f"✅ Searching for '{name}' in File Explorer."
#         else:
#             return "❌ Could not focus File Explorer window."

#     except Exception as e:
#         return f"❌ Error searching in File Explorer: {e}"


# def search_in_explorer(name: str) -> str:
#     return open_in_explorer_with_search(name, None)


# # ---------- Close File Explorer ----------
# def close_file_explorer() -> str:
#     """
#     Closes all File Explorer windows, then restarts Explorer so desktop/taskbar return.
#     """
#     try:
#         subprocess.run(
#             ["taskkill", "/f", "/im", "explorer.exe"], shell=True, check=False
#         )
#         time.sleep(1)
#         subprocess.Popen("explorer.exe", shell=True)
#         return "✅ Restarted File Explorer."
#     except Exception as e:
#         return f"❌ Error restarting File Explorer: {e}"


# _excel_app = None
# _excel_book = None
# _excel_sheet = None


# def _words_to_num(s: str) -> Optional[int]:
#     """
#     Convert an English number phrase (e.g. "fourteen", "one hundred twenty three")
#     into an integer. Returns None if conversion not possible.
#     Supports up to billions (basic).
#     """
#     s = s.lower().replace("-", " ").replace(",", " ")
#     if not s:
#         return None

#     units = {
#         "zero": 0,
#         "one": 1,
#         "two": 2,
#         "three": 3,
#         "four": 4,
#         "five": 5,
#         "six": 6,
#         "seven": 7,
#         "eight": 8,
#         "nine": 9,
#         "ten": 10,
#         "eleven": 11,
#         "twelve": 12,
#         "thirteen": 13,
#         "fourteen": 14,
#         "fifteen": 15,
#         "sixteen": 16,
#         "seventeen": 17,
#         "eighteen": 18,
#         "nineteen": 19,
#     }
#     tens = {
#         "twenty": 20,
#         "thirty": 30,
#         "forty": 40,
#         "fifty": 50,
#         "sixty": 60,
#         "seventy": 70,
#         "eighty": 80,
#         "ninety": 90,
#     }
#     scales = {
#         "hundred": 100,
#         "thousand": 1000,
#         "million": 1_000_000,
#         "billion": 1_000_000_000,
#     }

#     tokens = s.split()
#     total = 0
#     current = 0
#     negative = False

#     if tokens and tokens[0] in ("minus", "negative"):
#         negative = True
#         tokens = tokens[1:]

#     for token in tokens:
#         if token in units:
#             current += units[token]
#         elif token in tens:
#             current += tens[token]
#         elif token == "and":
#             continue
#         elif token in scales:
#             scale = scales[token]
#             if current == 0:
#                 current = 1
#             current *= scale
#             total += current
#             current = 0
#         else:
#             # unknown token -> cannot convert
#             return None

#     value = total + current
#     if negative:
#         value = -value
#     return value


# def _parse_number_from_text(text: str):
#     """
#     Return (value, is_number)
#     - If text indicates a numeric value (explicit digits or `number` prefix or english words)
#       returns (int/float, True).
#     - Else returns (text, False).
#     """
#     if not text:
#         return text, False

#     txt = text.strip().lower()

#     # If user said "number X ..." or "num X ..."
#     m = re.search(r"\bnumber\b\s*(.+)$", txt)
#     if not m:
#         m = re.search(r"\bnum\b\s*(.+)$", txt)
#     candidate = None
#     if m:
#         candidate = m.group(1).strip()
#     else:
#         # also allow raw digits without the "number" keyword (useful if parser already gives "14")
#         candidate = txt

#     # Try direct numeric match first (integers or floats)
#     num_match = re.search(r"^(-?\d+(?:\.\d+)?)$", candidate)
#     if num_match:
#         num_str = num_match.group(1)
#         if "." in num_str:
#             try:
#                 return float(num_str), True
#             except Exception:
#                 pass
#         else:
#             try:
#                 return int(num_str), True
#             except Exception:
#                 pass

#     # Try to extract digits from candidate (e.g., "14 in cell d3" - but parser should already strip cell)
#     extract = re.search(r"(-?\d+(?:\.\d+)?)", candidate)
#     if extract:
#         num_str = extract.group(1)
#         if "." in num_str:
#             try:
#                 return float(num_str), True
#             except Exception:
#                 pass
#         else:
#             try:
#                 return int(num_str), True
#             except Exception:
#                 pass

#     # Try to convert English words -> number (e.g., "fourteen", "one hundred twenty")
#     wn = _words_to_num(candidate)
#     if wn is not None:
#         return wn, True

#     # Not a number
#     return text, False


# _excel_app = None
# _excel_book = None
# _excel_sheet = None


# def _ensure_excel_open():
#     """Ensure Excel is running with an active workbook and sheet."""
#     global _excel_app, _excel_book, _excel_sheet
#     try:
#         if _excel_app is None:
#             try:
#                 _excel_app = win32.GetActiveObject("Excel.Application")
#             except Exception:
#                 _excel_app = win32.Dispatch("Excel.Application")
#                 _excel_app.Visible = True
#             time.sleep(0.2)

#         if _excel_book is None:
#             if _excel_app.Workbooks.Count > 0:
#                 _excel_book = _excel_app.ActiveWorkbook
#             else:
#                 _excel_book = _excel_app.Workbooks.Add()
#             time.sleep(0.1)

#         if _excel_sheet is None:
#             try:
#                 _excel_sheet = _excel_app.ActiveSheet
#             except Exception:
#                 _excel_sheet = _excel_book.Sheets(1)
#     except Exception as e:
#         raise RuntimeError(f"Could not open Excel: {e}")


# def open_excel(file_path=None) -> str:
#     """Open Excel or a specific file."""
#     global _excel_app, _excel_book, _excel_sheet
#     try:
#         if file_path:
#             try:
#                 _excel_app = win32.GetActiveObject("Excel.Application")
#             except Exception:
#                 _excel_app = win32.Dispatch("Excel.Application")
#                 _excel_app.Visible = True
#             _excel_book = _excel_app.Workbooks.Open(file_path)
#             _excel_sheet = _excel_book.ActiveSheet
#             return f"✅ Opened Excel file: {file_path}"
#         else:
#             _ensure_excel_open()
#             return "✅ Excel opened and ready."
#     except Exception as e:
#         return f"❌ Error opening Excel: {e}"


# def continue_writing_excel(text: str) -> str:
#     """
#     Continue writing spoken text into Excel.
#     Example:
#       - "Write Hello in cell B2"
#       - "Write 45 in C3"
#       - "Continue writing this line"
#     """
#     global _excel_app, _excel_book, _excel_sheet
#     try:
#         _ensure_excel_open()
#     except Exception as e:
#         return f"❌ Error opening Excel: {e}"

#     if not text.strip():
#         return "⚠️ Nothing to write."

#     # Detect if a cell is mentioned like "in cell B2" or "in B2"
#     cell_match = re.search(
#         r"\bcell\s*([A-Za-z]+\d+)\b|\bin\s*([A-Za-z]+\d+)\b", text, re.IGNORECASE
#     )
#     if cell_match:
#         cell = (cell_match.group(1) or cell_match.group(2)).upper()
#         # Remove command words like "write", "in cell B2"
#         value = re.sub(r"write|right", "", text, flags=re.IGNORECASE)
#         value = re.sub(
#             r"in\s*cell\s*[A-Za-z]+\d+|in\s*[A-Za-z]+\d+",
#             "",
#             value,
#             flags=re.IGNORECASE,
#         ).strip()
#         try:
#             _excel_sheet.Range(cell).Value = value
#             return f"✅ Wrote '{value}' into {cell}."
#         except Exception as e:
#             return f"❌ Error writing into {cell}: {e}"
#     else:
#         # Append text in next available cell in column A
#         try:
#             last_row = _excel_sheet.Cells(_excel_sheet.Rows.Count, 1).End(-4162).Row
#             if _excel_sheet.Cells(1, 1).Value in (None, "") and last_row == 1:
#                 row = 1
#             else:
#                 if _excel_sheet.Cells(last_row, 1).Value in (None, ""):
#                     row = last_row
#                 else:
#                     row = last_row + 1
#             _excel_sheet.Cells(row, 1).Value = text
#             return f"✅ Wrote '{text}' into A{row}."
#         except Exception as e:
#             return f"❌ Error writing to Excel: {e}"


# ===== EXCEL =====


# _excel_app = None
# _excel_book = None
# _excel_sheet = None


# def _words_to_num(s):
#     units = {
#         "zero": 0,
#         "one": 1,
#         "two": 2,
#         "three": 3,
#         "four": 4,
#         "five": 5,
#         "six": 6,
#         "seven": 7,
#         "eight": 8,
#         "nine": 9,
#         "ten": 10,
#         "eleven": 11,
#         "twelve": 12,
#         "thirteen": 13,
#         "fourteen": 14,
#         "fifteen": 15,
#         "sixteen": 16,
#         "seventeen": 17,
#         "eighteen": 18,
#         "nineteen": 19,
#     }
#     tens = {
#         "twenty": 20,
#         "thirty": 30,
#         "forty": 40,
#         "fifty": 50,
#         "sixty": 60,
#         "seventy": 70,
#         "eighty": 80,
#         "ninety": 90,
#     }
#     s = s.lower().replace("-", " ")
#     parts = s.split()
#     if not parts:
#         return None
#     total = 0
#     current = 0
#     for w in parts:
#         if w in units:
#             current += units[w]
#         elif w in tens:
#             current += tens[w]
#         elif w == "hundred":
#             current *= 100
#         else:
#             return None
#     return total + current


# def _ensure_excel():
#     global _excel_app, _excel_book, _excel_sheet
#     if win32 is None:
#         raise RuntimeError("pywin32 required")
#     if _excel_app is None:
#         try:
#             _excel_app = win32.GetActiveObject("Excel.Application")
#         except:
#             _excel_app = win32.Dispatch("Excel.Application")
#             _excel_app.Visible = True
#     if _excel_book is None:
#         _excel_book = (
#             _excel_app.ActiveWorkbook
#             if _excel_app.Workbooks.Count > 0
#             else _excel_app.Workbooks.Add()
#         )
#     if _excel_sheet is None:
#         _excel_sheet = _excel_app.ActiveSheet


# def open_excel():
#     _ensure_excel()
#     return "✅ Excel ready."


# def continue_writing_excel(text):
#     _ensure_excel()
#     if not text.strip():
#         return "⚠️ Nothing."

#     # row 3 column c fourteen
#     m = re.search(r"row\s*(\d+)\s*column\s*([a-z])\s*(.*)", text, re.IGNORECASE)
#     if m:
#         row = m.group(1).strip()
#         col = m.group(2).upper()
#         val = m.group(3).strip()
#         num = _words_to_num(val)
#         if num is not None:
#             val = num
#         cell = f"{col}{row}"
#         _excel_sheet.Range(cell).Value = val
#         return f"✅ Wrote '{val}' into {cell}."

#     # Fallback append in column A
#     last = _excel_sheet.Cells(_excel_sheet.Rows.Count, 1).End(-4162).Row
#     row = 1 if _excel_sheet.Cells(1, 1).Value in (None, "") and last == 1 else last + 1
#     _excel_sheet.Cells(row, 1).Value = text
#     return f"✅ Wrote '{text}' into A{row}."
# actions.py
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

        # NOTE: we do NOT special-case "excel" here anymore (Excel handled via Claude+MCP)
        # Fallback: try to open by name
        subprocess.Popen(f"start {name}", shell=True)
        return f"✅ Attempting to open {name}"

    except Exception as e:
        return f"❌ Error opening '{name}': {e}"


# ---------- Explorer search/open ----------
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
        ) or _focus_window_by_title_contains("explorer", wait=0.2):
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


# ---------- CLAUDE helpers ----------
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
