# interpreter.py
import os
import json
import re
from typing import Optional
from dotenv import load_dotenv

load_dotenv()
ENABLE_GEMINI = os.getenv("ENABLE_GEMINI", "false").lower() in ("1", "true", "yes")
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

genai_client = None
if ENABLE_GEMINI and GEMINI_API_KEY:
    try:
        from google import genai

        genai_client = genai.Client(api_key=GEMINI_API_KEY)
    except Exception as e:
        print("⚠️ Gemini requested but genai import failed:", e)
        genai_client = None

# Known app keywords to treat as open shortcuts
KNOWN_APPS = [
    "chrome",
    "brave",
    "notepad",
    "word",
    "winword",
    "calculator",
    "calc",
    "vscode",
    "visual studio code",
    "youtube",
    "google",
    "excel",
]


def _clean(text: str) -> str:
    if not text:
        return ""
    # remove annoying punctuation but keep ":" for drives like "D:"
    text = re.sub(r"[^\w\s:]", "", text)
    return text.lower().strip()


CURRENT_APP = None


def parse_command(text: str) -> dict:
    """
    Parse a short desktop voice command into a structured intent dict.

    Returns a dict like:
      {"intent": "write", "target": "notepad", "content": "hello"}
      {"intent": "open", "app": "excel"}
      {"intent": "write_excel", "target":"excel", "content":"10", "cell":"B2"}
      {"intent": "open_and_write_excel", "app":"excel", "content":"hello", "cell":"A1"}
      {"intent": "unknown", "raw": "<original>"}
    """
    global CURRENT_APP
    raw = text or ""
    t = _clean(raw)

    # --- Stop assistant ---
    if any(
        kw in t
        for kw in [
            "stop listening",
            "stop assistant",
            "exit assistant",
            "goodbye assistant",
        ]
    ):
        return {"intent": "stop"}

    # --- Dictation control ---
    if any(
        kw in t
        for kw in [
            "start writing",
            "start dictation",
            "begin dictation",
            "continue writing",
            "continue in",
            "resume writing",
            "resume dictation",
            "start dictation in",
        ]
    ):
        target = CURRENT_APP or "notepad"  # default to last opened app
        if "word" in t:
            target = "word"
        elif "notepad" in t:
            target = "notepad"
        elif "excel" in t:
            target = "excel"
        return {"intent": "start_dictation", "target": target}

    # --- Excel: "open excel and write X" ---
    m = re.match(r"open\s+excel\s+and\s+write\s+(.+)", t)
    if m:
        content_part = m.group(1).strip()
        # Try to detect "in cell A1"
        cell_match = re.search(r"in\s+cell\s+([a-z]+\d+)", content_part)
        if cell_match:
            cell = cell_match.group(1).upper()
            content = re.sub(r"in\s+cell\s+[a-z]+\d+", "", content_part).strip()
        else:
            cell = None
            content = content_part
        CURRENT_APP = "excel"
        return {
            "intent": "open_and_write_excel",
            "app": "excel",
            "content": content,
            "cell": cell,
        }

    # --- Explicit: "open excel" (plain) ---
    m = re.match(r"(?:open|start)\s+(?:microsoft\s+)?excel\b", t)
    if m:
        CURRENT_APP = "excel"
        return {"intent": "open", "app": "excel"}

    # --- Excel: "write X in cell A1" (without open) ---
    m = re.match(
        r"(?:write|type|put|insert)\s+(.+?)\s+(?:in|into)\s+cell\s+([a-z]+\d+)", t
    )
    if m:
        content = m.group(1).strip()
        cell = m.group(2).upper()
        target = "excel"
        CURRENT_APP = "excel"
        return {
            "intent": "write_excel",
            "target": target,
            "content": content,
            "cell": cell,
        }

    # --- Next line command ---
    if any(kw in t for kw in ["next line", "new line", "line break"]):
        return {"intent": "next_line"}

    if any(
        kw in t
        for kw in [
            "stop writing",
            "stop dictation",
            "end dictation",
            "finish writing",
            "stop writing now",
        ]
    ):
        return {"intent": "stop_dictation"}

    # --- Save commands ---
    if (
        "save this file" in t
        or t.startswith("save ")
        or "save file" in t
        or "save note" in t
    ):
        target = None
        if "word" in t:
            target = "word"
        elif "notepad" in t:
            target = "notepad"
        else:
            target = CURRENT_APP  # fallback
        return {"intent": "save", "target": target}

    # --- Close commands ---
    if (
        t.startswith("close ")
        or t.startswith("stop ")
        or t.startswith("exit ")
        or "close " in t
        or "stop " in t
        or "exit " in t
    ):
        target = None
        m = re.match(r"(?:close|stop|exit)\s+(.+)", t)
        if m:
            target = m.group(1).strip()
        if "word" in t:
            target = "word"
        elif "notepad" in t:
            target = "notepad"
        elif not target:
            target = CURRENT_APP  # fallback
        return {"intent": "close", "target": target}

    # --- Open X and write Y (generic, non-excel) ---
    m = re.match(r"open\s+([a-z0-9 ]+)\s+and\s+write\s+(.+)", t)
    if m:
        app = m.group(1).strip()
        content = m.group(2).strip()
        CURRENT_APP = app
        return {"intent": "open_and_write", "app": app, "content": content}

    # --- Write commands (generic) ---
    if t.startswith("write ") or t.startswith("type "):
        if t.startswith("write "):
            content = t[len("write ") :].strip()
        else:
            content = t[len("type ") :].strip()
        target = CURRENT_APP or "notepad"  # default to last app
        if " in word" in content:
            content = content.replace(" in word", "").strip()
            target = "word"
        if " in notepad" in content:
            content = content.replace(" in notepad", "").strip()
            target = "notepad"
        # If user said "write X in cell B2" this will have been matched earlier and returned.
        return {"intent": "write", "target": target, "content": content}

    # --- Web search handling ---
    if "search " in t or "google" in t or "wikipedia" in t or "youtube" in t:
        browser = None
        if "in brave" in t or "using brave" in t:
            browser = "brave"
        elif "in chrome" in t or "using chrome" in t:
            browser = "chrome"

        # Wikipedia
        if "wikipedia" in t:
            q = (
                t.replace("search", "")
                .replace("on wikipedia", "")
                .replace("in wikipedia", "")
                .replace("wikipedia", "")
                .strip()
            )
            if q:
                return {"intent": "search_web", "engine": "wikipedia", "query": q}

        # YouTube
        if "youtube" in t:
            q = (
                t.replace("search", "")
                .replace("on youtube", "")
                .replace("in youtube", "")
                .replace("youtube", "")
                .strip()
            )
            if q:
                return {"intent": "search_web", "engine": "youtube", "query": q}

        # Google / generic
        q = t
        if q.startswith("search "):
            q = q[len("search ") :].strip()
        q = (
            q.replace(" on google", "")
            .replace(" in google", "")
            .replace(" google", "")
            .replace("googling", "")
            .strip()
        )
        if q:
            return {
                "intent": "search_web",
                "engine": "google",
                "query": q,
                "browser": browser,
            }

    # --- File Explorer handling ---
    if "file explorer" in t or "windows explorer" in t or t.strip() in ["explorer"]:
        # Case 1: "open photos in file explorer"
        m = re.match(r"(?:open|show)\s+(.+)\s+in\s+(?:file|windows)?\s*explorer", t)
        if m:
            name = m.group(1).strip()
            return {"intent": "search", "name": name, "target": "explorer"}

        # Case 2: "in file explorer open photos"
        m = re.match(r"in\s+(?:file|windows)?\s*explorer\s+(?:open|show)\s+(.+)", t)
        if m:
            name = m.group(1).strip()
            return {"intent": "search", "name": name, "target": "explorer"}

        # Case 3: just "file explorer" / "explorer"
        return {"intent": "search", "name": "", "target": "explorer"}

    # --- File/folder search (generic) ---
    if t.startswith("search ") or t.startswith("show "):
        if t.startswith("search "):
            name = t[len("search ") :].strip()
        else:
            name = t[len("show ") :].strip()
        return {"intent": "search", "name": name}

    # --- Open app shortcuts fallback ---
    for app in KNOWN_APPS:
        # match "open chrome" or just "chrome"
        if re.search(r"(?:open|start)\s+" + re.escape(app) + r"\b", t) or re.search(
            r"\b" + re.escape(app) + r"\b", t
        ):
            if app == "google":
                return {"intent": "search_web", "engine": "google", "query": ""}
            CURRENT_APP = app  # remember app
            return {"intent": "open", "app": app}

    # --- Gemini fallback (optional) ---
    if ENABLE_GEMINI and genai_client:
        try:
            prompt = f"""
You are an assistant that converts short desktop voice commands into a JSON object with an intent.
Allowed intents: write, start_dictation, stop_dictation, open, open_and_write, search, search_web, save, close, stop, unknown, write_excel, open_and_write_excel.
User: "{raw}"
Output JSON only.
"""
            resp = genai_client.models.generate_content(
                model="gemini-2.5-flash-lite", contents=prompt, timeout=6
            )
            raw_text = getattr(resp, "text", None) or str(resp)
            parsed = json.loads(raw_text.strip())
            if isinstance(parsed, dict) and parsed.get("intent"):
                return parsed
        except Exception as e:
            print("⚠️ Gemini fallback failed:", e)

    # --- Fallback ---
    return {"intent": "unknown", "raw": raw}
