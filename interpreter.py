# interpreter.py
import os
import json
import re
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
]


def _clean(text: str) -> str:
    if not text:
        return ""
    # remove annoying punctuation but keep ":" for drives like "D:"
    text = re.sub(r"[^\w\s:]", "", text)
    return text.lower().strip()


import re


def parse_command(text: str) -> dict:
    """
    Returns dict with 'intent' (one of):
      - write (target: notepad|word, content)
      - start_dictation / stop_dictation
      - open (app)
      - open_and_write (app, content)
      - search (file/folder) -> use explorer if target=explorer
      - search_web (query, browser?) -> google/wikipedia/youtube
      - save (target optional)
      - close (target)
      - stop (stop assistant)
      - unknown (raw)
    """
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
        target = "notepad"
        if "word" in t:
            target = "word"
        elif "notepad" in t:
            target = "notepad"
        return {"intent": "start_dictation", "target": target}

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
        return {"intent": "close", "target": target}

    # --- Open X and write Y ---
    m = re.match(r"open\s+([a-z0-9 ]+)\s+and\s+write\s+(.+)", t)
    if m:
        app = m.group(1).strip()
        content = m.group(2).strip()
        return {"intent": "open_and_write", "app": app, "content": content}

    # --- Write commands ---
    if t.startswith("write ") or t.startswith("type "):
        if t.startswith("write "):
            content = t[len("write ") :].strip()
        else:
            content = t[len("type ") :].strip()
        target = "notepad"
        if " in word" in content:
            content = content.replace(" in word", "").strip()
            target = "word"
        if " in notepad" in content:
            content = content.replace(" in notepad", "").strip()
            target = "notepad"
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

    # --- Open app shortcuts ---
    for app in KNOWN_APPS:
        if re.search(r"\b" + re.escape(app) + r"\b", t):
            if app == "google":
                return {"intent": "search_web", "engine": "google", "query": ""}
            return {"intent": "open", "app": app}

        # --- Fallback ---
        return {"intent": "unknown", "raw": raw}

    # Gemini fallback
    if ENABLE_GEMINI and genai_client:
        try:
            prompt = f"""
You are an assistant that converts short desktop voice commands into a JSON object with an intent.
Allowed intents: write, start_dictation, stop_dictation, open, open_and_write, search, search_web, save, close, stop, unknown.
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

    return {"intent": "unknown", "raw": raw}
