# interpreter.py (patched)
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
    "form",
]

# Keeps track of the last field the user was filling in the form mode.
LAST_FORM_FIELD: Optional[str] = None

# Stores the current accumulated values per field while filling via voice
FORM_FIELD_VALUES: dict = (
    {}
)  # e.g. {"first_name":"John Doe", "address":"Flat 101, ..."}


def _clean(text: str) -> str:
    """Lowercase and remove punctuation except ':' (for drive letters)."""
    if not text:
        return ""
    text = re.sub(r"[^\w\s:]", "", text)
    return text.lower().strip()


def _spoken_to_email(spoken: str) -> str:
    """
    Convert a spoken email phrase into an email-like string.
    Handles: at, dot, period, point, underscore, dash/hyphen, space, dotcom, dotin, etc.
    """
    if not spoken:
        return ""
    s = spoken.lower().strip()
    s = re.sub(r"\b(at|@)\b", "@", s)
    s = re.sub(r"\b(dot|point|period)\b", ".", s)
    s = re.sub(r"\b(underscore|under score)\b", "_", s)
    s = re.sub(r"\b(dash|hyphen)\b", "-", s)
    s = re.sub(r"\b(space)\b", "", s)
    s = re.sub(r"\bdotcom\b", ".com", s)
    s = re.sub(r"\bdotin\b", ".in", s)
    s = re.sub(r"\bdotorg\b", ".org", s)
    s = re.sub(r"\bdotco\b", ".co", s)
    s = re.sub(r"\s*@\s*", "@", s)
    s = re.sub(r"\s*\.\s*", ".", s)
    s = re.sub(r"\s+", "", s)
    s = s.strip(".@")
    return s


CURRENT_APP = None


def parse_command(text: str) -> dict:
    """
    Parse voice command -> intent.
    Special for form mode: supports multi-part address/email and continues appending
    to LAST_FORM_FIELD until user changes field or stops dictation.

    IMPORTANT: Mentioning a field name ALWAYS resets/overwrites that field.
    Continuation (append) happens only when the user speaks without a field name
    and LAST_FORM_FIELD is set (form continuation).
    """
    global CURRENT_APP, LAST_FORM_FIELD, FORM_FIELD_VALUES

    raw = text or ""
    t = _clean(raw)

    # STOP / QUIT
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

    # OPEN FORM
    if re.search(r"\bopen\b.*\bform\b", t) or t in (
        "open form",
        "open the form",
        "start form",
    ):
        CURRENT_APP = "form"
        return {"intent": "open", "app": "form"}

    # START/CONTINUE DICTATION
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
        target = CURRENT_APP or "notepad"
        if "word" in t:
            target = "word"
        elif "notepad" in t:
            target = "notepad"
        elif "excel" in t:
            target = "excel"
        if CURRENT_APP == "form" or "form" in t:
            target = "form"
        return {"intent": "start_dictation", "target": target}

    # FORM MODE: raw email with @ (preserve)
    email_match_raw = re.search(r"\bemail\s+([A-Za-z0-9._%+\-@]+)\b", raw, flags=re.I)
    if email_match_raw:
        email_val = email_match_raw.group(1).strip()
        if "@" in email_val:
            LAST_FORM_FIELD = "email"
            FORM_FIELD_VALUES["email"] = email_val
            return {"intent": "fill_form", "field": "email", "value": email_val}

    # FORM MODE: spoken email patterns
    m = re.match(r"^(?:email|enter email|fill email|my email is)\s+(.+)$", t)
    if m:
        spoken = m.group(1).strip()
        email = _spoken_to_email(spoken)
        LAST_FORM_FIELD = "email"
        FORM_FIELD_VALUES["email"] = email
        return {"intent": "fill_form", "field": "email", "value": email}

    m = re.match(r"^(.+)\s+in\s+email$", t)
    if m:
        spoken = m.group(1).strip()
        email = _spoken_to_email(spoken)
        LAST_FORM_FIELD = "email"
        FORM_FIELD_VALUES["email"] = email
        return {"intent": "fill_form", "field": "email", "value": email}

    if t.strip() in ("email", "enter email", "start email"):
        LAST_FORM_FIELD = "email"
        FORM_FIELD_VALUES.setdefault("email", "")
        return {
            "intent": "fill_form",
            "field": "email",
            "value": FORM_FIELD_VALUES["email"],
        }

    # FORM FIELD FILLING: first_name, surname, address
    # NOTE: When the user explicitly says the field name, we OVERWRITE that field value.
    field_patterns = [
        (r"^(?:first name|firstname|first)\s+(.+)$", "first_name"),
        (r"^(?:surname|last name|lastname|last|family name)\s+(.+)$", "surname"),
        (r"^(?:address|home address|full address|residence)\s+(.+)$", "address"),
    ]

    for patt, fid in field_patterns:
        mm = re.match(patt, t)
        if mm:
            val = mm.group(1).strip()
            # OVERWRITE behaviour: replace previous value entirely
            LAST_FORM_FIELD = fid
            FORM_FIELD_VALUES[fid] = val
            return {"intent": "fill_form", "field": fid, "value": val}

    # "<value> in <field>" and "fill <value> in <field>"
    m = re.match(r"^(.+)\s+in\s+(first name|surname|address|email)$", t)
    if m:
        val = m.group(1).strip()
        fld_spoken = m.group(2)
        field_map = {
            "first name": "first_name",
            "surname": "surname",
            "address": "address",
            "email": "email",
        }
        fid = field_map[fld_spoken]
        # OVERWRITE when field name is explicitly mentioned
        if fid == "email":
            val_norm = _spoken_to_email(val)
            LAST_FORM_FIELD = "email"
            FORM_FIELD_VALUES["email"] = val_norm
            return {"intent": "fill_form", "field": "email", "value": val_norm}
        else:
            LAST_FORM_FIELD = fid
            FORM_FIELD_VALUES[fid] = val
            return {"intent": "fill_form", "field": fid, "value": val}

    m = re.match(r"^fill\s+(.+)\s+in\s+(first name|surname|address|email)$", t)
    if m:
        val = m.group(1).strip()
        fld_spoken = m.group(2)
        field_map = {
            "first name": "first_name",
            "surname": "surname",
            "address": "address",
            "email": "email",
        }
        fid = field_map[fld_spoken]
        # OVERWRITE when field name present
        if fid == "email":
            val_norm = _spoken_to_email(val)
            LAST_FORM_FIELD = "email"
            FORM_FIELD_VALUES["email"] = val_norm
            return {"intent": "fill_form", "field": "email", "value": val_norm}
        else:
            LAST_FORM_FIELD = fid
            FORM_FIELD_VALUES[fid] = val
            return {"intent": "fill_form", "field": fid, "value": val}

    # FORM SUBMIT: clear stored form state here so next session is fresh
    if re.search(r"\b(submit|send)\b.*\bform\b", t) or t in (
        "submit form",
        "submit the form",
        "finish form",
    ):
        # Clear stored form buffer & last-field pointer
        LAST_FORM_FIELD = None
        FORM_FIELD_VALUES.clear()
        return {"intent": "submit_form"}

    # STOP DICTATION
    if any(
        kw in t
        for kw in [
            "stop writing",
            "stop dictation",
            "end dictation",
            "finish writing",
            "stop writing now",
            "end writing",
        ]
    ):
        return {"intent": "stop_dictation"}

    # NEXT LINE
    if any(kw in t for kw in ["next line", "new line", "line break"]):
        return {"intent": "next_line"}

    # SAVE
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
            target = CURRENT_APP
        return {"intent": "save", "target": target}

    # CLOSE
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
            target = CURRENT_APP
        return {"intent": "close", "target": target}

    # OPEN & WRITE (generic)
    m = re.match(r"open\s+([a-z0-9 ]+)\s+and\s+write\s+(.+)", t)
    if m:
        app = m.group(1).strip()
        content = m.group(2).strip()
        CURRENT_APP = app
        return {"intent": "open_and_write", "app": app, "content": content}

    # GENERIC WRITE
    if t.startswith("write ") or t.startswith("type "):
        content = (
            t.replace("write ", "", 1)
            if t.startswith("write ")
            else t.replace("type ", "", 1)
        )
        target = CURRENT_APP or "notepad"
        if " in word" in content:
            content = content.replace(" in word", "").strip()
            target = "word"
        if " in notepad" in content:
            content = content.replace(" in notepad", "").strip()
            target = "notepad"
        return {"intent": "write", "target": target, "content": content}

    # EXCEL cases (kept)
    m = re.match(r"open\s+excel\s+and\s+write\s+(.+)", t)
    if m:
        content_part = m.group(1).strip()
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

    m = re.match(r"(?:open|start)\s+(?:microsoft\s+)?excel\b", t)
    if m:
        CURRENT_APP = "excel"
        return {"intent": "open", "app": "excel"}

    m = re.match(
        r"(?:write|type|put|insert)\s+(.+?)\s+(?:in|into)\s+cell\s+([a-z]+\d+)", t
    )
    if m:
        content = m.group(1).strip()
        cell = m.group(2).upper()
        CURRENT_APP = "excel"
        return {
            "intent": "write_excel",
            "target": "excel",
            "content": content,
            "cell": cell,
        }

    # WEB SEARCH
    if "search " in t or "google" in t or "wikipedia" in t or "youtube" in t:
        browser = None
        if "in brave" in t or "using brave" in t:
            browser = "brave"
        elif "in chrome" in t or "using chrome" in t:
            browser = "chrome"
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

    # EXPLORER / FILE SEARCH
    if "file explorer" in t or "windows explorer" in t or t.strip() in ["explorer"]:
        m = re.match(r"(?:open|show)\s+(.+)\s+in\s+(?:file|windows)?\s*explorer", t)
        if m:
            name = m.group(1).strip()
            return {"intent": "search", "name": name, "target": "explorer"}
        m = re.match(r"in\s+(?:file|windows)?\s*explorer\s+(?:open|show)\s+(.+)", t)
        if m:
            name = m.group(1).strip()
            return {"intent": "search", "name": name, "target": "explorer"}
        return {"intent": "search", "name": "", "target": "explorer"}

    if t.startswith("search ") or t.startswith("show "):
        name = (
            t[len("search ") :].strip()
            if t.startswith("search ")
            else t[len("show ") :].strip()
        )
        return {"intent": "search", "name": name}

    # KNOWN APPS FALLBACK
    for app in KNOWN_APPS:
        if re.search(r"(?:open|start)\s+" + re.escape(app) + r"\b", t) or re.search(
            r"\b" + re.escape(app) + r"\b", t
        ):
            if app == "google":
                return {"intent": "search_web", "engine": "google", "query": ""}
            CURRENT_APP = app
            return {"intent": "open", "app": app}

    # FORM CONTINUATION: if in form mode and last field set, treat raw as continuation
    # This is the only place where we append to an existing field.
    if CURRENT_APP == "form" and LAST_FORM_FIELD:
        incoming = raw.strip()
        if incoming:
            if LAST_FORM_FIELD == "email":
                new_piece = _spoken_to_email(incoming)
                prev = FORM_FIELD_VALUES.get("email", "")
                if prev:
                    # if incoming contains '@' assume correction/replacement
                    merged = (
                        new_piece if "@" in new_piece else (prev + new_piece).strip()
                    )
                else:
                    merged = new_piece
                FORM_FIELD_VALUES["email"] = merged
                LAST_FORM_FIELD = "email"
                return {"intent": "fill_form", "field": "email", "value": merged}
            else:
                prev = FORM_FIELD_VALUES.get(LAST_FORM_FIELD, "")
                new_val = (prev + " " + incoming).strip() if prev else incoming
                FORM_FIELD_VALUES[LAST_FORM_FIELD] = new_val
                return {
                    "intent": "fill_form",
                    "field": LAST_FORM_FIELD,
                    "value": new_val,
                }

    # Optional Gemini fallback omitted for brevity here (same as earlier if enabled)
    if ENABLE_GEMINI and genai_client:
        try:
            prompt = f"""
You are an assistant that converts short desktop voice commands into a JSON object with an intent.
Allowed intents: write, start_dictation, stop_dictation, open, open_and_write, search, search_web, save, close, stop, unknown, write_excel, open_and_write_excel, fill_form, submit_form.
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

    # FALLBACK: unknown
    return {"intent": "unknown", "raw": raw}
