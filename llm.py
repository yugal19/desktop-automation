import os
from dotenv import load_dotenv
from google import genai

load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")

client = genai.Client(api_key=API_KEY)


def text_to_command(text: str) -> str:
    """
    Converts natural language text into a Windows command using Gemini.
    """
    prompt = f"""
    You are an AI that converts user requests into Windows desktop commands.
    User request: "{text}"
    
    ✅ Rules:
    - Respond ONLY with the exact Windows command (no explanations, no formatting).
    - If the request cannot be mapped to a command, return "echo Unsupported command".
    """

    response = client.models.generate_content(model="gemini-1.5-flash", contents=prompt)

    command = response.text.strip()
    return command
