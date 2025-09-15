import os
import pyaudio
import asyncio
import subprocess
from deepgram import DeepgramClient, LiveTranscriptionEvents, LiveOptions
from dotenv import load_dotenv
from llm import text_to_command

load_dotenv()
DEEPGRAM_API_KEY = os.getenv("DEEPGRAM_API_KEY")
if not DEEPGRAM_API_KEY:
    raise ValueError("⚠️ Deepgram API key missing! Add it to your .env file.")

RATE = 16000
CHUNK = 1024
FORMAT = pyaudio.paInt16
CHANNELS = 1

p = pyaudio.PyAudio()
should_stop = asyncio.Event()


async def transcribe_live():
    global should_stop

    deepgram = DeepgramClient(DEEPGRAM_API_KEY)
    dg_connection = deepgram.listen.websocket.v("1")

    def on_transcript(self, result, **kwargs):
        transcript = result.channel.alternatives[0].transcript.strip()

        if transcript:
            print("📝 Transcript:", transcript)

            command = text_to_command(transcript)
            print("Suggested command:", command)

            try:
                subprocess.run(command, shell=True)
            except Exception as e:
                print("Error running command:", e)

        if "stop" in transcript.lower():
            print(" 'stop' detected. Ending transcription...")
            should_stop.set()

    dg_connection.on(LiveTranscriptionEvents.Transcript, on_transcript)

    options = LiveOptions(
        model="nova-2",
        punctuate=True,
        interim_results=False,
        encoding="linear16",
        sample_rate=RATE,
    )

    if not dg_connection.start(options):
        raise RuntimeError("❌ Failed to start Deepgram connection")

    stream = p.open(
        format=FORMAT, channels=CHANNELS, rate=RATE, input=True, frames_per_buffer=CHUNK
    )

    print("🎤 Speak into your mic. Say 'stop' to end.")

    try:
        while not should_stop.is_set():
            data = stream.read(CHUNK, exception_on_overflow=False)
            dg_connection.send(data)
    finally:
        dg_connection.finish()
        stream.stop_stream()
        stream.close()
        p.terminate()


asyncio.run(transcribe_live())
