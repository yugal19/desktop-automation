import asyncio
import json
import threading
import websockets

WS_PORT = 8765

_connected_client = None

_client_lock = threading.Lock()

_server_started = False


async def _handler(websocket):
    global _connected_client
    print("🌐 Browser connected to WebSocket server")

    with _client_lock:
        _connected_client = websocket

    try:
        async for msg in websocket:
            print("📩 Received from browser:", msg)
    except Exception as e:
        print("⚠️ Browser disconnected:", e)
    finally:
        with _client_lock:
            if _connected_client == websocket:
                _connected_client = None
        print("❌ Browser WebSocket closed")


async def _server_loop():
    global _server_loop
    _server_loop = asyncio.get_event_loop()
    print(f"🚀 Starting WebSocket server on ws://localhost:{WS_PORT}")

    async with websockets.serve(_handler, "localhost", WS_PORT):
        await asyncio.Future()  


def start_in_thread():
    """Starts the server only once in a background thread."""
    global _server_started
    if _server_started:
        return

    _server_started = True

    thread = threading.Thread(target=_thread_target, daemon=True)
    thread.start()
    print("🧵 WebSocket server thread started")


def _thread_target():
    try:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        loop.run_until_complete(_server_loop())
    except Exception as e:
        print("❌ WebSocket server error:", e)


def send(data: dict) -> bool:
    global _connected_client, _server_loop

    with _client_lock:
        client = _connected_client

    if not client:
        print("⚠️ No browser connected; cannot send.")
        return False

    if _server_loop is None:
        print("❌ No server loop available.")
        return False

    try:
        asyncio.run_coroutine_threadsafe(client.send(json.dumps(data)), _server_loop)
        return True
    except Exception as e:
        print("❌ Error sending to browser:", e)
        return False
