import asyncio
import concurrent.futures
import threading
from typing import Callable

try:
    from wecom_aibot_sdk import WSClient as _WSClient
    SMART_BOT_IMPORT_ERROR = None
except Exception as exc:  # pragma: no cover - depends on local environment
    _WSClient = None
    SMART_BOT_IMPORT_ERROR = exc


def smart_bot_sdk_available():
    return _WSClient is not None


def missing_smart_bot_fields(bot_id, secret, chat_id):
    missing = []
    if not (bot_id or "").strip():
        missing.append("Bot ID")
    if not (secret or "").strip():
        missing.append("Secret")
    if not (chat_id or "").strip():
        missing.append("Chat ID")
    return missing


def build_smart_bot_body(text):
    return {
        "msgtype": "markdown",
        "markdown": {"content": text},
    }


class SmartBotSender:
    def __init__(self, client_factory: Callable | None = None):
        self._client_factory = client_factory or _WSClient
        self._loop = None
        self._thread = None
        self._client = None
        self._credentials = None
        self._lock = threading.RLock()

    def close(self):
        with self._lock:
            if self._loop is None:
                return
            loop = self._loop
        try:
            future = asyncio.run_coroutine_threadsafe(self._disconnect_async(), loop)
            future.result(timeout=10)
        except Exception:
            pass
        with self._lock:
            if self._loop is not None:
                self._loop.call_soon_threadsafe(self._loop.stop)
            if self._thread is not None:
                self._thread.join(timeout=2)
            self._loop = None
            self._thread = None

    def send_markdown(self, bot_id, secret, chat_id, text, timeout=20):
        missing = missing_smart_bot_fields(bot_id, secret, chat_id)
        if missing:
            raise ValueError("Missing smart bot fields: " + ", ".join(missing))
        if self._client_factory is None:
            raise RuntimeError(f"wecom-aibot-sdk unavailable: {SMART_BOT_IMPORT_ERROR}")

        loop = self._ensure_loop()
        future = asyncio.run_coroutine_threadsafe(
            self._send_markdown_async(bot_id.strip(), secret.strip(), chat_id.strip(), text),
            loop,
        )
        return future.result(timeout=timeout)

    def _ensure_loop(self):
        with self._lock:
            if self._loop is not None and self._thread is not None and self._thread.is_alive():
                return self._loop

            loop = asyncio.new_event_loop()
            ready = threading.Event()

            def runner():
                asyncio.set_event_loop(loop)
                ready.set()
                loop.run_forever()

            thread = threading.Thread(target=runner, name="SmartBotSenderLoop", daemon=True)
            thread.start()
            ready.wait(timeout=2)
            self._loop = loop
            self._thread = thread
            return loop

    async def _send_markdown_async(self, bot_id, secret, chat_id, text):
        await self._ensure_client_async(bot_id, secret)
        return await self._client.send_message(chat_id, build_smart_bot_body(text))

    async def _ensure_client_async(self, bot_id, secret):
        credentials = (bot_id, secret)
        current = self._client
        connected = bool(current and getattr(current, "is_connected", False))
        if connected and self._credentials == credentials:
            return

        if current is not None:
            await self._disconnect_async()

        self._client = self._client_factory(bot_id, secret)
        await self._client.connect()
        self._credentials = credentials

    async def _disconnect_async(self):
        client = self._client
        self._client = None
        self._credentials = None
        if client is None:
            return
        try:
            await client.disconnect()
        except Exception:
            return
