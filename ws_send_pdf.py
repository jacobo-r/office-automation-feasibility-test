#!/usr/bin/env python3
"""
ws_send_pdf.py
Send a PDF file to a WebSocket server using the `websockets` library (asyncio).
Usage:
  python ws_send_pdf.py ws://127.0.0.1:9000 path/to/file.pdf
"""

import sys
import os
import asyncio
import websockets
from datetime import datetime

LOG_PATH = os.path.join(os.path.dirname(__file__), "ws_send_log.txt")

def log(msg: str):
    print(msg)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {msg}\n")

async def send_file(ws_url: str, pdf_path: str):
    if not os.path.exists(pdf_path):
        log(f"ERROR: File not found -> {pdf_path}")
        return

    try:
        log(f"Connecting to {ws_url} ...")
        async with websockets.connect(ws_url, max_size=None) as ws:
            log(f"Connected. Sending file: {pdf_path}")
            with open(pdf_path, "rb") as f:
                data = f.read()
            await ws.send(data)
            log(f"Sent {len(data)} bytes")

            # Optionally wait for confirmation (depends on your server)
            try:
                reply = await asyncio.wait_for(ws.recv(), timeout=3)
                log(f"Server replied: {reply}")
            except asyncio.TimeoutError:
                log("No reply from server (timeout)")

        log("WebSocket connection closed normally.")

    except Exception as e:
        log(f"ERROR sending file: {e}")

async def main():
    if len(sys.argv) < 3:
        log("Usage: python ws_send_pdf.py <ws_url> <pdf_path>")
        return
    ws_url = sys.argv[1]
    pdf_path = sys.argv[2]
    await send_file(ws_url, pdf_path)

if __name__ == "__main__":
    asyncio.run(main())
