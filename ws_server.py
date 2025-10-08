#!/usr/bin/env python3
"""
ws_send_pdf.py
Automatically send all PDF files in temp_pdf folder to a WebSocket server,
then delete each file after successful transmission.

Usage:
  python ws_send_pdf.py ws://127.0.0.1:9000 D:\path\to\temp_pdf
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
    try:
        async with websockets.connect(ws_url, max_size=None) as ws:
            with open(pdf_path, "rb") as f:
                data = f.read()
            await ws.send(data)
            log(f"Sent {os.path.basename(pdf_path)} ({len(data)} bytes)")

            try:
                reply = await asyncio.wait_for(ws.recv(), timeout=3)
                log(f"Server replied: {reply}")
            except asyncio.TimeoutError:
                log("No reply from server (timeout)")

        os.remove(pdf_path)
        log(f"Deleted local file: {pdf_path}")

    except Exception as e:
        log(f"ERROR sending {pdf_path}: {e}")

async def main():
    if len(sys.argv) < 3:
        log("Usage: python ws_send_pdf.py <ws_url> <folder_path>")
        return

    ws_url = sys.argv[1]
    folder = sys.argv[2]

    if not os.path.isdir(folder):
        log(f"ERROR: Folder not found: {folder}")
        return

    pdf_files = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]
    if not pdf_files:
        log(f"No PDFs to send in {folder}")
        return

    log(f"Found {len(pdf_files)} PDFs in {folder}")

    for fname in pdf_files:
        pdf_path = os.path.join(folder, fname)
        await send_file(ws_url, pdf_path)

    log("All pending PDFs processed.")

if __name__ == "__main__":
    asyncio.run(main())
