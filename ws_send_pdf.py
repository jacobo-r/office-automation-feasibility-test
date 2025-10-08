import sys, os, asyncio, websockets

async def send_pdf(ws_url, pdf_path):
    if not os.path.exists(pdf_path):
        print(f"File not found: {pdf_path}")
        sys.exit(1)
    async with websockets.connect(ws_url, max_size=None) as ws:
        print(f"Connected to {ws_url}")
        with open(pdf_path, "rb") as f:
            await ws.send(f.read())
        print(f"Sent {os.path.basename(pdf_path)}")
        try:
            ack = await asyncio.wait_for(ws.recv(), timeout=5)
            print("Server replied:", ack)
        except asyncio.TimeoutError:
            print("No ACK")

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: ws_send_pdf.py <ws_url> <pdf_path>")
        sys.exit(1)
    ws_url, pdf_path = sys.argv[1], sys.argv[2]
    asyncio.run(send_pdf(ws_url, pdf_path))
