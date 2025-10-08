import asyncio, websockets, os, datetime

# this piece of code simualtes the websocket server in the backend

SAVE_DIR = "./ws_received_from_macro"  # folder where incoming PDFs will be saved
os.makedirs(SAVE_DIR, exist_ok=True)

async def handle(ws):
    print("Client connected")
    try:
        data = await ws.recv()
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        file_path = os.path.join(SAVE_DIR, f"received_{ts}.pdf")
        with open(file_path, "wb") as f:
            f.write(data)
        print(f"Saved {file_path} ({len(data)} bytes)")
        await ws.send("OK")
    except Exception as e:
        print("Error:", e)
    finally:
        await ws.close()

async def main():
    async with websockets.serve(handle, "127.0.0.1", 9000, max_size=None):
        print("WebSocket server running on ws://127.0.0.1:9000")
        await asyncio.Future()  # run forever

if __name__ == "__main__":
    asyncio.run(main())
