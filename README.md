# UNO Export & Send – Feasibility Test

## How to Run (Quick Instructions)

1. **Make sure OpenOffice (or LibreOffice) on the Windows test machine is running with the UNO listener enabled:**

   ```cmd
   REM If OpenOffice:
   "C:\Program Files (x86)\OpenOffice 4\program\soffice.exe" --accept="socket,host=127.0.0.1,port=2002;urp;" --norestore

   REM Or if LibreOffice:
   "C:\Program Files\LibreOffice\program\soffice.exe" --accept="socket,host=127.0.0.1,port=2002;urp;" --norestore

	2.	Use the Office-provided Python binary (recommended) so pyuno is available. Example path:

"C:\Program Files (x86)\OpenOffice 4\program\python.exe" -m pip install requests websocket-client

If you use LibreOffice, change the path accordingly.

	3.	Save the script below as uno_export_send.py.
	4.	Run it with either a WebSocket URL or an HTTPS upload URL:

"C:\Program Files (x86)\OpenOffice 4\program\python.exe" uno_export_send.py --ws-url ws://127.0.0.1:9000/upload

or

"C:\Program Files (x86)\OpenOffice 4\program\python.exe" uno_export_send.py --upload-url https://your-backend.example.com/upload --token YOUR_API_TOKEN

If both --ws-url and --upload-url are provided, the script will prefer WebSockets (binary send) first.

⸻

The Script – uno_export_send.py

Save the following file and run as above:

#!/usr/bin/env python3
"""
uno_export_send.py
Simple feasibility test:
 - Connect to UNO at 127.0.0.1:2002
 - Export active document to a temporary PDF
 - Send PDF via WebSocket (binary) or HTTPS (multipart/form-data)

Usage:
  python uno_export_send.py --ws-url ws://127.0.0.1:9000/upload
  python uno_export_send.py --upload-url https://example.com/upload --token ABC

Note:
  Run with your Office Python (OpenOffice/LibreOffice python) so pyuno is available.
"""

(Insert the full script contents here.)

⸻

Notes, Tips, and Next Steps
	•	Run this with the Office Python (the one inside .../program/python.exe) — that guarantees pyuno is available.
	•	Install dependencies:
For WebSocket:

"C:\Program Files (x86)\OpenOffice 4\program\python.exe" -m pip install websocket-client

For HTTPS:

"C:\Program Files (x86)\OpenOffice 4\program\python.exe" -m pip install requests


	•	If you prefer not to install packages into the Office Python, you can run the script with a system Python, but you’ll need to ensure pyuno is importable — the easiest path is still to use the Office Python.
	•	The script tries to store the document silently.
If the document is untitled, it exports a temporary snapshot.
If the export fails, check that:
	•	The active document is focused, and
	•	OpenOffice was launched with the --accept flag.
	•	For production:
	•	Integrate the exporter into your native app (as a helper executable or imported module).
	•	Add robust retry/queueing.
	•	Pass proper metadata (e.g., filename, user ID) to your backend.

⸻

Optional Next Steps

If you’d like, I can provide:
	•	A tiny Node or Python test server that accepts the WebSocket binary (for local testing), or
	•	A .bat launcher that starts OpenOffice with the --accept flag and runs this script automatically.

⸻
