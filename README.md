Windows-ready feasibility script:
	•	Connects to the UNO socket (OpenOffice / LibreOffice) on localhost,
	•	Grabs the currently active document,
	•	Exports it to a temporary PDF using the suite’s own PDF exporter,
	•	Sends the PDF either over WebSocket (binary) or HTTPS (multipart/form-data), depending on the CLI argument you pass.

I kept it intentionally minimal and defensive so you can run it quickly and iterate.

⸻

How to run (quick instructions)
	1.	Make sure OpenOffice (or LibreOffice) on the Windows test machine is running with the UNO listener enabled, e.g. (PowerShell / cmd):

REM If OpenOffice:
"C:\Program Files (x86)\OpenOffice 4\program\soffice.exe" --accept="socket,host=127.0.0.1,port=2002;urp;" --norestore
REM Or if LibreOffice:
"C:\Program Files\LibreOffice\program\soffice.exe" --accept="socket,host=127.0.0.1,port=2002;urp;" --norestore

	2.	Use the Office-provided Python binary (recommended) so pyuno is available. Example path:

"C:\Program Files (x86)\OpenOffice 4\program\python.exe" -m pip install requests websocket-client

(If you use LibreOffice, change the path accordingly.)
	3.	Save the script below as uno_export_send.py.
	4.	Run it with either a WebSocket URL or an HTTPS upload URL:

"C:\Program Files (x86)\OpenOffice 4\program\python.exe" uno_export_send.py --ws-url ws://127.0.0.1:9000/upload

or

"C:\Program Files (x86)\OpenOffice 4\program\python.exe" uno_export_send.py --upload-url https://your-backend.example.com/upload --token YOUR_API_TOKEN

If both --ws-url and --upload-url are provided, the script will prefer websockets (binary send) first.

⸻

The script (uno_export_send.py)

Save the following file and run as above.

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
Note: Run with your office python (OpenOffice/LibreOffice python) so pyuno is available.

Notes, tips and next steps
	•	Run this with the Office python (the one inside .../program/python.exe) — that guarantees pyuno is available.
	•	If you want to use the WebSocket route, install websocket-client into that python with:

"C:\Program Files (x86)\OpenOffice 4\program\python.exe" -m pip install websocket-client

For HTTPS, install requests:

"C:\Program Files (x86)\OpenOffice 4\program\python.exe" -m pip install requests


	•	If you prefer not to install packages into the Office python, you can run the script with a system Python but you’ll need to ensure pyuno is importable — easiest is to use the Office python.
	•	The script tries to store the document silently. If the doc is untitled, it exports a temporary snapshot. If the export fails, check whether the active doc is indeed focused and that OpenOffice was launched with the --accept flag.
	•	For production, you’ll integrate the exporter into your native app (call as a helper executable or import the function), add robust retry/queueing, and pass proper metadata (filename, user id) to your backend.

⸻

If you want, I can:
	•	Give you a tiny Node/Python test server that accepts the WebSocket binary (for local testing), or
	•	Generate a small .bat file that launches OpenOffice with --accept and then runs this script automatically for testing.

Which would you like next?
