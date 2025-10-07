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
"""

import os
import sys
import tempfile
import argparse
import traceback

# Network libraries (may need to be installed into the office python)
try:
    import requests
except Exception:
    requests = None
try:
    import websocket  # websocket-client
except Exception:
    websocket = None

# UNO imports
try:
    import uno
    from com.sun.star.beans import PropertyValue
except Exception:
    uno = None
    PropertyValue = None


def pv(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def connect_uno(host="127.0.0.1", port=2002, timeout_seconds=5.0):
    if uno is None:
        raise RuntimeError("UNO (pyuno) not available in this Python environment.")
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx
    )
    url = f"uno:socket,host={host},port={port};urp;StarOffice.ComponentContext"
    # This can raise com.sun.star.connection.NoConnectException if not listening
    ctx = resolver.resolve(url)
    return ctx


def export_active_to_pdf(temp_pdf_path):
    """
    Connects to UNO, fetches the active component/document, and exports it to temp_pdf_path.
    Returns the path to the PDF file.
    """
    if uno is None:
        raise RuntimeError("UNO (pyuno) not available in this Python environment.")

    ctx = connect_uno()
    smgr = ctx.ServiceManager
    desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
    doc = desktop.getCurrentComponent()
    if doc is None:
        raise RuntimeError("No active document found (no focused document window).")

    # If the document is modified and has a disk location, save it silently
    try:
        if getattr(doc, "isModified", lambda: False)():
            has_location = getattr(doc, "hasLocation", lambda: False)()
            if has_location:
                # Save inplace
                try:
                    doc.store()
                except Exception:
                    # fallback: ignore if we can't store in place
                    pass
            else:
                # Untitled doc: store a temporary ODF snapshot so metadata exists
                tmp_odt = temp_pdf_path + ".odt"
                doc.storeAsURL(uno.systemPathToFileUrl(tmp_odt), ())
    except Exception:
        # ignore saving errors for the feasibility test
        pass

    # Decide filter based on service support
    filter_name = "writer_pdf_Export"
    try:
        if doc.supportsService("com.sun.star.sheet.SpreadsheetDocument"):
            filter_name = "calc_pdf_Export"
        elif doc.supportsService("com.sun.star.presentation.PresentationDocument"):
            filter_name = "impress_pdf_Export"
        elif doc.supportsService("com.sun.star.drawing.DrawingDocument"):
            filter_name = "draw_pdf_Export"
    except Exception:
        pass

    # Some PDF filter settings (optional)
    filter_data = ()
    try:
        fd = (
            pv("SelectPdfVersion", 1),  # standard PDF
            pv("ReduceImageResolution", True),
            pv("MaxImageResolution", 150),
        )
        filter_data = fd
    except Exception:
        filter_data = ()

    # Export to PDF
    pdf_url = uno.systemPathToFileUrl(temp_pdf_path)
    # Call storeToURL with FilterName and FilterData if supported
    try:
        doc.storeToURL(pdf_url, (pv("FilterName", filter_name), pv("FilterData", filter_data)))
    except Exception:
        # fallback: try minimal args
        doc.storeToURL(pdf_url, (pv("FilterName", filter_name),))

    if not os.path.exists(temp_pdf_path):
        raise RuntimeError("Export failed: PDF not found at expected path.")
    return temp_pdf_path


def send_via_https(upload_url, pdf_path, token=None):
    if requests is None:
        raise RuntimeError("requests library not available. Install it in the office python environment.")
    headers = {}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    with open(pdf_path, "rb") as f:
        files = {"file": (os.path.basename(pdf_path), f, "application/pdf")}
        print(f"[+] Uploading {pdf_path} -> {upload_url} ...")
        r = requests.post(upload_url, files=files, headers=headers, timeout=30)
        r.raise_for_status()
        print("[+] Upload success, server responded:", r.status_code, r.text[:400])
        return r.text


def send_via_websocket(ws_url, pdf_path, token=None):
    if websocket is None:
        raise RuntimeError("websocket-client library not available. Install it in the office python environment.")
    print(f"[+] Opening websocket to {ws_url} ...")
    # If you need headers (e.g., Authorization), websocket-client accepts header list
    header = None
    if token:
        header = [f"Authorization: Bearer {token}"]

    ws = websocket.create_connection(ws_url, header=header, timeout=20)
    try:
        with open(pdf_path, "rb") as f:
            data = f.read()
            print(f"[+] Sending {len(data)} bytes as binary frame over WebSocket ...")
            ws.send_binary(data)
            # Optionally wait for a server ack
            try:
                ack = ws.recv()
                print("[+] Received ACK:", ack)
            except Exception:
                pass
    finally:
        ws.close()
    return True


def main():
    parser = argparse.ArgumentParser(description="UNO export + send feasibility test")
    parser.add_argument("--ws-url", help="WebSocket URL to send binary PDF (ws://...)")
    parser.add_argument("--upload-url", help="HTTP(S) upload endpoint for multipart file upload")
    parser.add_argument("--token", help="Optional bearer token for auth (used for both WS and HTTPS)")
    parser.add_argument("--port", type=int, default=2002, help="UNO socket port (default 2002)")
    parser.add_argument("--host", default="127.0.0.1", help="UNO socket host (default 127.0.0.1)")
    args = parser.parse_args()

    if args.ws_url is None and args.upload_url is None:
        parser.error("Provide at least --ws-url or --upload-url for testing")

    # Build temp file path
    fd, temp_pdf = tempfile.mkstemp(prefix="uno_export_", suffix=".pdf")
    os.close(fd)
    try:
        print("[*] Connecting to UNO...")
        # connect_uno will attempt to resolve; connectivity errors get raised
        # we pass through host/port in export function by set globals or call directly
        # modify connect_uno to use passed args: quick monkey patch here:
        # For clarity we directly call connect_uno with args:
        try:
            ctx = None
            # attempt direct resolution to provide earlier failure messaging
            # (connect_uno uses default 127.0.0.1:2002; so call with args)
            local_ctx = uno.getComponentContext()
            resolver = local_ctx.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", local_ctx
            )
            url = f"uno:socket,host={args.host},port={args.port};urp;StarOffice.ComponentContext"
            ctx = resolver.resolve(url)
            # if ok, proceed to export using same mechanism but bypass connect_uno here
        except Exception as e:
            raise RuntimeError(f"Cannot connect to UNO at {args.host}:{args.port} — is soffice started with --accept? ({e})")

        print("[+] UNO reachable. Exporting active document to PDF...")
        pdf_path = export_active_to_pdf(temp_pdf)

        print("[+] PDF exported:", pdf_path)

        # prefer websocket if provided
        if args.ws_url:
            try:
                send_via_websocket(args.ws_url, pdf_path, token=args.token)
                print("[+] Sent via websocket OK")
                return
            except Exception as e:
                print("[!] WebSocket send failed:", e)
                print(traceback.format_exc())

        if args.upload_url:
            try:
                resp = send_via_https(args.upload_url, pdf_path, token=args.token)
                print("[+] Sent via HTTPS OK")
                print("Server response (truncated):", str(resp)[:500])
                return
            except Exception as e:
                print("[!] HTTPS upload failed:", e)
                print(traceback.format_exc())

        print("[!] No transport succeeded.")
    finally:
        try:
            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)
        except Exception:
            pass


if __name__ == "__main__":
    main()