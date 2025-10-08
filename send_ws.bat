@echo off
echo [%date% %time%] Batch started >> "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasibility-test-main\macro_trace.log"
cd /d "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasibility-test-main"
"C:\Users\Administrador\AppData\Local\Programs\Python\Python314\python.exe" ws_send_pdf.py ws://127.0.0.1:9000 "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasibility-test-main\temp_pdf" >> ws_send_log_from_macro.txt 2>&1
echo [%date% %time%] Batch finished >> "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasibility-test-main\macro_trace.log"
pause
