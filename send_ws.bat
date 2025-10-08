@echo off
setlocal

echo =============================================== >> "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasibility-test-main\macro_trace.log"
echo [%date% %time%] Batch started >> "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasibility-test-main\macro_trace.log"

cd /d "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasibility-test-main"
echo [%date% %time%] Now in directory: %CD% >> "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasibility-test-main\macro_trace.log"

echo [%date% %time%] Running Python line... >> "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasibility-test-main\macro_trace.log"
echo CMD: "C:\Users\Administrador\AppData\Local\Programs\Python\Python314\python.exe" ws_send_pdf.py ws://127.0.0.1:9000 "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasibility-test-main\temp_pdf" >> "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasibility-test-main\macro_trace.log"

"C:\Users\Administrador\AppData\Local\Programs\Python\Python314\python.exe" ws_send_pdf.py ws://127.0.0.1:9000 "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasibility-test-main\temp_pdf" >> ws_send_log_from_macro.txt 2>&1

echo [%date% %time%] Python exit code: %ERRORLEVEL% >> "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasibility-test-main\macro_trace.log"
echo [%date% %time%] Batch finished >> "D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasibility-test-main\macro_trace.log"
endlocal
pause
