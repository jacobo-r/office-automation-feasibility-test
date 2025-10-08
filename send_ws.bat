@echo off
setlocal

set LOGBASE=D:\USUARIOS (NO BORRAR)\ADMINISTRADOR\Downloads\office-automation-feasibility-test-main
set TRACEFILE=%LOGBASE%\macro_trace.log
set WSLOG=%LOGBASE%\ws_send_log_from_macro.txt
set PYTHON=C:\Users\Administrador\AppData\Local\Programs\Python\Python314\python.exe
set SCRIPT=%LOGBASE%\ws_send_pdf.py
set PDFFOLDER=%LOGBASE%\temp_pdf

echo =============================================== >> "%TRACEFILE%"
echo [%date% %time%] Batch started >> "%TRACEFILE%"
echo [%date% %time%] Current directory: %CD% >> "%TRACEFILE%"

REM --- Quick sanity test: can Python even start?
echo [%date% %time%] Testing Python startup... >> "%TRACEFILE%"
powershell -Command "Start-Process '%PYTHON%' -ArgumentList '-c \"import sys; print(\'Hello from\', sys.executable)\"' -Verb RunAs" >> "%TRACEFILE%" 2>&1
echo [%date% %time%] Test exit code: %ERRORLEVEL% >> "%TRACEFILE%"

REM --- Record Python version (useful for path mismatches)
"%PYTHON%" -V >> "%TRACEFILE%" 2>&1

REM --- Run the actual WebSocket sender
echo [%date% %time%] Running Python line... >> "%TRACEFILE%"
echo CMD: "%PYTHON%" "%SCRIPT%" ws://127.0.0.1:9000 "%PDFFOLDER%" >> "%TRACEFILE%"

"%PYTHON%" "%SCRIPT%" ws://127.0.0.1:9000 "%PDFFOLDER%" >> "%WSLOG%" 2>&1
echo [%date% %time%] Python exit code: %ERRORLEVEL% >> "%TRACEFILE%"

echo [%date% %time%] Batch finished >> "%TRACEFILE%"
echo. >> "%TRACEFILE%"

endlocal
pause
