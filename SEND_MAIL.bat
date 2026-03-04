@echo off
chcp 65001 >nul
title TrackMyDay - Send Report
echo ========================================
echo  TrackMyDay - EMAIL REPORT SENDING
echo ========================================
echo.

echo [1/3] Checking if app is running...
netstat -ano | findstr ":8501" >nul
if NOT errorlevel 1 goto server_running

echo   -> Server not running (port 8501 free).
goto check_done

:server_running
echo   -> Warning: Server running on port 8501.
echo        Recommended to stop before sending.
echo.
set /p choice=Continue? (Y/N): 
if /i "%choice%" neq "Y" (
    echo   Canceled.
    pause
    exit /b
)

:check_done
echo [2/3] Starting report sending...
echo   Using Python script...
python app.py --send-mail
if errorlevel 1 (
    echo   ERROR: Failed to send email!
) else (
    echo   OK: Sending completed.
)

echo.
echo [3/3] Summary...
if not exist server_log.txt goto no_logs

echo   Logs saved in server_log.txt
echo   Last 5 entries:
echo   --------------------------
setlocal enabledelayedexpansion
set count=0
for /f "usebackq delims=" %%a in ("server_log.txt") do (
    set /a count+=1
    set line!count!=%%a
)
set /a start=count-4
if !start! lss 1 set start=1
for /l %%i in (!start!,1,!count!) do (
    echo   !line%%i!
)
endlocal
goto summary_done

:no_logs
echo   No log file.

:summary_done
echo.
echo ========================================
echo  Sending completed.
echo  Check Outlook (Sent folder).
echo ========================================
echo.
pause