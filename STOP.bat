@echo off
title Zamykanie TrackMyDay
echo ========================================
echo  Zamykanie TrackMyDay - EWIDENCJA CZASU PRACY
echo ========================================
echo.

echo [1/3] Zamykanie serwera na porcie 8501...
for /f "tokens=5" %%a in ('netstat -ano ^| findstr :8501 ^| findstr LISTENING') do (
    echo Znaleziono proces PID: %%a
    taskkill /F /PID %%a >nul 2>&1
    if errorlevel 1 (
        echo   -> Nie uda?o si? zamkn??.
    ) else (
        echo   -> Proces %%a zosta? zamkni?ty.
    )
)

echo [2/3] Zamykanie proces?w Python...
taskkill /F /IM python.exe >nul 2>&1
taskkill /F /IM pythonw.exe >nul 2>&1
echo   -> Polecenie zamkni?cia wys?ane.

echo [3/3] Czyszczenie...
del /F /Q port_8501_pids.tmp cmdline.tmp guard_pids.tmp 2>nul
echo   -> Pliki tymczasowe usuni?te.

echo.
echo ========================================
echo Aplikacja TrackMyDay zosta?a zatrzymana.
echo ========================================
echo.
pause
