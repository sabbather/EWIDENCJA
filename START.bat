@echo off
cd /d "%~dp0"
setlocal

title TrackMyDay - EWIDENCJA CZASU PRACY
echo ========================================
echo  Uruchamianie TrackMyDay - EWIDENCJA CZASU PRACY
echo ========================================
echo.

:: Priorytet: lokalny Python 3.13, potem python z PATH
set "PYTHON_CMD=%LocalAppData%\Programs\Python\Python313\python.exe"
if not exist "%PYTHON_CMD%" (
    set "PYTHON_CMD=python"
    python --version >nul 2>&1
    if errorlevel 1 (
        echo ERROR: Nie znaleziono Pythona ani w %LocalAppData% ani w PATH.
        echo Zainstaluj 3.13 z opcja Add to PATH.
        pause
        exit /b 1
    )
)

echo Sprawdzam czy port 8501 jest wolny...
netstat -ano | findstr ":8501" | findstr "LISTENING" >nul
if %errorlevel% equ 0 (
    echo ERROR: Port 8501 jest juz zajety!
    echo Prawdopodobnie aplikacja juz dziala.
    echo Uruchom STOP.bat aby zamknac istniejace instancje.
    pause
    exit /b 1
)

echo Uruchamianie aplikacji...
echo Logi aplikacji beda zapisywane do server_log.txt
echo Logi Streamlit beda zapisywane do streamlit_output.txt
echo.
echo Uzywam polecenia: %PYTHON_CMD%

call "%PYTHON_CMD%" -m streamlit run app.py --server.port 8501 --logger.level=error > streamlit_output.txt 2>&1

echo.
echo Aplikacja zostala zamknieta.
echo Nacisnij dowolny klawisz aby kontynuowac...
pause >nul
