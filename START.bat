@echo off
cd /d "%~dp0"

title TrackMyDay - EWIDENCJA CZASU PRACY
echo ========================================
echo  Uruchamianie TrackMyDay - EWIDENCJA CZASU PRACY
echo ========================================
echo.

:: Sprawdzenie czy Python jest dostępny
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python nie znaleziony w zmiennej PATH!
    echo Sprawdz czy Python jest zainstalowany.
    pause
    exit /b 1
)

:: Sprawdzenie czy port 8501 jest już zajęty
echo Sprawdzam czy port 8501 jest wolny...
netstat -ano | findstr ":8501" | findstr "LISTENING" >nul
if %errorlevel% equ 0 (
    echo ERROR: Port 8501 jest już zajęty!
    echo Prawdopodobnie aplikacja już działa.
    echo Uruchom STOP.bat aby zamknąć istniejące instancje.
    pause
    exit /b 1
)

:: Tworzenie timestamp dla logów
for /f "tokens=2 delims==" %%a in ('wmic os get localdatetime /value') do set datetime=%%a
set log_date=%datetime:~0,4%-%datetime:~4,2%-%datetime:~6,2%
set log_time=%datetime:~8,2%-%datetime:~10,2%-%datetime:~12,2%

echo Uruchamianie aplikacji...
echo Logi aplikacji będą zapisywane do server_log.txt
echo Logi Streamlit będą zapisywane do streamlit_output.txt
echo.

:: Uruchomienie aplikacji z zapisem logów Streamlit do oddzielnego pliku (nadpisywanie przy każdym uruchomieniu)
python -m streamlit run app.py --server.port 8501 --logger.level=error > streamlit_output.txt 2>&1

echo.
echo Aplikacja została zamknięta.
echo Naciśnij dowolny klawisz aby kontynuować...
pause >nul