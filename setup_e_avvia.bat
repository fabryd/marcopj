@echo off
setlocal

echo Verifica presenza di Python...

where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Python non è installato!!!! :-(
    echo Esegui Miniconda3-latest-Windows-x86_64 che trovi in APP
    start https://www.python.org/downloads/
    pause
    exit /b
)

echo Python trovato :-).

echo ===============================
echo Aggiorno pip...
echo ===============================
python -m pip install --upgrade pip

echo ===============================
echo Installazione dei pacchetti...
echo ===============================
pip install -r requirements.txt

echo ===============================
echo Avvio dell'applicazione
echo ===============================

python -m streamlit run c:\marcopj\app\app_marcopj.py

pause
