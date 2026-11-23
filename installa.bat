@echo off
setlocal
cd /d "%~dp0"
title Installazione FastText

set "PY="
py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
if not defined PY python -c "import sys" >nul 2>nul && set "PY=python"
if not defined PY goto nopython

echo Python trovato:
%PY% --version
echo.
echo Installo le librerie necessarie (PySide6, python-docx, pillow)...
%PY% -m pip install --upgrade --user -r requirements.txt
if errorlevel 1 goto pipfail

echo.
echo ============================================================
echo  Fatto! Avvia FastText con un doppio clic su FastText.bat
echo  (facoltativo) crea_exe.bat  = crea FastText.exe
echo  (facoltativo) associa_ftx.bat = apri i .ftx con doppio clic
echo ============================================================
pause
exit /b 0

:nopython
echo Python non risulta installato.
echo.
echo Installalo in uno di questi modi, poi rilancia installa.bat:
echo   1) da https://www.python.org/downloads/
echo      (durante l'installazione spunta "Add python.exe to PATH")
echo   2) dal prompt dei comandi:  winget install Python.Python.3.12
pause
exit /b 1

:pipfail
echo.
echo Installazione delle librerie non riuscita. Controlla la connessione e riprova.
pause
exit /b 1
