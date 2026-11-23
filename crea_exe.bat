@echo off
setlocal
cd /d "%~dp0"
title Creazione FastText.exe

set "PY="
py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
if not defined PY python -c "import sys" >nul 2>nul && set "PY=python"
if not defined PY (
  echo Python non trovato: esegui prima installa.bat
  pause
  exit /b 1
)

echo Installo PyInstaller...
%PY% -m pip install --upgrade --user pyinstaller
if errorlevel 1 goto fail

echo Creo l'eseguibile (ci vuole qualche minuto)...
%PY% -m PyInstaller --noconfirm --clean --windowed --name FastText --icon fasttext.ico --hidden-import ftcore --hidden-import ftsearch --hidden-import fttextops --hidden-import ftpanels fasttext.py
if errorlevel 1 goto fail

echo.
echo ============================================================
echo  Pronto:  %~dp0dist\FastText\FastText.exe
echo  Puoi copiare l'intera cartella dist\FastText dove vuoi.
echo  Poi lancia associa_ftx.bat per aprire i .ftx con doppio clic.
echo ============================================================
pause
exit /b 0

:fail
echo Creazione non riuscita.
pause
exit /b 1
