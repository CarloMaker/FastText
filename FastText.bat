@echo off
rem Avvia FastText senza finestra nera. Accetta un file come argomento.
set "SCRIPT=%~dp0fasttext.py"
where pyw >nul 2>nul && goto usepyw
where pythonw >nul 2>nul && goto usepythonw
echo Python non trovato: esegui prima installa.bat
pause
exit /b 1
:usepyw
start "" pyw -3 "%SCRIPT%" %*
exit /b 0
:usepythonw
start "" pythonw "%SCRIPT%" %*
exit /b 0
