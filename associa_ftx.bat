@echo off
setlocal
cd /d "%~dp0"
rem Associa i file .ftx a FastText (solo per l'utente corrente, non servono permessi di amministratore)
rem e aggiunge "Apri con FastText" al menu del tasto destro dei file .docx e .txt

set "EXE=%~dp0dist\FastText\FastText.exe"
set "ICON=%~dp0fasttext.ico"
if exist "%EXE%" goto haveexe

set "PYW="
for /f "delims=" %%i in ('where pythonw 2^>nul') do if not defined PYW set "PYW=%%i"
if not defined PYW (
  echo Non trovo ne' FastText.exe ne' pythonw. Esegui prima installa.bat o crea_exe.bat
  pause
  exit /b 1
)
set "CMD=\"%PYW%\" \"%~dp0fasttext.py\" \"%%1\""
goto doreg

:haveexe
set "CMD=\"%EXE%\" \"%%1\""
set "ICON=%EXE%,0"

:doreg
reg add "HKCU\Software\Classes\.ftx" /ve /d "FastText.Document" /f >nul
reg add "HKCU\Software\Classes\FastText.Document" /ve /d "Documento FastText" /f >nul
reg add "HKCU\Software\Classes\FastText.Document\DefaultIcon" /ve /d "%ICON%" /f >nul
reg add "HKCU\Software\Classes\FastText.Document\shell\open\command" /ve /d "%CMD%" /f >nul
for %%e in (.docx .txt .md) do (
  reg add "HKCU\Software\Classes\SystemFileAssociations\%%e\shell\FastText" /ve /d "Apri con FastText" /f >nul
  reg add "HKCU\Software\Classes\SystemFileAssociations\%%e\shell\FastText\command" /ve /d "%CMD%" /f >nul
)
ie4uinit.exe -show >nul 2>nul
echo Fatto: i file .ftx si aprono con FastText.
echo Sui .docx, .txt e .md trovi "Apri con FastText" nel menu del tasto destro
echo (su Windows 11: "Mostra altre opzioni").
pause
