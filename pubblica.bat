@echo off
setlocal
cd /d "%~dp0"
title Pubblica una nuova versione di FastText

rem Uso: doppio clic (chiede il numero) oppure  pubblica.bat 1.4
rem Fa tutto: aggiorna il numero di versione, commit, push, tag e invio del tag.
rem GitHub poi crea da solo l'exe e la Release (Actions > "Crea exe e Release").

where git >nul 2>nul || (echo Git non trovato. & pause & exit /b 1)
set "PY="
py -3 -c "import sys" >nul 2>nul && set "PY=py -3"
if not defined PY python -c "import sys" >nul 2>nul && set "PY=python"
if not defined PY (echo Python non trovato: esegui prima installa.bat & pause & exit /b 1)

for /f "delims=" %%t in ('git describe --tags --abbrev^=0 2^>nul') do set "LAST=%%t"
if defined LAST echo Ultima versione pubblicata: %LAST%

set "VER=%~1"
if "%VER%"=="" set /p "VER=Numero della nuova versione (es. 1.4): "
if "%VER%"=="" exit /b 1
rem accetta anche "v1.4" o "v.1.4"
if /i "%VER:~0,1%"=="v" set "VER=%VER:~1%"
if "%VER:~0,1%"=="." set "VER=%VER:~1%"
echo %VER%| findstr /r "^[0-9][0-9]*\.[0-9][0-9.]*$" >nul || (echo Numero non valido: usa il formato 1.4 oppure 1.4.1 & pause & exit /b 1)

git rev-parse -q --verify "refs/tags/v%VER%" >nul 2>nul && (echo La versione v%VER% esiste gia'. Scegli un numero piu' alto. & pause & exit /b 1)

echo.
echo Pubblico la versione v%VER% ...
%PY% -c "import re,sys;p='ftcore.py';s=open(p,encoding='utf-8').read();s=re.sub(r'(?m)^APP_VERSION = .*$','APP_VERSION = %%r' %% sys.argv[1],s,count=1);open(p,'w',encoding='utf-8',newline='').write(s)" %VER%
if errorlevel 1 goto fail

git add -A
git diff --cached --quiet || git commit -q -m "Versione %VER%"
if errorlevel 1 goto fail
git push
if errorlevel 1 goto fail
git tag v%VER%
git push origin v%VER%
if errorlevel 1 goto fail

echo.
echo ============================================================
echo  Fatto! GitHub sta creando l'exe (circa 5 minuti).
echo  Avanzamento:  https://github.com/CarloMaker/FastText/actions
echo  Download:     https://github.com/CarloMaker/FastText/releases
echo ============================================================
start "" https://github.com/CarloMaker/FastText/actions
pause
exit /b 0

:fail
echo.
echo Qualcosa non e' andato a buon fine: leggi il messaggio qui sopra.
pause
exit /b 1
