@echo off
setlocal
cd /d "%~dp0"
py -3 "%~dp0orquestrador.py" %*
set "CODIGO=%ERRORLEVEL%"
endlocal & exit /b %CODIGO%
