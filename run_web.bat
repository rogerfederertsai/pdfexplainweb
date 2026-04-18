@echo off
cd /d "%~dp0"
set NO_COLOR=1
rem Starts server if needed (detached); opens browser tab; safe to double-click again (see project\run_web_launcher.py)

where wscript >nul 2>&1
if errorlevel 1 goto nowscript
start "" wscript //nologo //B "%~dp0set\run_web_hidden.vbs"
exit /b 0

:nowscript
where pyw >nul 2>&1
if errorlevel 1 goto trypythonw
start "" pyw -3 "%~dp0project\run_web_launcher.py"
exit /b 0

:trypythonw
where pythonw >nul 2>&1
if errorlevel 1 goto nointerpreter
start "" pythonw "%~dp0project\run_web_launcher.py"
exit /b 0

:nointerpreter
echo ERROR: Need wscript or pyw/pythonw. Install Python 3.10+ with PATH.
exit /b 1
