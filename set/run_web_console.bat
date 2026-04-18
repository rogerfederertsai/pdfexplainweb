@echo off
cd /d "%~dp0..\project"
set NO_COLOR=1
echo Console mode: server logs below. Press Ctrl+C or close window to stop.
echo Open http://127.0.0.1:8000/ in your browser when ready.
python -m uvicorn web.app:app --host 0.0.0.0 --port 8000 --no-use-colors
