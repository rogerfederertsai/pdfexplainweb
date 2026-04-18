@echo off
cd /d "%~dp0..\project"
setlocal EnableExtensions
set NO_COLOR=1

echo ============================================================
echo Project env setup (Python 3.10+)
echo ============================================================
echo.
echo Notes:
echo   1. Upgrade pip, install CPU torch, then requirements.txt
echo   2. For NVIDIA CUDA use install_gpu_env.bat
echo   3. Re-run after installing Python if PATH is not active yet
echo.

set "PYRUN="
where py >nul 2>&1
if errorlevel 1 goto try_python
py -3 -c "import sys" 2>nul
if errorlevel 1 goto try_python
set "PYRUN=py -3"
goto have_python

:try_python
where python >nul 2>&1
if errorlevel 1 goto no_python
python -c "import sys" 2>nul
if errorlevel 1 goto no_python
set "PYRUN=python"
goto have_python

:no_python
echo ERROR: Python not found. Install 3.10+ from https://www.python.org/downloads/
echo        Enable Add python.exe to PATH, then re-run.
pause
exit /b 1

:have_python
echo Using %PYRUN%
%PYRUN% -c "import sys; raise SystemExit(0 if sys.hexversion >= 0x30A00F0 else 1)" 2>nul
if errorlevel 1 goto need_py310
%PYRUN% -c "import sys; print('Python', sys.version.split()[0])"
goto pip_upgrade

:need_py310
echo ERROR: Need Python 3.10 or newer.
pause
exit /b 1

:pip_upgrade
echo.
%PYRUN% -m pip install -U pip setuptools wheel
if errorlevel 1 goto fail

echo.
echo [1/3] PyTorch CPU (PyPI)
%PYRUN% -m pip install torch torchvision
if errorlevel 1 goto fail

echo.
echo [2/3] requirements.txt
%PYRUN% -m pip install -r requirements.txt
if errorlevel 1 goto req_warn

echo.
echo [3/3] Import check
%PYRUN% -c "import torch; print('torch', torch.__version__, 'CUDA' if torch.cuda.is_available() else 'CPU')"
if errorlevel 1 goto fail
%PYRUN% -c "import fastapi, uvicorn; print('FastAPI / Uvicorn OK')"
if errorlevel 1 goto fail

echo.
echo ============================================================
echo Done. Double-click run_web.bat in the repo root for headless web UI.
echo NVIDIA: run set\install_gpu_env.bat for CUDA.
echo ============================================================
pause
exit /b 0

:req_warn
echo.
echo WARN: requirements.txt had failures (bitsandbytes on Windows is common).
echo        Web/EasyOCR may still work. See project notes for VLM 4-bit.
pause
exit /b 1

:fail
echo.
echo ERROR: Install stopped. Save log output for troubleshooting.
pause
exit /b 1
