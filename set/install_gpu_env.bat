@echo off
cd /d "%~dp0..\project"
setlocal EnableExtensions
set NO_COLOR=1

rem 與 run_web_hidden.vbs 一致：優先 py -3，否則 python（避免裝到 A 顆 Python、跑後端卻用 B 顆）
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
rem PATH 找不到時，嘗試使用使用者層級安裝路徑（例如 LocalAppData\Programs\Python\Python312）
for /d %%D in ("%LocalAppData%\Programs\Python\Python3*") do (
  if exist "%%~fD\python.exe" (
    set "PYRUN=%%~fD\python.exe"
    goto have_python
  )
)
echo ERROR: 找不到 Python。請安裝 3.10+ 並勾選 PATH，或安裝 Python Launcher ^(py^)。
echo        也可直接使用完整路徑執行 python.exe -m pip ...
exit /b 1

:have_python
echo Using %PYRUN%
echo [1/3] PyTorch with CUDA 12.8 (NVIDIA, from pytorch.org)
echo      RTX 50 系列 ^(Blackwell / sm_120^) 需 cu128 以上；cu124 版 torch 常導致 CUDA 對 PyTorch 不可用。
%PYRUN% -m pip install -U pip
if errorlevel 1 goto :fail
%PYRUN% -m pip install torch torchvision --index-url https://download.pytorch.org/whl/cu128
if errorlevel 1 goto :fail
echo [2/3] Other packages (easyocr etc. will not replace CUDA torch if version matches)
%PYRUN% -m pip install -r requirements.txt
if errorlevel 1 goto :fail
echo [3/3] Force CUDA wheels again (fixes accidental CPU torch from transitive deps)
%PYRUN% -m pip install torch torchvision --index-url https://download.pytorch.org/whl/cu128 --force-reinstall
if errorlevel 1 goto :fail
echo.
echo Verify:
%PYRUN% -c "import torch; print('torch=', torch.__version__, 'cuda=', torch.version.cuda, 'available=', torch.cuda.is_available()); print('device=', torch.cuda.get_device_name(0) if torch.cuda.is_available() else None)"
%PYRUN% -c "import torch; assert torch.cuda.is_available(), 'CUDA not available (check GPU torch wheel, not CPU)'; x=torch.zeros(1, device='cuda'); print('cuda tensor ok', x.device)"
%PYRUN% -c "import importlib.util; print('transformers OK' if importlib.util.find_spec('transformers') else 'transformers MISSING - check pip install')"
echo.
echo Done. Double-click run_web.bat in the repo root, or run_web_console.bat in this folder for debug logs.
exit /b 0
:fail
echo FAILED. Check errors above.
exit /b 1
