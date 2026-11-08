@echo off
setlocal
cd /d "%~dp0"

set "PYEXE="
python --version >nul 2>&1 && set "PYEXE=python"
if not defined PYEXE (
    for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python*") do (
        if exist "%%D\python.exe" if not defined PYEXE set "PYEXE=%%D\python.exe"
    )
)

if not defined PYEXE (
    echo [ERROR] Python not found. Install Python 3.9+ with "Add Python to PATH".
    pause
    exit /b 1
)

echo Python: %PYEXE%

%PYEXE% -c "import streamlit" >nul 2>&1
if errorlevel 1 (
    echo [1/2] Installing dependencies, please wait...
    %PYEXE% -m pip install -r "%~dp0requirements.txt" -i https://pypi.tuna.tsinghua.edu.cn/simple
    if errorlevel 1 (
        echo [ERROR] Install failed.
        pause
        exit /b 1
    )
) else (
    echo [1/2] Dependencies OK.
)

echo [2/2] Starting...
start "" http://127.0.0.1:8501
%PYEXE% -m streamlit run "%~dp0app.py"

pause
endlocal
