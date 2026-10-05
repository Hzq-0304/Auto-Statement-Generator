@echo off
chcp 65001 >nul
cd /d "%~dp0"
if exist ".venv\Scripts\pythonw.exe" (
    start "" ".venv\Scripts\pythonw.exe" main.py
    exit /b
)
where py >nul 2>nul
if errorlevel 1 (
    echo 请先按照 README 安装 Python，或使用 dist 中已打包的程序。
    pause
    exit /b 1
)
py main.py
if errorlevel 1 pause
