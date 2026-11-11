@echo off
chcp 65001 >nul
title Conversor e Juntador Word para PDF (Python)
cd /d "%~dp0"

if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" (
    "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" "%~dp0app.py"
) else (
    python "%~dp0app.py"
)
pause
