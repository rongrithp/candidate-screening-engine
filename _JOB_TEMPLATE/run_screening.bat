@echo off
chcp 65001 > nul
echo Starting Candidate Screening Engine...

if exist "%~dp0..\screener_core.exe" (
    "%~dp0..\screener_core.exe" "%~dp0"
) else if exist "%~dp0..\screener_core.py" (
    python "%~dp0..\screener_core.py" "%~dp0"
) else (
    echo Error: Neither screener_core.exe nor screener_core.py was found in parent directory.
)
pause
