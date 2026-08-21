@echo off
chcp 65001 > nul
echo Starting Screening Engine...
python "%~dp0..\screener_core.py" "%~dp0"
pause

