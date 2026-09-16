@echo off
title SupportDesk Tool

set "PYFILE=%~dp0whd_reply_tool_v3.py"

echo ============================================================
echo   SupportDesk Reply Draft Generator
echo   Example Company
echo ============================================================
echo.
echo Checking libraries...
pip install tkcalendar pywin32 pdfplumber google-genai --quiet
echo.
python "%PYFILE%"
pause
