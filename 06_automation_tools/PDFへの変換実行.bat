@echo off
setlocal
cd /d %~dp0
python convert_to_pdf.py
pause