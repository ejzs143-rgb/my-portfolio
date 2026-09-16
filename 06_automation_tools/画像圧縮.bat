@echo off
cd /d "%~dp0"
python -c "import PIL" 2>nul || pip install Pillow
python compressor.py
