@echo off
setlocal
cd /d %~dp0

echo ==================================================
echo   Video Compression Tool (Python Launcher)
echo ==================================================

:: 1. Pythonがインストールされているか確認
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Pythonが見つかりません。
    echo Pythonをインストールし、「Add Python to PATH」にチェックを入れてください。
    pause
    exit
)

:: 2. Pythonファイルが存在するか確認
if not exist "video_tool.py" (
    echo [ERROR] video_tool.py が見つかりません。
    echo このバッチファイルと同じフォルダに置いてください。
    pause
    exit
)

:: 3. 実行
echo アプリを起動しています...
python video_tool.py

echo.
echo プログラムが終了しました。
pause