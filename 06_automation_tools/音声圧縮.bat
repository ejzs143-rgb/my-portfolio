@echo off
chcp 65001 > nul
echo =========================================
echo 音声一括圧縮ツール (ファイル選択方式)
echo =========================================

REM 自身のディレクトリにあるPythonスクリプトを実行
set PYTHON_SCRIPT="%~dp0compress_audio.py"
python %PYTHON_SCRIPT%

echo.
pause