@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo =========================================
echo Excel to PDF 変換ツール (ファイル選択版)
echo =========================================

set PYTHON_SCRIPT="%~dp0excel_to_pdf.py"

:: 仮想環境(venv)を使用する場合は "python" をフルパスに変更
python %PYTHON_SCRIPT%

echo =========================================
echo すべての処理が完了しました。
pause