@echo off
setlocal
chcp 65001 > nul

echo === 診断開始 ===

:: 1. Python自体の存在チェック
where python >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo 【致命的】パソコンがPythonを認識できていません。
    echo 対策: Pythonをインストールし直すか、「Path」の設定が必要です。
    pause
    exit /b
)

:: 2. ファイルの存在チェック
if not exist "recompress_ultra.py" (
    echo 【致命的】recompress_ultra.py が同じフォルダに見当たりません。
    echo 現在のフォルダにあるファイル:
    dir /b
    pause
    exit /b
)

:: 3. 実行
echo 処理を実行します...
python recompress_ultra.py "%~1"

echo.
echo === 処理終了（この画面が消えずに残っているはずです） ===
pause
