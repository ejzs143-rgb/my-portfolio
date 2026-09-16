@echo off
chcp 65001 > nul
setlocal

echo ========================================
echo   QualityMetrics Monthly Update Tool
echo   (QualityMetrics月次実績 自動更新ツール)
echo ========================================
echo.

set BASEDIR=%~dp0
set SRC_DIR=%BASEDIR%input_hata
set WORK_DIR=%BASEDIR%input_working
set OUT_DIR=%BASEDIR%output

REM --- Pythonの存在確認 ---
where python >nul 2>nul
if errorlevel 1 (
    echo [エラー] Pythonが見つかりません。
    echo https://www.python.org/ からインストールしてください。
    echo インストール時は「Add python.exe to PATH」に必ずチェックを入れてください。
    echo.
    pause
    exit /b 1
)

echo [OK] Python が見つかりました。
python --version
echo.

REM --- openpyxlの確認・自動インストール ---
python -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
    echo openpyxl が未インストールのため、インストールします...
    python -m pip install openpyxl
    if errorlevel 1 (
        echo.
        echo [エラー] openpyxl のインストールに失敗しました。
        echo 社内ネットワーク／プロキシ設定が原因の可能性があります。IT部門にご確認ください。
        echo.
        pause
        exit /b 1
    )
    echo [OK] openpyxl をインストールしました。
    echo.
)

REM --- フォルダ確認（中身のファイル選択はPython側が全件スキャンして行う） ---
if not exist "%SRC_DIR%" mkdir "%SRC_DIR%"
if not exist "%WORK_DIR%" mkdir "%WORK_DIR%"
if not exist "%OUT_DIR%" mkdir "%OUT_DIR%"

echo 入力フォルダの中身を確認しています...
echo.

set OUT_FILE=%OUT_DIR%\QualityMetrics実績_更新後.xlsx

python "%BASEDIR%msqp_import.py" --source "%SRC_DIR%" --working "%WORK_DIR%" --output "%OUT_FILE%"
set PYRESULT=%ERRORLEVEL%

echo.
echo ========================================
if %PYRESULT% EQU 0 (
    echo 完了しました。
    echo 結果ファイル: %OUT_FILE%
    echo.
    echo 内容を確認のうえ、問題なければ元のworking fileへ反映してください。
    echo （このツールは元ファイルを直接上書きしません）
) else (
    echo 処理中にエラーが発生しました。上記のメッセージをご確認ください。
)
echo ========================================
echo.
pause
