@echo off
chcp 65001 >nul
echo =========================================
echo M4A to WAV 自動変換ツール (FFmpeg Direct版)
echo =========================================

echo [1/2] 必要なパッケージのインストール状況を確認しています...

python -c "import imageio_ffmpeg" >nul 2>&1
if %ERRORLEVEL% equ 0 goto SKIP_INSTALL

echo [未検出] imageio-ffmpegをインストールします...
python -m pip install imageio-ffmpeg

if %ERRORLEVEL% neq 0 (
    echo [エラー] インストールに失敗しました。
    pause
    exit /b
)
goto RUN_TOOL

:SKIP_INSTALL
echo [OK] 準備完了。インストール処理をスキップします。

:RUN_TOOL
echo.
echo [2/2] ツールを起動しています...
python converter.py

echo.
echo 全ての処理が終了しました。
pause