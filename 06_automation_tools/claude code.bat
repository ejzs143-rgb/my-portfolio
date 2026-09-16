@echo off
chcp 65001 >nul
setlocal

:: ==========================================
:: 1. ドラッグ＆ドロップ対応（そのまま残します）
:: ==========================================
if not "%~1"=="" (
    set "TARGET_DIR=%~1"
    goto :LAUNCH
)

:: ==========================================
:: 2. いつものエクスプローラー画面での選択
:: ==========================================
echo フォルダ選択画面を起動しています...
for /f "delims=" %%I in ('powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Title = '目的のフォルダに入った状態で、右下の【開く】を押してください'; $d.FileName = 'このフォルダで起動'; $d.CheckFileExists = $false; $d.CheckPathExists = $true; $d.ValidateNames = $false; $d.Filter = 'フォルダ|*.none'; if($d.ShowDialog() -eq 'OK'){ [System.IO.Path]::GetDirectoryName($d.FileName) }"') do set "TARGET_DIR=%%I"

:: キャンセルボタンが押された場合はそのまま終了
if "%TARGET_DIR%"=="" exit /b

:LAUNCH
:: ==========================================
:: 3. 対象フォルダでAIを起動
:: ==========================================
start powershell -NoExit -Command "Set-Location -LiteralPath '%TARGET_DIR%'; Write-Host '📂 選択されたフォルダ: %TARGET_DIR%' -ForegroundColor Cyan; Write-Host '🚀 Claude Code を起動します...' -ForegroundColor Green; claude"
