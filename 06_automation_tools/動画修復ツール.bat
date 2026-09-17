@echo off
chcp 65001 >nul
setlocal EnableExtensions DisableDelayedExpansion

echo ==========================================
echo Untrunc MP4修復ツール
echo シンプル都度選択版
echo ==========================================
echo.

echo untrunc.exe を選択してください...

for /f "delims=" %%I in ('powershell -NoProfile -STA -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName System.Windows.Forms; $f=New-Object System.Windows.Forms.OpenFileDialog; $f.Filter='untrunc.exe|untrunc.exe|EXEファイル (*.exe)|*.exe|すべてのファイル (*.*)|*.*'; $f.Title='untrunc.exe を選択してください'; if($f.ShowDialog() -eq 'OK'){ $f.FileName }"') do set "UNTRUNC_EXE=%%I"

if not defined UNTRUNC_EXE (
    echo キャンセルされました。
    pause
    exit /b 1
)

if not exist "%UNTRUNC_EXE%" (
    echo 【エラー】untrunc.exe が見つかりません。
    echo(%UNTRUNC_EXE%
    pause
    exit /b 1
)

for %%F in ("%UNTRUNC_EXE%") do set "UNTRUNC_DIR=%%~dpF"

echo.
echo 選択された untrunc.exe:
echo(%UNTRUNC_EXE%
echo.

echo 修復したい破損MP4を選択してください...

for /f "delims=" %%I in ('powershell -NoProfile -STA -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName System.Windows.Forms; $f=New-Object System.Windows.Forms.OpenFileDialog; $f.Filter='MP4動画 (*.mp4)|*.mp4|すべてのファイル (*.*)|*.*'; $f.Title='修復したい破損MP4を選択してください'; if($f.ShowDialog() -eq 'OK'){ $f.FileName }"') do set "CORRUPT_FILE=%%I"

if not defined CORRUPT_FILE (
    echo キャンセルされました。
    pause
    exit /b 1
)

if not exist "%CORRUPT_FILE%" (
    echo 【エラー】破損MP4が見つかりません。
    echo(%CORRUPT_FILE%
    pause
    exit /b 1
)

echo.
echo 選択された破損MP4:
echo(%CORRUPT_FILE%
echo.

echo 正常な参照MP4を選択してください...

for /f "delims=" %%I in ('powershell -NoProfile -STA -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName System.Windows.Forms; $f=New-Object System.Windows.Forms.OpenFileDialog; $f.Filter='MP4動画 (*.mp4)|*.mp4|すべてのファイル (*.*)|*.*'; $f.Title='同じ機器・同じ設定で録画された正常MP4を選択してください'; if($f.ShowDialog() -eq 'OK'){ $f.FileName }"') do set "NORMAL_FILE=%%I"

if not defined NORMAL_FILE (
    echo キャンセルされました。
    pause
    exit /b 1
)

if not exist "%NORMAL_FILE%" (
    echo 【エラー】正常MP4が見つかりません。
    echo(%NORMAL_FILE%
    pause
    exit /b 1
)

echo.
echo 選択された正常MP4:
echo(%NORMAL_FILE%
echo.

echo 修復後ファイルの保存先フォルダを選択してください...

for /f "delims=" %%I in ('powershell -NoProfile -STA -ExecutionPolicy Bypass -Command "Add-Type -AssemblyName System.Windows.Forms; $f=New-Object System.Windows.Forms.FolderBrowserDialog; $f.Description='修復後のファイルを保存するフォルダを選択してください'; $f.ShowNewFolderButton=$true; if($f.ShowDialog() -eq 'OK'){ $f.SelectedPath }"') do set "SAVE_DIR=%%I"

if not defined SAVE_DIR (
    echo キャンセルされました、または保存先フォルダを取得できませんでした。
    pause
    exit /b 1
)

if not exist "%SAVE_DIR%" (
    echo 【エラー】保存先フォルダが見つかりません。
    echo(%SAVE_DIR%
    pause
    exit /b 1
)

echo.
echo 選択された保存先:
echo(%SAVE_DIR%
echo.

for %%F in ("%CORRUPT_FILE%") do (
    set "CORRUPT_DIR=%%~dpF"
    set "CORRUPT_NAME=%%~nF"
    set "CORRUPT_EXT=%%~xF"
)

set "FIXED_FILENAME=%CORRUPT_NAME%_fixed%CORRUPT_EXT%"
set "FIXED_FILE_1=%CORRUPT_DIR%%FIXED_FILENAME%"
set "FIXED_FILE_2=%UNTRUNC_DIR%%FIXED_FILENAME%"
set "FIXED_FILE_3=%CD%\%FIXED_FILENAME%"

echo.
echo ==========================================
echo 修復処理を開始します
echo ==========================================
echo.
echo 実行コマンド:
echo "%UNTRUNC_EXE%" -s "%NORMAL_FILE%" "%CORRUPT_FILE%"
echo.

pushd "%CORRUPT_DIR%"

"%UNTRUNC_EXE%" -s "%NORMAL_FILE%" "%CORRUPT_FILE%"

set "UNTRUNC_EXIT=%ERRORLEVEL%"

popd

echo.
echo untrunc 終了コード: %UNTRUNC_EXIT%
echo.

if exist "%FIXED_FILE_1%" (
    move /y "%FIXED_FILE_1%" "%SAVE_DIR%\%FIXED_FILENAME%" >nul
    goto SUCCESS
)

if exist "%FIXED_FILE_2%" (
    move /y "%FIXED_FILE_2%" "%SAVE_DIR%\%FIXED_FILENAME%" >nul
    goto SUCCESS
)

if exist "%FIXED_FILE_3%" (
    move /y "%FIXED_FILE_3%" "%SAVE_DIR%\%FIXED_FILENAME%" >nul
    goto SUCCESS
)

goto FAILED

:SUCCESS
echo.
echo ==========================================
echo 修復が完了しました！
echo 保存先:
echo(%SAVE_DIR%\%FIXED_FILENAME%
echo ==========================================
echo.
pause
exit /b 0

:FAILED
echo.
echo 【エラー】修復されたファイルが見つかりませんでした。
echo.
echo 探した場所:
echo(%FIXED_FILE_1%
echo(%FIXED_FILE_2%
echo(%FIXED_FILE_3%
echo.
echo 原因候補:
echo 1. untrunc の修復に失敗した
echo 2. 参照MP4と破損MP4の録画条件が違う
echo 3. 破損が深刻すぎる
echo 4. untrunc の出力ファイル名が想定と違う
echo.
pause
exit /b 1