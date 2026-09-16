@echo off
echo 動画圧縮ツール
echo.
cd ffmpeg-8.0-essentials_build
set /p input="動画ファイルをドラッグ&ドロップ: "
set input=%input:"=%
for %%F in ("%input%") do set filename=%%~nF
bin\ffmpeg.exe -i "%input%" -vcodec libx264 -crf 23 "%filename%_圧縮済み.mp4"
pause
