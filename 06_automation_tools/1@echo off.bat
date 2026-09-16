@echo off
echo FFmpeg Simple Compressor
echo.
set /p input="Drag video file here and press Enter: "
set input=%input:"=%
for %%F in ("%input%") do set filename=%%~nF
bin\ffmpeg.exe -i "%input%" -vcodec libx264 -crf 23 -preset medium "%filename%_compressed.mp4"
echo.
echo Done! Check for %filename%_compressed.mp4
pause
