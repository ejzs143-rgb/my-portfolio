@echo off
chcp 65001 >nul
echo ==========================================
echo    FFmpeg 動画圧縮ツール v2.1（修正版）
echo ==========================================
echo.

REM 現在のディレクトリを確認
echo 現在の場所: %CD%
echo.

REM FFmpegの存在確認
if exist "ffmpeg.exe" (
    echo ✅ FFmpeg found: ffmpeg.exe
    set FFMPEG_PATH=ffmpeg.exe
) else if exist "bin\ffmpeg.exe" (
    echo ✅ FFmpeg found: bin\ffmpeg.exe
    set FFMPEG_PATH=bin\ffmpeg.exe
) else (
    echo ❌ FFmpegが見つかりません
    echo 以下を確認してください：
    echo 1. このバッチファイルがffmpegフォルダ内にあるか
    echo 2. bin\ffmpeg.exe または ffmpeg.exe が存在するか
    pause
    exit /b
)

echo.
echo 使い方：動画ファイルをドラッグ&ドロップしてEnterを押してください
echo.
set /p "input=ファイル: "

REM 引用符を除去
set input=%input:"=%

REM ファイルの存在確認
if not exist "%input%" (
    echo ❌ エラー: ファイルが見つかりません
    echo ファイルパス: %input%
    pause
    exit /b
)

REM ファイル名と拡張子を取得
for %%F in ("%input%") do (
    set "filename=%%~nF"
    set "extension=%%~xF"
    set "filepath=%%~dpF"
)

REM 出力ファイル名を設定
set "output=%filepath%%filename%_圧縮済み%extension%"

echo.
echo 📋 処理情報:
echo 入力: %input%
echo 出力: %output%
echo FFmpeg: %FFMPEG_PATH%
echo.
echo 処理開始...

REM FFmpeg実行
"%FFMPEG_PATH%" -i "%input%" -vcodec libx264 -crf 23 -preset medium "%output%"

if %errorlevel%==0 (
    echo.
    echo ✅ 圧縮完了！
    echo 📁 保存場所: %output%
) else (
    echo.
    echo ❌ 圧縮中にエラーが発生しました
)

echo.
pause
```

## 🗂️ **ファイル配置の確認**

正しいフォルダ構成になっているか確認してください：
```
📁 ffmpeg/
├── 📄 動画圧縮.bat ← 新しく作成するバッチファイル
├── 📁 bin/
│   ├── ffmpeg.exe ← 本体
│   ├── ffplay.exe
│   └── ffprobe.exe
├── 📁 doc/
└── 📁 presets/
```

## 🔄 **修正手順**

1. **古いバッチファイルを削除**
2. **新しいバッチファイルを作成**
   - ファイル名: `動画圧縮_修正版.bat`
   - 保存場所: ffmpegフォルダの一番上（binフォルダと同じ階層）

3. **実行してテスト**

## 🆘 **それでもエラーが出る場合の代替方法**

### **方法A: 直接コマンド実行**
1. **コマンドプロンプトを開く**
2. **ffmpegフォルダに移動**
```
   cd "C:\Users\Public\Desktop\ffmpeg"
```
3. **以下のコマンドを実行**
```
   bin\ffmpeg.exe -i "動画ファイル名.mp4" -vcodec libx264 -crf 23 "圧縮済み.mp4"