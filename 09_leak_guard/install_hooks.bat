@echo off
chcp 65001 > nul
rem leak_guard のフックをこのリポジトリで有効にする（1回だけ実行すればよい）
cd /d "%~dp0.."
git config core.hooksPath 09_leak_guard/hooks
if errorlevel 1 (
  echo [エラー] git の設定に失敗しました。このフォルダが Git リポジトリか確認してください。
  pause
  exit /b 1
)
if not exist "%USERPROFILE%\.leak_guard" mkdir "%USERPROFILE%\.leak_guard"
if not exist "%USERPROFILE%\.leak_guard\denylist.txt" (
  copy "%~dp0denylist.sample.txt" "%USERPROFILE%\.leak_guard\denylist.txt" > nul
  echo 禁止語リストのひな形を作成しました: %USERPROFILE%\.leak_guard\denylist.txt
  echo 社名・人名・社内システム名などを書き足してください（このファイルは公開されません）。
)
echo.
echo フックを有効にしました。以降のコミット／プッシュは自動で検査されます。
echo 過去の履歴も確認する場合: python 09_leak_guard\leak_guard.py --history
pause
