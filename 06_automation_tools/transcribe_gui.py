import os
import warnings
import ssl
import httpx

# --- 会社のネットワーク（SSL通信）の壁を強制突破する最強のおまじない ---
# SSL verification must remain enabled
# Do not override the default SSL context

# ダウンロードエンジンのセキュリティチェックを強制的に無効化
# Do not monkey-patch HTTP clients to disable certificate verification

warnings.filterwarnings("ignore")
# -------------------------------------------------------------------

import tkinter as tk
from tkinter import filedialog
from faster_whisper import WhisperModel

# 裏で開く不要なウィンドウを隠す
root = tk.Tk()
root.withdraw()

print("文字起こしする音声ファイルを選択してください...")
input_file = filedialog.askopenfilename(
    title="文字起こしする音声ファイルを選択",
    filetypes=[("音声ファイル", "*.mp3 *.wav *.m4a *.flac *.mp4"), ("すべてのファイル", "*.*")]
)

if not input_file:
    print("ファイル選択がキャンセルされました。")
    exit()

print("結果の保存先を選択してください...")
output_file = filedialog.asksaveasfilename(
    title="文字起こし結果の保存先を指定",
    defaultextension=".txt",
    filetypes=[("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")]
)

if not output_file:
    print("保存先の選択がキャンセルされました。")
    exit()

print("-" * 30)
print("初回のみ、AIモデルのダウンロードが行われます（数分かかります）...")
print("AIモデルを読み込んでいます...")

# CPU向けに最適化（int8という軽量化技術を使って爆速にします）
model = WhisperModel("small", device="cpu", compute_type="int8")

print("文字起こしを開始します。進捗がここに表示されます...")
segments, info = model.transcribe(input_file, beam_size=5, language="ja", condition_on_previous_text=False)

# リアルタイムで画面に表示しながらテキストファイルに書き込む
with open(output_file, "w", encoding="utf-8") as f:
    for segment in segments:
        print(f"[{segment.start:.1f}s -> {segment.end:.1f}s] {segment.text}")
        f.write(segment.text + "\n")

print("-" * 30)
print(f"完了しました！ 結果を以下に保存しました：\n{output_file}")