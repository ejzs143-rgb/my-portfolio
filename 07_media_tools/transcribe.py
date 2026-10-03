"""
音声ファイル文字起こしツール（faster-whisper / CPU動作）

音声・動画ファイルを選択すると、日本語で文字起こししてテキストファイルに保存する。
初回のみ Whisper モデルのダウンロードが発生する。

使い方:
    python transcribe.py
"""
import sys
import tkinter as tk
from tkinter import filedialog

from faster_whisper import WhisperModel

MODEL_SIZE = "small"       # tiny / base / small / medium / large-v3
COMPUTE_TYPE = "int8"      # CPU向けの軽量量子化


def main() -> None:
    root = tk.Tk()
    root.withdraw()

    print("文字起こしする音声ファイルを選択してください...")
    input_file = filedialog.askopenfilename(
        title="文字起こしする音声ファイルを選択",
        filetypes=[("音声ファイル", "*.mp3 *.wav *.m4a *.flac *.mp4"), ("すべてのファイル", "*.*")],
    )
    if not input_file:
        print("ファイル選択がキャンセルされました。")
        sys.exit(0)

    print("結果の保存先を選択してください...")
    output_file = filedialog.asksaveasfilename(
        title="文字起こし結果の保存先を指定",
        defaultextension=".txt",
        filetypes=[("テキストファイル", "*.txt"), ("すべてのファイル", "*.*")],
    )
    if not output_file:
        print("保存先の選択がキャンセルされました。")
        sys.exit(0)

    print("-" * 30)
    print("AIモデルを読み込んでいます（初回はダウンロードに数分かかります）...")
    model = WhisperModel(MODEL_SIZE, device="cpu", compute_type=COMPUTE_TYPE)

    print("文字起こしを開始します...")
    segments, _info = model.transcribe(
        input_file, beam_size=5, language="ja", condition_on_previous_text=False
    )

    # 進捗を画面に出しながら逐次書き込む
    with open(output_file, "w", encoding="utf-8") as f:
        for segment in segments:
            print(f"[{segment.start:.1f}s -> {segment.end:.1f}s] {segment.text}")
            f.write(segment.text + "\n")

    print("-" * 30)
    print(f"完了しました。結果を保存しました: {output_file}")


if __name__ == "__main__":
    main()
