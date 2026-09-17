import sys
import os
import subprocess
import tkinter as tk
from tkinter import filedialog

def compress_audio(input_file):
    if not os.path.exists(input_file):
        print(f"エラー: ファイルが見つかりません: {input_file}")
        return False

    file_dir = os.path.dirname(input_file)
    file_name = os.path.basename(input_file)
    name, _ = os.path.splitext(file_name)
    
    # 修正箇所1: 保存先を元ファイルと同一のフォルダ直下に変更
    # ※元ファイルと同名にするとFFmpegが読込/書込競合でクラッシュするため、末尾の付与は維持
    output_file = os.path.join(file_dir, f"{name}_compressed.m4a")

    # 修正箇所2: 音声に特化した高圧縮パラメータの設定
    target_bitrate = "32k" # ビットレートを32kbpsへ引き下げ
    channels = "1"         # ステレオからモノラルへ変換しデータ量半減
    sample_rate = "24000"  # サンプリングレートを下げて高音域のデータ削減

    command = [
        "ffmpeg",
        "-y",
        "-i", input_file,
        "-vn",
        "-c:a", "aac",
        "-b:a", target_bitrate,
        "-ac", channels,         # チャンネル数の指定追加
        "-ar", sample_rate,      # サンプリングレートの指定追加
        "-map_metadata", "0",
        output_file
    ]

    print(f"処理中: {file_name} -> {target_bitrate}, モノラル, {sample_rate}Hz で高圧縮中...")
    
    try:
        result = subprocess.run(
            command, 
            stdout=subprocess.PIPE, 
            stderr=subprocess.PIPE, 
            text=True,
            encoding='utf-8',
            errors='replace'
        )
        if result.returncode == 0:
            print(f"完了: {os.path.normpath(output_file)}")
            return True
        else:
            print(f"エラー: 処理中にエラーが発生しました。\n{result.stderr}")
            return False
    except FileNotFoundError:
        print("致命的エラー: FFmpegがインストールされていないか、環境変数PATHが設定されていません。")
        sys.exit(1)

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    file_paths = filedialog.askopenfilenames(
        title="圧縮する音声ファイルを選択してください（複数選択可）",
        filetypes=[
            ("ExampleBrand Bo Files", "*.m4a;*.mp3;*.wav;*.flac;*.aac;*.wma"),
            ("All Files", "*.*")
        ]
    )

    if not file_paths:
        print("ファイルの選択がキャンセルされました。処理を終了します。")
        sys.exit(0)
    
    for file_path in file_paths:
        compress_audio(file_path)
        
    print("すべての処理が完了しました。")