import os
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
import imageio_ffmpeg

def select_files():
    """ファイル選択ダイアログを表示し、選択されたファイルのパスリストを返す"""
    root = tk.Tk()
    root.withdraw()
    
    file_paths = filedialog.askopenfilenames(
        title="文字起こし用に変換するM4Aファイルを選択してください",
        filetypes=[("M4A Audio", "*.m4a"), ("All Files", "*.*")]
    )
    return file_paths

def main():
    files = select_files()
    if not files:
        print("キャンセルされました。ファイルが選択されていません。")
        return

    # imageio_ffmpegからFFmpegの実行ファイルパスを直接取得
    ffmpeg_exe = imageio_ffmpeg.get_ffmpeg_exe()
    success_count = 0

    for file_path in files:
        try:
            print(f"変換中: {os.path.basename(file_path)} ...")
            base_name, _ = os.path.splitext(file_path)
            output_path = f"{base_name}_converted.wav"
            
            # FFmpegを直接叩くコマンドの構成（-y は上書き許可の設定）
            command = [
                ffmpeg_exe,
                "-y",
                "-i", file_path,
                output_path
            ]
            
            # 変換実行（ログが大量に出るのを防ぐため出力を抑制）
            subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=True)
            
            print(f"完了: {output_path}")
            success_count += 1
            
        except Exception as e:
            print(f"[エラー] {os.path.basename(file_path)} の処理中にエラーが発生しました: {e}")

    # 完了時のポップアップ通知
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("変換完了", f"{success_count}個のファイルをWAV形式に変換しました。")

if __name__ == "__main__":
    main()
