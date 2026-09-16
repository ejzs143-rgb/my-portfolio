import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os

def run_app():
    root = tk.Tk()
    root.withdraw()

    # --- 1. FFmpegの場所を確定させる ---
    # まずは予想される場所をチェック
    default_ffmpeg = r"C:\Users\Public\Desktop\動画圧縮\ffmpeg-8.0-essentials_build\bin\ffmpeg.exe"
    
    if os.path.exists(default_ffmpeg):
        ffmpeg_exe = default_ffmpeg
    else:
        # 見つからない場合は、ユーザーに選んでもらう
        messagebox.showinfo("確認", "FFmpeg本体(ffmpeg.exe)を選択してください。")
        ffmpeg_exe = filedialog.askopenfilename(
            title="ffmpeg.exe を選択してください",
            filetypes=[("FFmpeg実行ファイル", "ffmpeg.exe")]
        )

    if not ffmpeg_exe:
        return

    # --- 2. 圧縮したい動画を選択 ---
    input_path = filedialog.askopenfilename(title="圧縮したい元の動画を選択してください")
    if not input_path: return

    # --- 3. 保存先を指定 ---
    output_path = filedialog.asksaveasfilename(
        title="保存先と名前を決めてください",
        defaultextension=".mp4",
        filetypes=[("MP4 files", "*.mp4")]
    )
    if not output_path: return

    # --- 4. 圧縮レベルを選択 ---
    level = simpledialog.askstring("圧縮設定", "レベルを選んでください (1:高画質 2:標準 3:低画質 4:最小)", initialvalue="2")
    crf_map = {"1": "18", "2": "23", "3": "28", "4": "32"}
    crf = crf_map.get(level, "23")

    # --- 5. 実行 ---
    cmd = [
        ffmpeg_exe, "-i", input_path,
        "-vcodec", "libx264", "-crf", crf,
        "-pix_fmt", "yuv420p", "-y", output_path
    ]

    print(f"実行中: {input_path} -> {output_path}")
    try:
        # 実行中の進捗が見えるようにコマンドプロンプトを開く
        subprocess.run(cmd, check=True)
        messagebox.showinfo("成功", "動画の圧縮が完了しました！")
    except Exception as e:
        messagebox.showerror("エラー", f"失敗しました:\n{e}")

if __name__ == "__main__":
    run_app()
