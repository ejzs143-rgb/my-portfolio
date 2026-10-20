import os
import tkinter as tk
from tkinter import filedialog
from pypdf import PdfWriter

def select_folder(title):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    path = filedialog.askdirectory(title=title)
    root.destroy()
    return path

def python_only_compress():
    # フォルダ選択（2回）
    input_dir = select_folder("【1/2】圧縮したいPDFがあるフォルダを選択")
    if not input_dir: return
    output_dir = select_folder("【2/2】圧縮後の保存先フォルダを選択")
    if not output_dir: return

    files = [f for f in os.listdir(input_dir) if f.lower().endswith(".pdf")]
    print(f"合計 {len(files)} 件のPython画像圧縮を開始します...")

    for filename in files:
        in_path = os.path.abspath(os.path.join(input_dir, filename))
        out_path = os.path.abspath(os.path.join(output_dir, "Py圧縮_" + filename))
        
        try:
            print(f"処理中: {filename}...")
            writer = PdfWriter(clone_from=in_path)

            # 各ページの画像を JPEG（画質30）で置き換える。
            # pypdf の replace を使うと圧縮形式の情報（/Filter 等）も合わせて書き換わるため、
            # JPEG 以外（スキャン画像に多い Flate 形式など）も正しく圧縮できる
            for page in writer.pages:
                for image in page.images:
                    try:
                        img = image.image
                        if img.mode in ("RGBA", "LA", "PA"):
                            continue  # 透過のある画像は JPEG にすると背景が変わるため残す
                        if img.mode not in ("RGB", "L", "CMYK"):
                            img = img.convert("RGB")
                        image.replace(img, quality=30)
                    except Exception:
                        # 抽出できない形式の画像はそのまま残す
                        continue
            
            with open(out_path, "wb") as f:
                writer.write(f)
            
            orig = os.path.getsize(in_path) / 1024
            new = os.path.getsize(out_path) / 1024
            print(f" 【成功】 {orig:.1f}KB -> {new:.1f}KB")

        except Exception as e:
            print(f" 【エラー】 {filename}: {e}")

    print("\nすべて完了しました！元データはそのままです。")

if __name__ == "__main__":
    python_only_compress()