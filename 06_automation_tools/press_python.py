import os
import tkinter as tk
from tkinter import filedialog
from pypdf import PdfReader, PdfWriter
from PIL import Image
import io

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
            reader = PdfReader(in_path)
            writer = PdfWriter()

            for page in reader.pages:
                writer.add_page(page)

            # 各ページの画像を探して圧縮
            for page in writer.pages:
                if "/Resources" in page and "/XObject" in page["/Resources"]:
                    xObject = page["/Resources"]["/XObject"].get_object()
                    for obj in xObject:
                        if xObject[obj]["/Subtype"] == "/Image":
                            try:
                                # 画像を抽出してリサイズ・画質落とし
                                img_data = xObject[obj].get_data()
                                img = Image.open(io.BytesIO(img_data))
                                
                                # 白黒化せずカラーを維持しつつ圧縮（画質30）
                                out_img = io.BytesIO()
                                img.save(out_img, format="JPEG", quality=30)
                                xObject[obj]._data = out_img.getvalue()
                            except:
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
