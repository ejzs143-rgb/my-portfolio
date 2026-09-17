import os
import tkinter as tk
from tkinter import filedialog
import win32com.client

def select_folder(title):
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    path = filedialog.askdirectory(title=title)
    root.destroy()
    return path

def word_direct_compress():
    input_dir = select_folder("【1/2】圧縮したいPDFがあるフォルダを選択")
    if not input_dir: return
    output_dir = select_folder("【2/2】保存先フォルダを選択")
    if not output_dir: return

    files = [f for f in os.listdir(input_dir) if f.lower().endswith(".pdf")]
    print(f"合計 {len(files)} 件のWord直接圧縮を開始します...")

    # Wordを起動
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False 
    except Exception as e:
        print(f"Wordの起動に失敗しました: {e}")
        return

    for filename in files:
        in_path = os.path.abspath(os.path.join(input_dir, filename))
        out_path = os.path.abspath(os.path.join(output_dir, "最強圧縮_" + filename))
        
        try:
            print(f"--- 処理中: {filename} ---")
            # WordでPDFを開く
            doc = word.Documents.Open(in_path)
            
            # PDFとしてエクスポート (17=wdExportFormatPDF, OptimizeFor=1=最小サイズ)
            doc.ExportAsFixedFormat(out_path, 17, OptimizeFor=1)
            
            doc.Close(False) # 保存せずに閉じる
            
            orig = os.path.getsize(in_path) / 1024
            new = os.path.getsize(out_path) / 1024
            print(f"【成功】 {orig:.1f}KB -> {new:.1f}KB")
            
        except Exception as e:
            print(f"【エラー】 {filename}: {e}")

    word.Quit()
    print("\nすべての処理が完了しました！元データはそのままです。")

if __name__ == "__main__":
    word_direct_compress()