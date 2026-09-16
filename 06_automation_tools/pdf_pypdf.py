import pypdf
import os
import tkinter as tk
from tkinter import filedialog

def extract():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    # 入力フォルダ選択
    print("変換したいPDFがあるフォルダを選択してください...")
    input_dir = filedialog.askdirectory(title="PDFフォルダを選択")
    if not input_dir: 
        print("キャンセルされました")
        return
    
    # 出力先選択
    print("変換結果の保存先を選択してください...")
    output_dir = filedialog.askdirectory(title="保存先フォルダを選択")
    if not output_dir:
        print("キャンセルされました")
        return
    
    # PDF処理
    count = 0
    for filename in os.listdir(input_dir):
        if filename.lower().endswith(".pdf"):
            print(f"処理中: {filename}")
            in_path = os.path.join(input_dir, filename)
            out_path = os.path.join(output_dir, filename.replace(".pdf", ".txt"))
            
            try:
                reader = pypdf.PdfReader(in_path)
                with open(out_path, "w", encoding="utf-8") as f:
                    for i, page in enumerate(reader.pages):
                        f.write(f"\n--- Page {i+1} ---\n")
                        f.write(page.extract_text())
                print(f"✓ 完了: {filename}")
                count += 1
            except Exception as e:
                print(f"✗ エラー: {filename} - {e}")
    
    print(f"\n{'='*50}")
    print(f"変換完了: {count}個のファイル")
    print(f"保存先: {output_dir}")
    print(f"{'='*50}")

if __name__ == '__main__':
    extract()
