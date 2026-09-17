import os
import tkinter as tk
from tkinter import filedialog
import win32com.client
import time

def select_files(title):
    """複数のファイルを選択するダイアログを表示"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    # 複数ファイルを選択できるように askopenfilenames を使用
    file_paths = filedialog.askopenfilenames(
        title=title,
        filetypes=[
            ("Officeファイル", "*.docx;*.doc;*.xlsx;*.xls;*.pptx;*.ppt"),
            ("すべてのファイル", "*.*")
        ]
    )
    root.destroy()
    return file_paths

def select_folder(title):
    """保存先フォルダを選択するダイアログを表示"""
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    path = filedialog.askdirectory(title=title)
    root.destroy()
    return path

def convert_selected_to_pdf():
    # 1. 変換したいファイルを選択（複数選択可能）
    input_files = select_files("【1/2】PDFに変換したいOfficeファイルを選択（複数選択可）")
    if not input_files: 
        print("ファイルの選択がキャンセルされました。")
        return

    # 2. PDFを保存するフォルダを選択
    output_dir = select_folder("【2/2】PDFを保存するフォルダを選択")
    if not output_dir: 
        print("保存先の選択がキャンセルされました。")
        return

    apps = {"word": None, "excel": None, "powerpoint": None}
    
    print(f"変換を開始します...")

    # 選択されたファイルのフルパスリストをループ処理
    for in_path in input_files:
        # パスからファイル名を取得
        filename = os.path.basename(in_path)
        
        # 開いているOfficeファイルの一時ファイル（~$から始まるファイル）はスキップ
        if filename.startswith("~$"):
            continue

        ext = os.path.splitext(filename)[1].lower()
        
        # 入力パスと出力パスを絶対パスに変換
        in_path_abs = os.path.abspath(in_path)
        out_path_abs = os.path.abspath(os.path.join(output_dir, os.path.splitext(filename)[0] + ".pdf"))

        try:
            # Word / Excel 
            if ext in [".docx", ".doc"]:
                if not apps["word"]: apps["word"] = win32com.client.Dispatch("Word.Application")
                doc = apps["word"].Documents.Open(in_path_abs, ReadOnly=True)
                doc.ExportAsFixedFormat(out_path_abs, 17)
                doc.Close(False)
                print(f"【成功】 {filename}")

            elif ext in [".xlsx", ".xls"]:
                if not apps["excel"]: apps["excel"] = win32com.client.Dispatch("Excel.Application")
                wb = apps["excel"].Workbooks.Open(in_path_abs, ReadOnly=True)
                wb.ExportAsFixedFormat(0, out_path_abs)
                wb.Close(False)
                print(f"【成功】 {filename}")

            # PowerPoint (ブロック回避の最終形態)
            elif ext in [".pptx", ".ppt"]:
                if not apps["powerpoint"]:
                    # アプリ本体をまず「完全に」起動させる
                    apps["powerpoint"] = win32com.client.Dispatch("PowerPoint.Application")
                
                # これが重要：アプリを表に見せて「動いている」とWindowsに認識させる
                apps["powerpoint"].Visible = True 
                time.sleep(1) # 起動待ちの余裕を持たせる

                # プレゼンテーションを開く
                pres = apps["powerpoint"].Presentations.Open(in_path_abs, -1, -1, -1) 
                
                # 32 = ppSaveAsPDF
                pres.SaveAs(out_path_abs, 32)
                pres.Close()
                print(f"【成功】 {filename}")
                
            else:
                print(f"【スキップ】 {filename} (未対応の拡張子です)")

        except Exception as e:
            print(f"【失敗】 {filename}: {e}")

    print("\nアプリを終了しています...")
    try:
        if apps["word"]: apps["word"].Quit()
        if apps["excel"]: apps["excel"].Quit()
        if apps["powerpoint"]: apps["powerpoint"].Quit()
    except:
        pass
    
    print(f"すべて完了しました！")

if __name__ == "__main__":
    convert_selected_to_pdf()