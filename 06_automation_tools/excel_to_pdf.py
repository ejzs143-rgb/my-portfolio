import os
import sys
import pathlib
import win32com.client
import tkinter as tk
from tkinter import filedialog

def convert_excel_to_pdf(excel_path: str):
    abs_excel_path = str(pathlib.Path(excel_path).resolve())
    abs_pdf_path = str(pathlib.Path(excel_path).with_suffix('.pdf').resolve())

    if not os.path.exists(abs_excel_path):
        print(f"[スキップ] ファイルが見つかりません: {abs_excel_path}")
        return

    excel_app = win32com.client.DispatchEx("Excel.Application")
    excel_app.Visible = False
    excel_app.DisplayAlerts = False

    wb = None
    try:
        wb = excel_app.Workbooks.Open(abs_excel_path, ReadOnly=True)
        wb.ExportAsFixedFormat(0, abs_pdf_path, 0, True, False)
        print(f"[成功] {abs_pdf_path}")
    except Exception as e:
        print(f"[失敗] {abs_excel_path} - エラー詳細: {e}")
    finally:
        if wb:
            wb.Close(SaveChanges=False)
        if excel_app:
            excel_app.Quit()

def select_files():
    """ファイル選択ダイアログを表示し、選択されたファイルのパスリストを返す"""
    root = tk.Tk()
    root.withdraw() # 背面の空のメインウィンドウを非表示にする
    
    # 最前面に表示するための処理
    root.attributes('-topmost', True)
    
    file_paths = filedialog.askopenfilenames(
        title="変換するExcelファイルを選択してください（複数選択可）",
        filetypes=[("Excelファイル", "*.xlsx *.xls *.xlsm")]
    )
    return file_paths

if __name__ == "__main__":
    print("ファイル選択ダイアログを起動しています...")
    target_files = select_files()

    if not target_files:
        print("ファイルが選択されませんでした。処理をキャンセルします。")
        sys.exit(0)

    print(f"{len(target_files)}件のファイルが選択されました。変換を開始します。")
    for file_path in target_files:
        convert_excel_to_pdf(file_path)
