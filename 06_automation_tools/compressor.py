import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image

def main():
    # GUIウィンドウの初期化と非表示
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)

    # 1. ファイル選択ダイアログ
    file_paths = filedialog.askopenfilenames(
        title="圧縮する画像ファイルを選択してください（複数選択可）",
        filetypes=[("Image files", "*.jpg;*.jpeg;*.png;*.bmp")]
    )
    
    if not file_paths:
        return 

    # 2. 保存先フォルダ選択ダイアログ
    save_dir = filedialog.askdirectory(
        title="保存先のフォルダを選択してください"
    )

    if not save_dir:
        return 

    # 画質優先の設定に変更（85〜95が一般的。ここでは85とし、最適化フラグを使用）
    quality_setting = 85
    success_count = 0

    # 3. 圧縮処理ループ
    for path in file_paths:
        try:
            img = Image.open(path)
            
            # 透過PNG等をJPEG保存可能にするためのRGB変換
            if img.mode in ("RGBA", "P"):
                img = img.convert("RGB")
            
            _, file_name = os.path.split(path)
            base_name, _ = os.path.splitext(file_name)
            out_path = os.path.join(save_dir, f"{base_name}_compressed.jpg")
            
            # optimize=True を追加し、画質を維持しつつファイル構造を最適化
            img.save(out_path, "JPEG", quality=quality_setting, optimize=True)
            success_count += 1
            
        except Exception as e:
            messagebox.showerror("エラー", f"画像処理中にエラーが発生しました。\nファイル: {path}\n詳細: {e}")

    # 4. 完了通知
    if success_count > 0:
        messagebox.showinfo("処理完了", f"{success_count}件の画像圧縮が完了しました。")

if __name__ == "__main__":
    main()