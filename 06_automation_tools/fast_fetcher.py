import os
import shutil
import tkinter as tk
from tkinter import messagebox, ttk, filedialog

# --- ユーザー情報の取得 ---
USER_PROFILE = os.environ.get("USERPROFILE")

# --- デフォルトで開くサーバーのパスを指定 ---
DEFAULT_SERVER_PATH = r"\\fileserver\shared"
# ------------------------------------------------------

class FastFetcherApp:
    def __init__(self, root):
        self.root = root
        self.root.title("爆速ファイルフェッチャー (保存先選択版)")
        self.root.geometry("600x500")
        
        # 実際にリストに表示されているフォルダのパスを保持する変数
        self.loaded_path = ""
        
        # ダイアログで開く初期フォルダ（最初はデスクトップ、以降は最後に保存した場所を記憶）
        self.last_save_dir = os.path.join(USER_PROFILE, "Desktop")

        # --- UI構築 ---
        # 1. パス入力欄
        tk.Label(root, text="1. サーバーのフォルダパスをコピペ (Enterで取得):", font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 0))
        
        input_frame = tk.Frame(root)
        input_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self.path_var = tk.StringVar()
        self.path_combo = ttk.Combobox(input_frame, textvariable=self.path_var, font=("Arial", 10))
        self.path_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        self.path_combo.bind("<Return>", lambda event: self.load_file_list())
        
        self.load_btn = tk.Button(input_frame, text="一覧を爆速取得", command=self.load_file_list, bg="#2196F3", fg="white", font=("Arial", 9, "bold"))
        self.load_btn.pack(side=tk.RIGHT)

        # 2. ファイル/フォルダ一覧リストボックス
        tk.Label(root, text="2. 項目を選択 (フォルダはWクリックで中へ潜れます):", font=("Arial", 10, "bold")).pack(anchor="w", padx=10, pady=(10, 0))
        
        list_frame = tk.Frame(root)
        list_frame.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        
        self.scrollbar = tk.Scrollbar(list_frame)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.file_listbox = tk.Listbox(list_frame, yscrollcommand=self.scrollbar.set, selectmode=tk.SINGLE, font=("Arial", 10))
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.scrollbar.config(command=self.file_listbox.yview)

        # ダブルクリック時の挙動をバインド
        self.file_listbox.bind("<Double-1>", self.on_double_click)

        # 3. ダウンロードボタン
        self.download_btn = tk.Button(root, text="選択した項目(ファイル/フォルダ)をローカルに落とす", command=self.download_btn_clicked, bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), pady=5)
        self.download_btn.pack(fill=tk.X, padx=10, pady=15)

        # --- 起動時の自動読み込み処理 ---
        # デフォルトのパスを入力欄にセットする
        self.path_var.set(DEFAULT_SERVER_PATH)
        # 画面の描画が完了した直後(0.2秒後)に、自動でリスト取得処理を走らせる
        self.root.after(200, self.load_file_list)

    def load_file_list(self):
        """指定パス内のファイル・フォルダ一覧を取得して表示"""
        target_path = self.path_var.get().strip()
        target_path = target_path.strip('"').strip("'")
        
        if not target_path:
            return
            
        self.file_listbox.delete(0, tk.END)
        
        if not os.path.exists(target_path):
            messagebox.showerror("エラー", f"アクセスできません。\nパスが間違っているか、VPN接続が切れています。\n\n入力パス:\n{target_path}")
            return
            
        try:
            # os.listdir で中身を取得し、フォルダとファイルに分ける
            items = os.listdir(target_path)
            self.loaded_path = target_path
            
            folders = []
            files = []
            
            for item in items:
                item_full_path = os.path.join(self.loaded_path, item)
                if os.path.isdir(item_full_path):
                    folders.append(item)
                else:
                    files.append(item)
            
            # アルファベット順に並び替え
            folders.sort()
            files.sort()
            
            # --- リストへの追加処理 ---
            # 1. 「上の階層へ戻る」ボタンを追加
            parent_dir = os.path.dirname(self.loaded_path)
            if self.loaded_path != parent_dir:
                self.file_listbox.insert(tk.END, "[..] (上の階層へ)")
            
            # 2. フォルダを先に追加
            for folder in folders:
                self.file_listbox.insert(tk.END, f"[フォルダ] {folder}")
                
            # 3. ファイルをその後に追加
            for f in files:
                self.file_listbox.insert(tk.END, f)
            
            # 履歴（コンボボックス）の更新
            current_history = list(self.path_combo["values"])
            if self.loaded_path not in current_history:
                current_history.insert(0, self.loaded_path)
                self.path_combo["values"] = current_history
                
        except Exception as e:
            messagebox.showerror("読み込みエラー", f"一覧の取得に失敗しました:\n{e}")

    def on_double_click(self, event):
        """リスト内の項目をダブルクリックした時の動作を制御"""
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            return
            
        selection = self.file_listbox.get(selected_indices[0])
        
        # 1. 「上の階層へ」をダブルクリックした場合 -> 親フォルダへ移動
        if selection == "[..] (上の階層へ)":
            parent_dir = os.path.dirname(self.loaded_path)
            self.path_var.set(parent_dir)
            self.load_file_list()
            
        # 2. フォルダをダブルクリックした場合 -> そのフォルダの中へ潜る
        elif selection.startswith("[フォルダ] "):
            folder_name = selection.replace("[フォルダ] ", "", 1)
            new_path = os.path.join(self.loaded_path, folder_name)
            self.path_var.set(new_path)
            self.load_file_list()
            
        # 3. ファイルをダブルクリックした場合 -> 保存先を選んでダウンロード
        else:
            self.execute_download(selection, is_folder=False)

    def download_btn_clicked(self):
        """緑色のダウンロードボタンを押した時の動作を制御"""
        if not self.loaded_path:
            return

        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("警告", "ダウンロードする項目を選択してください。")
            return
            
        selection = self.file_listbox.get(selected_indices[0])
        
        if selection == "[..] (上の階層へ)":
            messagebox.showinfo("案内", "この項目はダウンロードできません。\nダブルクリックで上の階層へ移動します。")
            return
            
        # フォルダが選択された場合
        if selection.startswith("[フォルダ] "):
            folder_name = selection.replace("[フォルダ] ", "", 1)
            self.execute_download(folder_name, is_folder=True)
        # ファイルが選択された場合
        else:
            self.execute_download(selection, is_folder=False)

    def execute_download(self, item_name, is_folder):
        """実際のコピー処理を実行する（保存先ダイアログ付き）"""
        src_path = os.path.join(self.loaded_path, item_name)
        
        # --- 追加: 保存先フォルダを選択するダイアログを表示 ---
        dest_dir = filedialog.askdirectory(
            title=f"「{item_name}」の保存先フォルダを選択",
            initialdir=self.last_save_dir # 前回開いた場所を初期位置にする
        )
        
        # キャンセルボタンが押された場合（空文字が返る）は処理を中断
        if not dest_dir:
            return
            
        # 次回のために選んだフォルダを記憶しておく
        self.last_save_dir = dest_dir
        
        # 最終的な保存先のフルパスを作成
        dst_path = os.path.join(dest_dir, item_name)
        # ----------------------------------------------------
        
        try:
            if is_folder:
                # フォルダごとコピー（既に同じフォルダがあっても上書きする設定）
                shutil.copytree(src_path, dst_path, dirs_exist_ok=True)
            else:
                # ファイル単体をコピー
                shutil.copy2(src_path, dst_path)
                
            messagebox.showinfo("完了", f"保存が完了しました！\n\n【保存先】\n{dst_path}")
        except Exception as e:
            messagebox.showerror("コピーエラー", f"コピーに失敗しました:\n{e}")

if __name__ == "__main__":
    root = tk.Tk()
    app = FastFetcherApp(root)
    root.mainloop()