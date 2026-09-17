import os
import sys
import time
import string
import traceback
from datetime import datetime

# --- Windows特有の文字コードエラー（UnicodeEncodeError）を先回りで防止 ---
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except AttributeError:
        pass

# --- 基本設定 ---
USER_HOME = os.path.expanduser("~")
DATA_DIR = os.path.join(USER_HOME, ".my_search_data")
if not os.path.exists(DATA_DIR):
    os.makedirs(DATA_DIR)

def get_index_path(target_dir):
    safe_name = target_dir.replace(":\\", "_").replace("\\", "_").replace("/", "_").replace(" ", "_")
    return os.path.join(DATA_DIR, f"{safe_name}_index.tsv")

def detect_targets():
    targets = []
    targets.append(("現在のユーザーフォルダ", USER_HOME))
    
    try:
        for item in os.listdir(USER_HOME):
            if "OneDrive" in item or "SharePoint" in item:
                full_path = os.path.join(USER_HOME, item)
                if os.path.isdir(full_path):
                    targets.append((f"クラウド同期 ({item})", full_path))
    except Exception:
        pass

    for d in string.ascii_uppercase:
        drive = f"{d}:\\"
        if os.path.exists(drive):
            label = "システムドライブ" if d == "C" else "ネットワーク/外部ドライブ"
            targets.append((f"{label} ({drive})", drive))
            
    return targets

def build_index(target_dir, index_file):
    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] インデックス作成を開始します: {target_dir}")
    start_time = time.time()
    count = 0
    
    with open(index_file, "w", encoding="utf-8") as f:
        stack = [target_dir]
        while stack:
            current_dir = stack.pop()
            try:
                # follow_symlinks=True にすることで、SharePointのショートカットフォルダの奥まで辿る
                with os.scandir(current_dir) as it:
                    for entry in it:
                        try:
                            if entry.is_dir(follow_symlinks=True):
                                stack.append(entry.path)
                            elif entry.is_file(follow_symlinks=True):
                                stat = entry.stat(follow_symlinks=True)
                                f.write(f"{entry.path}\t{stat.st_mtime}\t{stat.st_ctime}\n")
                                count += 1
                        except (OSError, PermissionError):
                            continue
            except (OSError, PermissionError):
                continue

    elapsed = time.time() - start_time
    print(f"完了: {count}件のファイルを {elapsed:.2f}秒 で登録しました。")

def search_local(keyword, index_file, target_dir):
    if not os.path.exists(index_file):
        print("\nインデックスが存在しません。初回作成を自動で実行します。")
        build_index(target_dir, index_file)

    results = []
    keyword_lower = keyword.lower()
    
    try:
        with open(index_file, "r", encoding="utf-8") as f:
            for line in f:
                if keyword_lower in line.lower():
                    parts = line.rstrip("\n").split("\t")
                    if len(parts) == 3:
                        path, mtime_str, ctime_str = parts
                        results.append({
                            'path': path,
                            'mtime': float(mtime_str),
                            'ctime': float(ctime_str)
                        })
    except Exception as e:
        print(f"検索エラー: {e}")
        return

    results.sort(key=lambda x: x['mtime'], reverse=True)

    print(f"\n[{keyword}] の検索結果: {len(results)}件 (更新日時の新しい順)")
    print("-" * 110)
    print(f"{'更新日時':<19} | {'作成日時':<19} | パス")
    print("-" * 110)

    for res in results:
        m_date = datetime.fromtimestamp(res['mtime']).strftime('%Y-%m-%d %H:%M:%S')
        c_date = datetime.fromtimestamp(res['ctime']).strftime('%Y-%m-%d %H:%M:%S')
        # Windowsコマンドプロンプトで表示できない文字を「?」に置換してクラッシュを防ぐ
        safe_path = res['path'].encode('cp932', errors='replace').decode('cp932')
        print(f"{m_date} | {c_date} | {safe_path}")
    print("-" * 110)

def main():
    print("=== ローカル・クラウド同期フォルダ 高速検索システム ===")
    targets = detect_targets()
    
    print("\n検索対象を選択してください:")
    for i, (name, path) in enumerate(targets, 1):
        print(f" [{i}] {name}")
    print(f" [{len(targets) + 1}] パスを手動で入力する")

    while True:
        try:
            choice = input("\n番号を入力 > ").strip()
            if choice.isdigit():
                choice_idx = int(choice) - 1
                if 0 <= choice_idx < len(targets):
                    selected_name, selected_path = targets[choice_idx]
                    break
                elif choice_idx == len(targets):
                    selected_path = input("検索対象のフルパスを入力してください > ").strip()
                    selected_name = "カスタムパス"
                    if os.path.exists(selected_path):
                        break
                    else:
                        print("無効なパスです。")
            print("正しい番号を入力してください。")
        except KeyboardInterrupt:
            sys.exit()

    print(f"\n>> ターゲット: {selected_name} を選択しました。")
    print("終了: 'exit' または 'quit' / インデックス手動更新: 'update'")

    index_file = get_index_path(selected_path)

    # コマンドライン引数実行時の処理
    if len(sys.argv) > 1:
        arg = sys.argv[1]
        if arg.lower() == "update":
            build_index(selected_path, index_file)
        else:
            search_local(arg, index_file, selected_path)
        return

    # インタラクティブループ
    while True:
        try:
            word = input(f"\n[{selected_name}] 検索キーワード > ").strip()
            if not word:
                continue
            if word.lower() in ['exit', 'quit']:
                break
            elif word.lower() == 'update':
                build_index(selected_path, index_file)
            else:
                search_local(word, index_file, selected_path)
        except KeyboardInterrupt:
            break

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("\n" + "="*50)
        print("【予期せぬエラーが発生しました】")
        traceback.print_exc()
        print("="*50)
        input("\nEnterキーを押して終了してください...")