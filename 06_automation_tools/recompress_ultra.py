import zipfile
import os
import tempfile
from pathlib import Path

def recompress_lzma():
    print("=== 高圧縮ZIP生成ツール (LZMA版) ===")
    
    # ユーザーに入力を求める（ドラッグ&ドロップの代わりにパスを貼り付け）
    input_str = input("圧縮したいZIPファイルをここに貼り付けてEnterを押してください: ").strip('"')
    input_path = Path(input_str)

    if not input_path.exists() or input_path.suffix.lower() != '.zip':
        print(f"\n【エラー】指定されたファイルがZIPではありません: {input_str}")
        input("\n何かキーを押すと終了します...")
        return

    output_path = input_path.parent / f"(ultra)_{input_path.name}"

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            print(f"\n1/2: 解凍中...")
            with zipfile.ZipFile(input_path, 'r') as zip_read:
                zip_read.extractall(tmpdir)
            
            print(f"2/2: LZMA方式で再圧縮中（時間がかかります）...")
            with zipfile.ZipFile(output_path, 'w', compression=zipfile.ZIP_LZMA) as zip_write:
                for root, dirs, files in os.walk(tmpdir):
                    for file in files:
                        file_full_path = Path(root) / file
                        arcname = file_full_path.relative_to(tmpdir)
                        zip_write.write(file_full_path, arcname)

        old_size = input_path.stat().st_size
        new_size = output_path.stat().st_size
        print(f"\n--- 完了 ---")
        print(f"出力先: {output_path}")
        print(f"サイズ: {old_size:,} -> {new_size:,} bytes")
        print(f"削減率: {(1 - new_size/old_size)*100:.2f}%")

    except Exception as e:
        print(f"\n【実行エラー】\n{e}")

    input("\n処理が終了しました。Enterキーを押して閉じてください...")

if __name__ == "__main__":
    recompress_lzma()
