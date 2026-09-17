import pandas as pd
from pathlib import Path
from datetime import datetime

class ExcelDataAggregator:
    "\""
    複数のExcelファイルを読み込み、データを統合・集計して
    サマリーレポート（統合Excel）を自動生成する汎用ツール。
    "\""
    def __init__(self, input_dir: str, output_dir: str):
        self.input_dir = Path(input_dir)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def process_files(self) -> None:
        if not self.input_dir.exists():
            print(f"入力ディレクトリが存在しません: {self.input_dir}")
            return

        all_data = []
        for file_path in self.input_dir.glob("*.xlsx"):
            try:
                df = pd.read_excel(file_path)
                # データの出所を追跡できるようにファイル名を追加
                df['Source_File'] = file_path.stem
                all_data.append(df)
                print(f"読み込み成功: {file_path.name} ({len(df)}行)")
            except Exception as e:
                print(f"エラー発生 ({file_path.name}): {e}")

        if not all_data:
            print("集計対象のデータが見つかりませんでした。")
            return

        # 全データの結合 (縦積み)
        merged_df = pd.concat(all_data, ignore_index=True)
        self._generate_report(merged_df)

    def _generate_report(self, df: pd.DataFrame) -> None:
        # 出力ファイル名の生成
        date_str = datetime.now().strftime("%Y%m%d")
        output_file = self.output_dir / f"Consolidated_Report_{date_str}.xlsx"
        
        # 数値列を抽出してファイルごとに合計を集計するサマリーロジック
        numeric_cols = df.select_dtypes(include='number').columns
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 1. 生データの結合シート
            df.to_excel(writer, sheet_name="All_Data", index=False)
            
            # 2. サマリー集計シート (数値データがある場合のみ)
            if len(numeric_cols) > 0 and 'Source_File' in df.columns:
                summary_df = df.groupby('Source_File')[numeric_cols].sum().reset_index()
                summary_df.to_excel(writer, sheet_name="Summary_by_Source", index=False)
                
        print(f"集計レポートを出力しました: {output_file}")

if __name__ == "__main__":
    base_path = Path(__file__).parent
    input_path = base_path / "inputs"
    output_path = base_path / "outputs"
    
    aggregator = ExcelDataAggregator(str(input_path), str(output_path))
    aggregator.process_files()