import pandas as pd
from pathlib import Path
import logging

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class DataCleaningPipeline:
    '''
    月次データなどの生データ（CSV）を読み込み、
    欠損値の処理、重複排除、データ型の正規化などのクリーンアップを自動実行するツール。
    '''
    def __init__(self, input_dir: str, output_dir: str):
        self.input_dir = Path(input_dir)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def process_pipeline(self) -> None:
        if not self.input_dir.exists():
            return
        for file_path in self.input_dir.glob("*.csv"):
            df = pd.read_csv(file_path, encoding='utf-8-sig', on_bad_lines='skip')
            df = df.drop_duplicates().dropna(how='all')
            df.to_csv(self.output_dir / f"Cleaned_{file_path.name}", index=False, encoding='utf-8-sig')
            logging.info(f"Processed: {file_path.name}")

if __name__ == "__main__":
    base_path = Path(__file__).parent
    pipeline = DataCleaningPipeline(str(base_path / "inputs"), str(base_path / "outputs"))
    pipeline.process_pipeline()
