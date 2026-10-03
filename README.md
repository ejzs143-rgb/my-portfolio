# Kensuke Ishioka — 業務自動化ポートフォリオ

自動車アフターサービス領域の実務で、手作業の集計・転記・資料作成を自動化してきた経験をもとに、
**汎用化して公開できる部分**をまとめたリポジトリです。
掲載コードはすべて一般化済みで、実務データ・社内固有の情報は含みません（[公開前チェック](#09_leak_guard--公開前の機密情報チェック)を参照）。

## 掲載ツール

| フォルダ | 内容 | 技術 |
|---|---|---|
| [01_meeting_minutes](01_meeting_minutes) | 会議ログ（JSON）から決定事項・TODO を抽出し、Markdown 議事録を生成 | Python |
| [02_excel_aggregator](02_excel_aggregator) | 複数の Excel を統合し、ファイル別の集計シート付きレポートを出力 | Python / pandas |
| [03_data_cleaning](03_data_cleaning) | CSV の重複・空行除去と型の正規化を一括実行 | Python / pandas |
| [04_excel_vba](04_excel_vba) | 全シートの件数を集計したサマリーシートを自動作成 | VBA |
| [05_text_grouping](05_text_grouping) | 日本語の短文コメントを類似度でグルーピングし、件数・構成比を集計 | Python（標準ライブラリのみ） |
| [06_document_tools](06_document_tools) | Excel→PDF 一括変換、PDF→テキスト抽出、PDF 内画像の圧縮、画像圧縮 | Python / pypdf / Pillow / COM |
| [07_media_tools](07_media_tools) | 動画・音声のトリミング GUI、文字起こし、WAV 変換、動画圧縮 | Python / FFmpeg / faster-whisper |
| [08_xml_merge](08_xml_merge) | 複数 XML ファイルを GUI で選択して 1 ファイルに統合 | PowerShell（バッチ埋め込み） |
| [09_leak_guard](09_leak_guard) | コミット／プッシュ前に機密情報の混入を検出して止める Git フック | Python / Git hooks |

### 05_text_grouping — 日本語コメントの類似グルーピング
大量の自由記述コメントを「同じ趣旨」ごとにまとめ、どの指摘が多いかを数える用途で作成しました。
文字列一致率・文字 bigram の Jaccard 係数・キーワード一致度（同義語の正規化つき）を加重平均し、
1 パス目で逐次クラスタリング、2 パス目で小グループ同士を再マージします。外部ライブラリは使っていません。

```
python 05_text_grouping/comment_grouper.py 05_text_grouping/sample_input.csv result.csv
```

### 09_leak_guard — 公開前の機密情報チェック
公開リポジトリへの業務情報の混入を、仕組みで防ぐためのツールです。
- 組み込みルール：メールアドレス、Windows のユーザーフォルダ、UNC パス、社内 IP、API キー・秘密鍵、パスワード直書き
- 禁止語リスト：社名・人名・社内システム名などを利用者が登録（リストはリポジトリ外に置き、公開しない）
- 中身を検査できない文書ファイル（Excel / Word / PDF / メール等）はコミット自体を止める
- `pre-commit` で変更内容を、`pre-push` で送信するコミットの内容・メッセージ・作成者メールを検査
- `--history` で過去の全コミットも検査可能。検出値は画面上でも伏字にして二次流出を防ぐ

```
09_leak_guard\install_hooks.bat          :: フックを有効化（初回のみ）
python 09_leak_guard\leak_guard.py --all   :: 手動で全ファイル検査
```

## 実務で開発したツール（社内業務のためコード非公開）
- **月次指標レポートの自動作成（Python）**：複数システムの月次エクスポート（Excel / Access 出力）を拠点コードで突合し、月次レポートを生成。手作業の参照式に依存していた運用を、誰が実行しても同じ結果になる計算ロジックに置き換え、前回結果との差異を要因別に説明できる形にした
- **月次品質データの取り込みツール（Python）**：入力ファイルの妥当性検証と、マスタとの差分（新規・廃止拠点など）の検知つき
- **報告資料の一括生成（VBA）**：集計結果から拠点別の PowerPoint 資料を自動作成
- **業務時間ログの集計・報告書作成（Python）**：作業ログを業務区分に分類し、報告用の Excel を生成

## 動作環境
- Windows 10 / 11、Python 3.10 以上（`pip install -r requirements.txt`）
- Excel→PDF 変換は Microsoft Office、動画系ツールは FFmpeg が必要
- `.bas` ファイルは表示用に UTF-8 で保存しています。VBA エディタへインポートする場合は Shift_JIS に変換してください

## 技術スタック
Python（pandas / openpyxl / pypdf / Pillow / tkinter）、VBA、PowerShell、Windows Batch、FFmpeg、Git
