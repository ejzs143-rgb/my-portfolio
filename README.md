# Kensuke Ishioka - Data Engineering & Automation Portfolio

## 概要 (Overview)
自動車業界のアフターセールス・保証業務（Warranty / Field Force）領域における、実務課題を解決するための自動化・データ分析ポートフォリオです。
バックオフィス業務のDX推進、データパイプライン構築、レガシー環境からの脱却を目的とした自作ツール群をまとめています。

## プロジェクト一覧 (Projects)

### 1. 会議運営・議事録自動化システム (`01_meeting_automation`)
- **技術**: Python
- **概要**: 会議のテキストログから決定事項やアクションアイテム（TODO）を自動抽出し、構造化されたMarkdown形式の議事録を生成。業務フローのシステム化を実現。

### 2. 総合レポート集約基盤 (`02_operation_report_builder`)
- **技術**: Python (pandas)
- **概要**: 複数拠点や部門から提出される散在したExcelデータを自動で読み込み、統合・集計。経営・現場向けのサマリーレポートを生成し、集計工数を大幅に削減。

### 3. 月次データETLパイプライン (`03_monthly_data_pipeline`)
- **技術**: Python (pandas)
- **概要**: 月次で発生する大量の生データに対し、欠損値の補完、重複排除、データ型の正規化などのクレンジング処理を自動実行。分析可能な高品質データへ変換するバッチ処理。

### 4. 高度Excel業務自動化マクロ (`04_excel_vba_automation`)
- **技術**: VBA
- **概要**: モダンな環境が即座に導入できない現場向けに、複数シートからの条件抽出やデータ転記を完全自動化。現場に即導入可能な実用的マクロ。

## 技術スタック (Tech Stack)
- **Languages**: Python, VBA, PowerShell
- **Libraries**: pandas, json, pathlib
- **Tools**: Git, GitHub, Excel, Power BI



### 新規追加: 自動化ツール群 (`06_automation_tools`)
- **技術**: Python, VBA, PowerShell, Windows Batch, JavaScript
- **概要**: ファイル変換、データ集計、メディア処理、業務レポート作成などの自動化ツール集。公開用に固有名詞、個人情報、社内パス、認証情報を一般化し、第三者配布物は除外注記へ置換。




### 新規追加: 自動化ツール群 (`06_automation_tools`)
- **技術**: Python, VBA, PowerShell, Windows Batch, JavaScript, FFmpeg, openpyxl
- **概要**: 会議・データ集計・ファイル変換・メディア処理・業務レポート作成を支援する汎用自動化ツール集。固有名詞、個人名、メールアドレス、ローカルパス、サーバー名を公開用の一般表現へ置換済みです。
