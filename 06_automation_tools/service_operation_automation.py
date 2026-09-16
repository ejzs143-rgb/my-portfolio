# -*- coding: utf-8 -*-
"""
ServiceMetrics 月次レポート自動作成ツール
================================================
Example Company / FF課 担当者向け(最終版)

■設計思想
  前任担当者のお手本ファイルは「特定フォルダに各種Excelが揃っている前提」の
  外部リンク数式で組まれているが、本ツールはその前提を引き継がない。
  Pythonが各ソースファイルから直接値を計算し、数式に依存しない
  「静的な値だけの最終形」を出力する。出力ファイルには他環境のファイルパスを
  参照する数式を一切残さない(過去月シート分も含めて値に固定する)。

■できること
  実行時に2つのモードを選べる:
    (a) 今ある最新月のシートが未完成なら、そのシートを完成させる(全データ更新)
    (b) 最新月が完成済みなら、次の月のシートを新規作成する
  ReceptionMetrics(ダイアログレセプション/H列)は、入手済みなら(a)(b)いずれの中でも
  そのまま反映できる(未入手なら空欄のまま出力し、後日同じ「全データ更新」を
  再実行すれば良い。以前あった「H列だけ追加」専用モードは、ReceptionMetricsを毎回
  入手できる前提になったため廃止した)。
■①(基準ファイル)についての重要な原則
  ①には前任担当者/担当者が継続管理する系譜ファイル(前回の出力)のみを使う。
  データ提供者から届くファイルを①に使ってはいけない
  (データ提供者は自分の手元のコピーにMediaCaptureを追記して返送してくるだけで、
  彼の店舗リスト自体が最新・網羅的である保証がないため。実際にデータ提供者
  経由のファイルが248店舗しかなく、276店舗版と混同する事故が発生した)。
  MediaCapture(I・J列)は、データ提供者から届くファイルから⑥で店舗コードを
  キーに値だけを取り込み、①の店舗リストにマージする。①のファイル自体は
  店舗数の基準として常に前任担当者系譜を使い続ける。

  いずれの場合も:
    - 全シートの外部リンク数式を「保存済みの計算値」に置き換えて静的化する
    - TTL(合計)行はPythonが再計算して静的な値で書き込む
    - K・M(目標値)は前月シート(または指定した参照ファイル)の値を
      静的な数値として引き継ぐ(目標値はいじらない)

■手動が残る部分(構造上避けられないもの)
  1. 各ソースファイルの取得そのもの
     UsagePortal(GRP/ビッフィー)・LegacyServiceTool・NewServiceTool・QualityMetricsは別システムからの
     エクスポートが必要。ログイン権限・操作はPythonでは代替できない。
  2. MediaCapture(I・J列)
     データ提供者がServiceMetricsワークブック自体にMediaCaptureデータを
     入力した状態でメール送付してくる運用。つまり①で選ぶベースファイルに
     最初から入っているのが通常形。⑥のダイアログは、別途「利用率集計」
     ファイルから取り込みたい場合の予備手段(通常は「いいえ」でよい)。
  3. ReceptionMetrics(H列)
     次のどちらのファイルでも読める(2026年7月分で数値照合済み)。
       (a) 「ReceptionMetrics [YYYYMM].xlsx」の「[YYYYMM]DR」シート(B列=店舗コード,
           E列='*'列)…従来形式
       (b) 「[YYYYMM]_T_Jisseki_su_sum_2.xlsx」…(a)の元になっている
           Access出力そのもの。ReceptionMetricsブックへの貼り付け工程が不要になる
     いずれも見出し名で列を特定するため、列位置のズレやシート名末尾の
     空白では失敗しない。月初時点では存在せず、翌週月曜に取得可能。
     モード(b)で後から追加する。

■確定済みのデータソース対応(2026年6月実績で数値照合済み)
  F列 UsagePortal利用数            = UsagePortal Usage Report(VW+商用車)「# VIN Requests」合算
  G列 メンテナンステーブル   = LegacyServiceTool「Created MTs Different VINs」
                              + NewServiceTool「MT: Tables created with unique Vins」
                              ※複数ブランド(VW/商用車/Audi等)を持つ店舗は
                              全ブランド行を合算する(下記【運用方針】参照)
  H列 ダイアログレセプション = ReceptionMetrics [YYYYMM].xlsx の '*' 列、または
                              [YYYYMM]_T_Jisseki_su_sum_2.xlsx(Access出力)の
                              「CHECKOUT_YMDのカウント*」列。両者は同一値。
                              K列(カウントK)は前任担当者仕様どおり取らない
  I列 MediaCaptureテク撮影数   = データ提供者がベースファイルに入力済み
  J列 MediaCaptureアド送信数   = I列と同値(暫定仕様。下記【運用方針】参照)
  K・M列 目標                = 前月(または参照ファイル)から静的値で引き継ぎ
  L・N列 QualityMetrics実績            = QualityMetrics実績シートの当月列(Q=G〜R, 最終=U〜AF)
                              ※QualityMetrics実績に存在しない店舗コードは「非計測店舗」
                              として扱う(下記【運用方針】参照)

■運用方針(2026年7月、前任担当者からの個別確認を前提とせず担当者の判断で確定)
  担当者が単独運用に移行するにあたり、前任担当者の手作業(の再現・都度確認)
  に依存しない、独立して説明可能な計算方式を正式な運用ルールとする。
  前任者(前任担当者)の集計結果との差異は「バグ」ではなく、下記の意図的な
  設計判断による既知の差異として扱う。

  1) G列は全ブランド機械合算(V+N)とする(2026/6実績で+67/276店舗差・0.28%)
     根拠(2026/7/9 前任担当者本人への確認と法医学的解析で確定):
     ・生データ(LegacyServiceTool/NewServiceTool)は前任担当者と完全同一(ハッシュ一致を確認済み)。
       差はデータではなく「集計のやり方」だけから生じている。
     ・前任担当者の実際の手順(本人証言): NewServiceToolをV/A/Nに手作業で仕分け→
       Audiを除外→VとNを別々にVLOOKUPで表へ移して合算。「面倒くさい」
       「本当はやりたくない」と本人も認める多段階の手作業。
     ・差異91店舗の逆算分解: ①Vが丸ごと抜けている=37件(H+Nのみで天野
       値と一致) ②NewServiceTool寄与ゼロなのに+1=11件(IFERROR既定値の疑い)
       ③同一入力なのに異なる出力=43件(店舗ごとの手作業補正の跡。
       数式では再現不能)。→手作業ミスの蓄積が差の正体。
     ・本ツールの全ブランド機械合算は、再現性・正確性ともに手作業より
       優れ、公式配布ファイルのみから誰でも同じ結果を再現できる
       (監査・引き継ぎ耐性が高い)。よってこちらを正式方式とする。

  2) J列(アド送信数)はI列(テク撮影数)と同値のまま運用する
     根拠: 前任担当者の旧ファイルでも実際にI列と全店舗で完全一致しており
     (独自の値が入っていた形跡なし)、現行運用と実質差はない。

  3) QualityMetrics実績に存在しない店舗(2026/6時点で徳島86010のみ)は、
     前月引き継ぎの静的値をそのまま「非計測店舗の暫定値」として扱う
     (前任担当者の旧ファイルでも6ヶ月間 L=N=4 のまま更新されておらず、
     実測値ではなく同様の暫定運用だったことを確認済み)。
     QualityMetrics側で計測対象化されない限り、この店舗の実績は変動しない前提。

  ※将来、前任担当者から追加情報(合算方針の公式回答・別データソースの共有)
  があれば、上記1)を再考する。ただし現行の全ブランド合算方式は、それ単体
  で独立して正当化可能であるため、情報が得られないことは業務停止の理由
  にならない。

■使い方
  run_service_operation.bat をダブルクリック → ダイアログに従って
  ファイルを選ぶだけ。実行後に反映件数レポートが表示される。
"""

import csv
import json
import os
import re
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter


# ============================================================
# 1. 各データソースの読み込み
# ============================================================

def load_ppso(paths):
    """UsagePortal Usage Reportから店舗コード別のUsagePortal利用数(単月)を集計する。
    Partner列 "981/V/51020" 形式から5桁コードを抽出し、
    "# VIN Requests"(G列)を合算する。VW・商用車の2ファイルを渡すと合算される。
    """
    rows = {}
    for path in paths:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb['DS_DEALER_DETAIL_ONE_LINE']
        for r in range(2, ws.max_row + 1):
            partner = ws[f'A{r}'].value
            vin_req = ws[f'G{r}'].value
            if partner is None:
                continue
            m = re.search(r'/(\d{5})$', str(partner))
            if not m:
                continue
            code = int(m.group(1))
            try:
                val = float(vin_req)
            except (TypeError, ValueError):
                val = 0.0
            rows[code] = rows.get(code, 0.0) + val
        wb.close()
    return rows


def load_maintenance_table(elsapro_path, elsa2go_path):
    """メンテナンステーブル利用数を店舗コード別に集計する。
    LegacyServiceTool「Created MTs Different VINs」(H列・ユニークVIN数)と、
    NewServiceTool「MT: Tables created with unique Vins」を店舗コードで合算する。
    ※「Created MTs」(重複VIN込み)ではなくDifferent VINsの方が
      実績との一致度が高かったためこちらを採用(2026/6分で検証済み)。
    NewServiceToolのCSVは先頭列名が毎月変わる('6-2026'等)ため位置で参照し、
    集計列は固定の列名で参照する。末尾の'Sum'行は除外する。
    """
    rows = {}
    wb = openpyxl.load_workbook(elsapro_path, data_only=True)
    ws = wb['Bericht 1']
    for r in range(10, ws.max_row + 1):
        b = ws[f'B{r}'].value
        h = ws[f'H{r}'].value  # Created MTs Different VINs
        if b is None:
            continue
        bs = str(b).strip()
        if not (bs.isdigit() and len(bs) == 5):
            continue  # 合計行・不正コードを除外
        code = int(bs)
        val = float(h) if isinstance(h, (int, float)) else 0.0
        rows[code] = rows.get(code, 0.0) + val
    wb.close()

    with open(elsa2go_path, newline='', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        header = next(reader)
        try:
            col_idx = header.index('MT: Tables created with unique Vins')
        except ValueError:
            raise ValueError(
                'NewServiceTool CSVに「MT: Tables created with unique Vins」列が見つかりません。'
                'ファイルが正しいか確認してください。')
        for row in reader:
            if not row or row[0] == 'Sum':
                continue
            m = re.match(r'JPN(\d{5})', str(row[0]))
            if not m:
                continue
            code = int(m.group(1))
            try:
                val = float(row[col_idx])
            except (TypeError, ValueError, IndexError):
                val = 0.0
            rows[code] = rows.get(code, 0.0) + val  # ブランド違いの重複コードも合算
    return rows


def load_service_cam_from_so_file(path, sheet_name):
    """データ提供者から届くファイル(ServiceMetricsワークブック形式)の該当月
    シートから、店舗コード(B列)別にMediaCapture(I・J列)を取り込む。

    【設計判断】①(店舗リストの基準)には前任担当者/担当者が継続管理する
    系譜ファイルのみを使い、データ提供者のファイルを①に使うことはしない
    (データ提供者側の店舗リストが古い可能性があり、店舗数の基準にすべきでは
    ないため)。MediaCaptureの値だけを店舗コードで拾い、①の店舗リストに
    マージする。path未指定・シート無しなら空を返す(=未反映のまま進める)。
    """
    rows = {}
    if not path or not sheet_name:
        return rows
    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        return rows
    ws = wb[sheet_name]
    for r in range(2, ws.max_row + 1):
        code = ws[f'B{r}'].value
        if code is None or ws[f'C{r}'].value == 'TTL':
            continue
        i_val = ws[f'I{r}'].value
        j_val = ws[f'J{r}'].value
        rows[int(code)] = (
            float(i_val) if isinstance(i_val, (int, float)) else 0.0,
            float(j_val) if isinstance(j_val, (int, float)) else 0.0,
        )
    wb.close()
    return rows


def load_service_cam(path, sheet_name):
    """(旧形式向け)VW_Service_Cam_利用率集計ファイルの指定月シートから
    撮影数(G列)を取得する。データ提供者から「利用率集計」形式そのものが
    届いた場合の予備手段。path未指定・シート無しなら空を返す。
    """
    rows = {}
    if not path or not sheet_name:
        return rows
    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        return rows
    ws = wb[sheet_name]
    for r in range(2, ws.max_row + 1):
        code = ws[f'A{r}'].value
        photo = ws[f'G{r}'].value
        if code is None:
            continue
        val = float(photo) if isinstance(photo, (int, float)) else 0.0
        rows[int(code)] = val
    wb.close()
    return rows


def _col_letter(base_letter, month):
    """base_letter(1月の列)からmonth-1列だけ右の列名。例: ('G',6)->'L'"""
    return get_column_letter(column_index_from_string(base_letter) + month - 1)


def load_msqp(path, month):
    """QualityMetrics実績シートから当月のQチェック/最終チェック「実績」を取得する。
    Qチェック実績: G(1月)〜R(12月) / 最終チェック実績: U(1月)〜AF(12月)。
    目標(F・T列)はここでは扱わない(K・M列は前月からの静的引き継ぎのため)。
    """
    rows = {}
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb['QualityMetrics実績']
    q_col = _col_letter('G', month)
    f_col = _col_letter('U', month)
    for r in range(6, ws.max_row + 1):
        code = ws[f'B{r}'].value
        if code is None:
            continue
        rows[int(code)] = {
            'q_actual': ws[f'{q_col}{r}'].value,
            'final_actual': ws[f'{f_col}{r}'].value,
        }
    wb.close()
    return rows


def _find_sheet(wb, sheet_name):
    """シート名を「前後の空白を無視して」探す。
    ※前任担当者作成の ReceptionMetrics 202607.xlsx はシート名が「202607DR 」と
      末尾に半角スペースが入っていた。完全一致だと見つからず、
      エラーにならないままH列が空で通る(静かな失敗)ため、緩く探す。
    """
    if not sheet_name:
        return None
    target = str(sheet_name).strip().lower()
    for name in wb.sheetnames:
        if str(name).strip().lower() == target:
            return name
    return None


def _dr_header_map(ws, scan_rows=5):
    """ダイアログレセプション表のヘッダー行を探し、
    {'row':見出し行, 'code':店舗コード列, 'value':値列, 'ym':YYYYMM列} を返す。

    2形式に共通対応する:
      (1) ReceptionMetrics [YYYYMM].xlsx の [YYYYMM]DR シート
          → 2行目が見出し、B列=DEALER_CODE / E列='*' / G列=YYYYMM
      (2) Access出力 [YYYYMM]_T_Jisseki_su_sum_2.xlsx
          → 1行目が見出し、A列=DEALER_CODE / D列='CHECKOUT_YMDのカウント*'
             / F列=YYYYMM
    見出し名で列を特定するため、列位置がずれても追随できる。
    """
    for r in range(1, min(scan_rows, ws.max_row) + 1):
        code_col = value_col = ym_col = None
        for c in range(1, ws.max_column + 1):
            v = ws.cell(r, c).value
            if v is None:
                continue
            s = str(v).strip()
            up = s.upper()
            if code_col is None and up in ('DEALER_CODE', 'DEALERCODE', '販売店コード'):
                code_col = c
            elif value_col is None and (s == '*' or s.endswith('カウント*')):
                value_col = c
            elif ym_col is None and up in ('YYYYMM', 'YYYY/MM'):
                ym_col = c
        if code_col and value_col:
            return {'row': r, 'code': code_col, 'value': value_col, 'ym': ym_col}
    return None


def load_dialog_reception(path, sheet_name, expect_yyyymm=None):
    """ダイアログレセプション数(H列の元データ)を店舗コード別に取得する。

    受け付けるファイル(どちらを選んでも同じ結果になる):
      (a) ReceptionMetrics [YYYYMM].xlsx  … 従来どおり。[YYYYMM]DRシートを読む
      (b) [YYYYMM]_T_Jisseki_su_sum_2.xlsx … ReceptionMetricsの元になっているAccess出力。
          ReceptionMetricsブックへの貼り付け工程を経ずに直接読める

    取得する値は「CHECKOUT_YMDのカウント*」列(ReceptionMetrics上の '*' 列)。
    ※前任担当者お手本内の数式
      =VLOOKUP($B2,'[ReceptionMetrics 202605.xlsx]202605DR'!$B$3:$E$242,4,FALSE)
      と同じ列。K列(D列)は取らない仕様をそのまま踏襲する。

    expect_yyyymm('202607'等)を渡すと、対象月が違うファイルを掴んだ場合に
    空の辞書を返す(古い月のファイルで静かに上書きする事故を防ぐ)。
    """
    rows = {}
    if not path:
        return rows
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        name = _find_sheet(wb, sheet_name)
        candidates = [name] if name else list(wb.sheetnames)
        for cand in candidates:
            ws = wb[cand]
            hdr = _dr_header_map(ws)
            if not hdr:
                continue
            found = {}
            months = set()
            for r in range(hdr['row'] + 1, ws.max_row + 1):
                code = ws.cell(r, hdr['code']).value
                val = ws.cell(r, hdr['value']).value
                if code is None:
                    continue
                try:
                    found[int(code)] = float(val) if isinstance(val, (int, float)) else 0.0
                except (TypeError, ValueError):
                    continue
                if hdr['ym']:
                    ym = ws.cell(r, hdr['ym']).value
                    if ym is not None:
                        months.add(str(ym).strip())
            if not found:
                continue
            if expect_yyyymm and months and str(expect_yyyymm) not in months:
                continue  # 対象月が違う → 採用しない
            return found
    finally:
        wb.close()
    return rows


def _check_dialog_reception(config):
    """ReceptionMetrics / Access出力のどちらでも、当月データが実際に読めるかを検証する。
    シート名の一致ではなく「読めた店舗数」で判定するため、
    シート名末尾の空白や、ファイル形式の違いでは落ちない。
    """
    path = config.get('ro_dr_path')
    if not path:
        return None
    try:
        dr = load_dialog_reception(path, config.get('ro_dr_sheet'),
                                   expect_yyyymm=config.get('write_sheet'))
    except Exception as e:
        return (f"★ReceptionMetrics: ファイルを開けません({e})", True)
    if dr:
        return (f"・ReceptionMetrics: {config['write_sheet']}のデータ{len(dr)}店舗 読み取りOK", False)
    return ("★ReceptionMetrics: 当月データを読み取れません"
            f"(対象月{config['write_sheet']}のシート/行が見つからない可能性)", True)


def load_targets(path, sheet_name):
    """指定ワークブックの指定シートから、店舗コード別のK・M(目標値)の
    「計算済みの値」を取得する。数式そのものではなく必ず値を読む。"""
    rows = {}
    if not path or not sheet_name:
        return rows
    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        return rows
    ws = wb[sheet_name]
    for r in range(2, ws.max_row + 1):
        code = ws[f'B{r}'].value
        c_val = ws[f'C{r}'].value
        if code is None or c_val == 'TTL':
            continue
        k = ws[f'K{r}'].value
        m = ws[f'M{r}'].value
        if k in (None, '') and m in (None, ''):
            continue
        rows[int(code)] = (k, m)
    wb.close()
    return rows


# ============================================================
# 2. ワークブックの静的化・TTL再計算
# ============================================================

_EXTERNAL_REF = re.compile(r'\[\d+\]')


def freeze_external_formulas(wb, wb_values):
    """全シートを走査し、外部ファイル参照([n]付き)の数式セルを
    「保存済みの計算値」に置き換えて静的化する。
    内部数式(=SUM等、[n]を含まないもの)はここでは触らない。

    【なぜ必要か】openpyxlはファイルを読み込んで保存し直すと、数式セルの
    「キャッシュ済み計算値」を失う。前任担当者お手本には過去月シート含め
    1.4万セル超の外部リンク数式があり、放置すると過去月のH・K・M列等が
    担当者の環境で全て空欄・リンク切れになる(実測で確認済み)。
    そのため保存前に、元ファイルに保存されていた計算値でセルを置き換える。
    """
    frozen = 0
    for sn in wb.sheetnames:
        if sn not in wb_values.sheetnames:
            continue
        ws = wb[sn]
        ws_v = wb_values[sn]
        for row in ws.iter_rows():
            for cell in row:
                if (cell.data_type == 'f'
                        and isinstance(cell.value, str)
                        and _EXTERNAL_REF.search(cell.value)):
                    cell.value = ws_v[cell.coordinate].value
                    frozen += 1
    return frozen


def find_ttl_row(ws):
    """C列が'TTL'の行番号を返す(無ければNone)。
    前任担当者お手本はB列空欄・作業用ファイルはB列99999だが、
    C列'TTL'は両ファイル共通のため、これで判定する。"""
    for r in range(2, ws.max_row + 1):
        if ws[f'C{r}'].value == 'TTL':
            return r
    return None


def recompute_ttl(ws):
    """TTL(合計)行のF〜N列を、店舗行の合計から静的な値で再計算する。
    元のTTL行は=SUM数式だが、静的な最終形として値で確定させる。
    列内に1件もデータが無い場合は合計も空欄にする(未取得の明示)。"""
    ttl_r = find_ttl_row(ws)
    if ttl_r is None:
        return None
    totals = {}
    for col in 'FGHIJKLMN':
        vals = []
        for r in range(2, ws.max_row + 1):
            if r == ttl_r or ws[f'B{r}'].value is None or ws[f'C{r}'].value == 'TTL':
                continue
            v = ws[f'{col}{r}'].value
            if isinstance(v, (int, float)):
                vals.append(v)
        total = sum(vals) if vals else None
        ws[f'{col}{ttl_r}'] = total
        totals[col] = total
    return totals


# ============================================================
# 3.5 ディーラーマスタ照合(閉店・改称の検知と反映)
# ============================================================
# 【経緯】86010(閉店)・86300(改称)のような変化は、放置すると来月以降
# 誤ったまま自動継続してしまう。マスタと現在のシートを突合し、
# 差異が見つかった項目だけを対話的に確認して反映する。
# 新規追加候補は件数が多くなりやすいため、個別ダイアログにはせず
# レポートに一覧化し、追加自体は次回別途行う(閉店・改称は削除/更新の
# ワンアクションだが、新規追加は行データの作成が必要で性質が異なるため)。

def load_dealer_master(path):
    """VW販売店マスターファイルから、有効な店舗コード→店名のdictを作る。
    「店舗マスター」(新車店舗コード/専売店舗名)と「サテライトマスター」
    (サテライト拠点コード/サテライト拠点名称)の両シートを結合する。
    シートが無い/形式が違う場合はそのシートだけ無視し、処理は継続する。
    """
    master = {}
    wb = openpyxl.load_workbook(path, data_only=True)

    def scan(sheet_name, code_col, name_col, header_row=2, start_row=3):
        if sheet_name not in wb.sheetnames:
            return
        ws = wb[sheet_name]
        for r in range(start_row, ws.max_row + 1):
            code = ws[f'{code_col}{r}'].value
            name = ws[f'{name_col}{r}'].value
            if isinstance(code, int):
                master[code] = name

    scan('店舗マスター', 'D', 'E')
    scan('サテライトマスター', 'C', 'D')
    wb.close()
    return master


def _normalize_name(name):
    """店名比較用に、ブランド接頭辞や全角/半角スペースの違いを吸収する。"""
    if not name:
        return ''
    s = str(name).replace('　', ' ').replace('Volkswagen', '').strip()
    return re.sub(r'\s+', '', s)


def reconcile_dealer_master(ws, master):
    """シートの店舗コード(B列)・店名(C列)をマスタと突合する。
    閉店候補は1件ずつダイアログで確認して削除を反映する。
    改称候補・新規追加候補は反映せず一覧としてレポートに残す(別途手動対応)。
    ※改称候補を個別ダイアログにしなかった理由: ServiceMetricsは
    「Volkswagen北見」のような略称、マスタは「Volkswagen（旭川）北見認定
    中古車センター」のような正式名称、という慣習差が大量にあり(検証時に
    276店舗中21件が該当)、個別確認では実質ノイズになる。店名の一致判定は
    人が一覧を見て判断する方が確実。

    戻り値: {'閉店として削除': [...], '店名差分(未反映・要確認)': [...],
             '新規追加候補(未反映・要確認)': [...]}
    """
    result = {
        '閉店として削除': [],
        '店名差分(未反映・要確認)': [],
        '新規追加候補(未反映・要確認)': [],
    }

    so_rows = {}
    for r in range(2, ws.max_row + 1):
        b = ws[f'B{r}'].value
        if b is None or ws[f'C{r}'].value == 'TTL':
            continue
        so_rows[int(b)] = r

    # 閉店候補: シートにあるがマスタに無いコード(=確実な信号のみ)。
    # 以前は1件ずつダイアログで確認していたが、件数が増えると確認作業が
    # 長くなるため、候補を一括表示して1回のYes/Noで反映するようにした。
    closure_candidates = []
    for code, r in sorted(so_rows.items()):
        if code not in master:
            closure_candidates.append((code, r, ws[f'C{r}'].value))

    if closure_candidates:
        listing = '\n'.join(f'  {c}: {n}' for c, _, n in closure_candidates[:20])
        more = f'\n  ...他{len(closure_candidates)-20}件' if len(closure_candidates) > 20 else ''
        if messagebox.askyesno(
                'ディーラーマスタ差異: 閉店の疑い(まとめて確認)',
                f'マスタに見当たらない店舗が{len(closure_candidates)}件あります。\n'
                f'{listing}{more}\n\n'
                '閉店等とみなして、まとめてこのシートから削除しますか?\n'
                '(「いいえ」の場合、全件そのまま残して記録のみ行います)'):
            for code, _, name in closure_candidates:
                result['閉店として削除'].append((code, name))
            for _, r, _ in sorted(closure_candidates, key=lambda x: -x[1]):
                ws.delete_rows(r, 1)
                so_rows = {c: (rr - 1 if rr > r else rr) for c, rr in so_rows.items() if rr != r}

    # 店名差分: 一覧化のみ(自動反映しない)。マスタは略称/正式名称の慣習差が
    # 大きく、機械的な自動更新は誤爆リスクが高いため、人の目での確認に委ねる。
    for code, r in sorted(so_rows.items()):
        if code not in master:
            continue
        old_name = ws[f'C{r}'].value
        new_name = master[code]
        if not new_name or _normalize_name(old_name) == _normalize_name(new_name):
            continue
        result['店名差分(未反映・要確認)'].append((code, old_name, new_name))

    # 新規追加候補: マスタにあるがシートに無いコード(件数が多くなりうるため一覧化のみ)
    missing = sorted(set(master) - set(so_rows))
    result['新規追加候補(未反映・要確認)'] = [(c, master[c]) for c in missing]

    return result

def build_month_sheet(config):
    """configに従いServiceMetricsワークブックを更新し、レポートdictを返す。

    モード:
      overwrite_existing=True, h_only=False : 既存シートを全データ更新
      overwrite_existing=True, h_only=True  : 既存シートにH列(ReceptionMetrics)だけ追加
      overwrite_existing=False              : template_sheetをコピーして新月シート作成

    共通処理:
      - 外部リンク数式を全シートで静的値に置換(freeze)
      - 外部リンクparts自体も除去(keep_links=False)
      - 書き込み先シートのTTL行をPythonで再計算して静的値化
    """
    base = config['base_workbook']
    out = config['output_path']
    if os.path.abspath(base) == os.path.abspath(out):
        raise ValueError('保存先が元ファイルと同じです。別の名前・場所を指定してください。')

    # 元ファイルの「計算済みの値」(freezeの置換元)
    wb_values = openpyxl.load_workbook(base, data_only=True)

    shutil.copy(base, out)
    # keep_links=False: 外部リンクparts(他環境のファイルパス情報)を除去。
    # 数式文字列は残るが、直後のfreezeで全て値に置き換えるため問題ない。
    wb = openpyxl.load_workbook(out, keep_links=False)

    frozen_count = freeze_external_formulas(wb, wb_values)
    wb_values.close()

    h_only = bool(config.get('h_only'))
    overwrite = bool(config.get('overwrite_existing'))

    if overwrite:
        if config['write_sheet'] not in wb.sheetnames:
            raise ValueError(f"シート {config['write_sheet']} が見つかりません。")
        ws_new = wb[config['write_sheet']]
    else:
        template = config['template_sheet']
        if template not in wb.sheetnames:
            raise ValueError(f"コピー元シート {template} が見つかりません。")
        if config['write_sheet'] in wb.sheetnames:
            del wb[config['write_sheet']]
        ws_new = wb.copy_worksheet(wb[template])
        ws_new.title = config['write_sheet']

    # ディーラーマスタが指定されていれば、データ反映の前に閉店・改称を
    # 反映しておく(以降のF〜N列反映はこの後の店舗リストを対象に行う)
    dealer_master_report = None
    if config.get('dealer_master_path') and not h_only:
        master = load_dealer_master(config['dealer_master_path'])
        dealer_master_report = reconcile_dealer_master(ws_new, master)

    matched = {'F': 0, 'G': 0, 'H': 0, 'I_J': 0, 'L_N': 0, 'K_M_目標値': 0}
    dealer_codes_in_sheet = set()

    if h_only:
        dr = load_dialog_reception(config.get('ro_dr_path'), config.get('ro_dr_sheet'),
                                   expect_yyyymm=config.get('write_sheet'))
        if not dr:
            raise ValueError('ReceptionMetricsファイルからデータを読み取れませんでした。'
                             'ファイルとシート名([YYYYMM]DR)を確認してください。')
        for r in range(2, ws_new.max_row + 1):
            code = ws_new[f'B{r}'].value
            if code is None or ws_new[f'C{r}'].value == 'TTL':
                continue
            code = int(code)
            dealer_codes_in_sheet.add(code)
            if code in dr:
                ws_new[f'H{r}'] = dr[code]
                matched['H'] += 1
        unmatched_sources = {'ro_dr_未使用コード': sorted(set(dr) - dealer_codes_in_sheet)}
    else:
        ppso = load_ppso(config['ppso_paths'])
        mt = load_maintenance_table(config['elsapro_path'], config['elsa2go_path'])
        # MediaCaptureは、データ提供者から届くServiceMetrics形式ファイル(新方式)を
        # 優先する。無ければ「利用率集計」形式(旧方式)を試す。
        cam_so = load_service_cam_from_so_file(
            config.get('haga_file_path'), config.get('haga_file_sheet'))
        cam_legacy = load_service_cam(
            config.get('service_cam_path'), config.get('service_cam_sheet'))
        msqp = load_msqp(config['msqp_path'], config['write_month_num'])
        dr = load_dialog_reception(config.get('ro_dr_path'), config.get('ro_dr_sheet'),
                                   expect_yyyymm=config.get('write_sheet'))
        targets = load_targets(config.get('targets_path'), config.get('targets_sheet'))

        for r in range(2, ws_new.max_row + 1):
            code = ws_new[f'B{r}'].value
            if code is None or ws_new[f'C{r}'].value == 'TTL':
                continue  # TTL行は最後にまとめて再計算する
            code = int(code)
            dealer_codes_in_sheet.add(code)
            ws_new[f'E{r}'] = config['write_month_num']

            # F・G・L・Nは必須ソースから毎回網羅取得するため、常にクリアして埋め直す
            for col in ('F', 'G', 'L', 'N'):
                ws_new[f'{col}{r}'] = None
            # H・I・J(ReceptionMetrics/MediaCapture)は、既存シート更新時は今回データが
            # 無ければ既存値(前回投入分等)を保持。新規シート作成時のみ
            # 前月由来の値をクリアする。
            if not overwrite:
                for col in ('H', 'I', 'J'):
                    ws_new[f'{col}{r}'] = None
            # K・Mは必ずクリアしてから静的値で引き継ぐ(数式残存の余地を残さない)
            ws_new[f'K{r}'] = None
            ws_new[f'M{r}'] = None

            if code in ppso:
                ws_new[f'F{r}'] = ppso[code]
                matched['F'] += 1
            if code in mt:
                ws_new[f'G{r}'] = mt[code]
                matched['G'] += 1
            if code in dr:
                ws_new[f'H{r}'] = dr[code]
                matched['H'] += 1
            if code in cam_so:
                ws_new[f'I{r}'], ws_new[f'J{r}'] = cam_so[code]
                matched['I_J'] += 1
            elif code in cam_legacy:
                ws_new[f'I{r}'] = cam_legacy[code]
                ws_new[f'J{r}'] = cam_legacy[code]  # 暫定:I列と同値(仕様か否か前任担当者に確認中)
                matched['I_J'] += 1
            if code in msqp:
                mrow = msqp[code]
                if isinstance(mrow['q_actual'], (int, float)):
                    ws_new[f'L{r}'] = mrow['q_actual']
                if isinstance(mrow['final_actual'], (int, float)):
                    ws_new[f'N{r}'] = mrow['final_actual']
                matched['L_N'] += 1
            if code in targets:
                k_val, m_val = targets[code]
                if k_val not in (None, ''):
                    ws_new[f'K{r}'] = k_val
                if m_val not in (None, ''):
                    ws_new[f'M{r}'] = m_val
                matched['K_M_目標値'] += 1

        unmatched_sources = {
            'ppso_未使用コード': sorted(set(ppso) - dealer_codes_in_sheet),
            'mt_未使用コード': sorted(set(mt) - dealer_codes_in_sheet),
        }

    # TTL行の静的値化は書き込みシートだけでなく全YYYYMMシートに適用する。
    # 過去月のTTLは内部SUM数式のままだと、openpyxl保存でキャッシュ値が失われ、
    # Excel以外(pandas等)から読むと空欄に見えるため、値で確定させる。
    ttl = None
    for sn in wb.sheetnames:
        if re.fullmatch(r'\d{6}', sn):
            t = recompute_ttl(wb[sn])
            if sn == config['write_sheet']:
                ttl = t
    wb.save(out)
    wb.close()

    return {
        'モード': 'H列のみ追加' if h_only else ('既存シート更新' if overwrite else '新規シート作成'),
        '書き込みシート': config['write_sheet'],
        '反映店舗数': matched,
        '対象店舗数': len(dealer_codes_in_sheet),
        'TTL行(再計算後)': ttl,
        '静的化した外部リンク数式セル数': frozen_count,
        'ソース側の未使用コード(新規店舗の可能性・要確認)': unmatched_sources,
        'ディーラーマスタ照合結果': dealer_master_report,
        '★既知の注意事項(毎回表示)': [
            'G列は全ブランド機械合算(V+N)方式=正式運用(docstring【運用方針】参照)。'
            '前任担当者旧方式(手作業VLOOKUP)との差は手作業ミス由来と確認済み'
            '(2026/6実績: 91/276店舗差, 合計+67=0.28%)',
            '86010(徳島)は2026年6月末閉店・86300が7/1付で「VW徳島」に改称。'
            '7月分作成時に86010の行削除と86300の店名更新を行うこと'
            '(86010はマスタにコードが残存しているため自動検知されない)',
        ],
    }


# ============================================================
# 4. 月整合チェック(誤った月への書き込み防止ガード)
# ============================================================
# 【経緯】書き込み先シートの月とソースファイルの月が食い違ったまま実行し、
# 5月シートに6月データが混入する事故が実際に発生した(2026/7)。
# 実行前に各ソースの対象月を推定し、書き込み先と照合して警告する。

def _yyyymm_from_filename(path):
    """ファイル名から対象月(YYYYMM)を推定する。判定不能ならNone。"""
    name = os.path.basename(path or '')
    m = re.search(r'(20\d{2})\s*年\s*(\d{1,2})\s*月', name)      # 2026年6月
    if m:
        return f'{int(m.group(1))}{int(m.group(2)):02d}'
    m = re.search(r'(\d{2})[._](20\d{2})', name)                 # 06.2026 / 06_2026
    if m:
        return f'{m.group(2)}{m.group(1)}'
    m = re.search(r'(20\d{2})(0[1-9]|1[0-2])(?!\d)', name)       # 202606
    if m:
        return m.group(1) + m.group(2)
    return None


def _month_of_elsapro(path):
    """LegacyServiceToolファイルの対象月。ヘッダー部の'Time period:'(例 2026-06)を最優先、
    無ければファイル名から推定。"""
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb['Bericht 1']
        for r in range(1, 10):
            if str(ws[f'D{r}'].value).strip() == 'Time period:':
                m = re.match(r'(20\d{2})-(\d{2})', str(ws[f'F{r}'].value))
                wb.close()
                if m:
                    return m.group(1) + m.group(2)
        wb.close()
    except Exception:
        pass
    return _yyyymm_from_filename(path)


def _month_of_elsa2go(path):
    """NewServiceTool CSVの対象月。ヘッダー先頭列(例 '6-2026')から判定。"""
    try:
        with open(path, newline='', encoding='utf-8-sig') as f:
            header = next(csv.reader(f))
        m = re.match(r'(\d{1,2})-(20\d{2})', str(header[0]).strip())
        if m:
            return f'{m.group(2)}{int(m.group(1)):02d}'
    except Exception:
        pass
    return _yyyymm_from_filename(path)


def validate_month_consistency(config):
    """各ソースの対象月を書き込み先シートと照合する。
    戻り値: (確認用の行リスト, 不整合の有無)"""
    target = config['write_sheet']
    lines, mismatch = [], False

    def add(label, detected):
        nonlocal mismatch
        if detected is None:
            lines.append(f'・{label}: 月を判定できず(要目視確認)')
        elif detected == target:
            lines.append(f'・{label}: {detected} 一致')
        else:
            lines.append(f'★{label}: {detected} ≠ 書き込み先{target} 【月違い!】')
            mismatch = True

    if config.get('h_only'):
        res = _check_dialog_reception(config)
        if res:
            lines.append(res[0])
            mismatch = mismatch or res[1]
        return lines, mismatch

    labels = ['UsagePortal(VW)', 'UsagePortal(商用車)']
    for i, p in enumerate(config.get('ppso_paths') or []):
        add(labels[i] if i < 2 else f'UsagePortal{i+1}', _yyyymm_from_filename(p))
    add('LegacyServiceTool', _month_of_elsapro(config['elsapro_path']))
    add('NewServiceTool', _month_of_elsa2go(config['elsa2go_path']))

    # QualityMetricsは年間累積ファイルのため、当月実績列にデータがあるかで判定する
    msqp = load_msqp(config['msqp_path'], config['write_month_num'])
    n_actual = sum(1 for v in msqp.values()
                   if isinstance(v['q_actual'], (int, float))
                   or isinstance(v['final_actual'], (int, float)))
    if n_actual > 0:
        lines.append(f"・QualityMetrics実績: {config['write_month_num']}月の実績あり({n_actual}店舗) 一致")
    else:
        lines.append(f"★QualityMetrics実績: {config['write_month_num']}月の実績が0件(古いファイルの可能性)")
        mismatch = True

    res = _check_dialog_reception(config)
    if res:
        lines.append(res[0])
        mismatch = mismatch or res[1]

    return lines, mismatch


# ============================================================
# 5. ダイアログでファイルを選ばせてconfigを組み立てる
# ============================================================

_XLSX_TYPES = [('Excelファイル', '*.xlsx'), ('すべてのファイル', '*.*')]


# ------------------------------------------------------------
# 「前回使った正しいファイル」の記憶(248店舗版への後戻り防止)
# ------------------------------------------------------------
# 【経緯】276店舗版(前任担当者最終版)と248店舗版(古い作業用ファイル)を
# 取り違える事故が3回続いた。都度①でファイルを選び直す限り、いつでも
# 古い方を誤って選び得る。そこで、店舗数チェックに合格した(=270店舗以上の)
# 出力ファイルのパスを記憶しておき、次回起動時に「前回のファイルを続けて
# 使うか」を最初に聞く。「はい」を選び続ける限り、276店舗の系譜が
# そのまま毎月引き継がれ、古いファイルを選ぶ機会自体が減る。

def _state_file_path():
    """状態ファイル(前回の正しいファイルの記憶)の保存先。
    スクリプトと同じフォルダに直接置くと、①BATファイル・②Excel等と
    混ざって見た目が煩雑になるため、Windowsの正規のアプリ設定用フォルダ
    (%LOCALAPPDATA%)配下の専用サブフォルダに保存する。
    (%LOCALAPPDATA%が無い環境ではスクリプトのフォルダにフォールバックする)
    """
    base = os.getenv('LOCALAPPDATA') or os.path.dirname(os.path.abspath(__file__))
    folder = os.path.join(base, 'ServiceMetricsTool')
    try:
        os.makedirs(folder, exist_ok=True)
    except Exception:
        folder = os.path.dirname(os.path.abspath(__file__))  # 作成に失敗したら従来の場所へ
    new_path = os.path.join(folder, 'state.json')

    # 旧バージョンがスクリプト直下に残した状態ファイルがあれば、
    # 一度だけ新しい場所へ移し、フォルダの汚れを掃除する。
    old_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.so_tool_state.json')
    if os.path.exists(old_path) and not os.path.exists(new_path):
        try:
            shutil.move(old_path, new_path)
        except Exception:
            pass
    elif os.path.exists(old_path) and os.path.exists(new_path):
        try:
            os.remove(old_path)  # 新しい場所に既にあるなら、古い方は削除するだけ
        except Exception:
            pass
    return new_path


def _load_last_good_file():
    try:
        with open(_state_file_path(), encoding='utf-8') as f:
            data = json.load(f)
        path = data.get('last_good_file')
        if path and os.path.exists(path):
            return path
    except Exception:
        pass
    return None


def _save_last_good_file(path):
    try:
        with open(_state_file_path(), 'w', encoding='utf-8') as f:
            json.dump({'last_good_file': os.path.abspath(path)}, f, ensure_ascii=False)
    except Exception:
        pass  # 記憶の保存に失敗しても本体処理は止めない
_CSV_TYPES = [('CSVファイル', '*.csv'), ('すべてのファイル', '*.*')]
_MONTH_EN = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
             'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']


def _month_before(yyyymm):
    y, m = int(yyyymm[:4]), int(yyyymm[4:6])
    m -= 1
    if m == 0:
        m, y = 12, y - 1
    return f'{y}{m:02d}'


def _month_after(yyyymm):
    y, m = int(yyyymm[:4]), int(yyyymm[4:6])
    m += 1
    if m == 13:
        m, y = 1, y + 1
    return f'{y}{m:02d}'


def _pick_file(title, filetypes):
    path = filedialog.askopenfilename(title=title, filetypes=filetypes)
    return path or None


_PERMISSION_HINT = (
    'ファイルへのアクセスが拒否されました(PermissionError)。\n'
    'よくある原因と対処:\n'
    '1) そのファイルをExcelで開いたまま\n'
    '   → Excelで閉じてから「はい」で再試行\n'
    '2) OneDrive上の「クラウドのみ」ファイル(雲マーク)\n'
    '   → エクスプローラーで右クリック→\n'
    '     「このデバイス上に常に保持する」を選んでから再試行\n'
    '3) 上記で解決しない場合\n'
    '   → OneDrive外のローカルフォルダ(例: C:\\Temp)に\n'
    '     コピーし、そちらを選択')


def _try_open_check(path):
    """選択されたファイルが実際に読み取れるか事前確認する。
    OneDriveのクラウドのみファイル・Excelでのロック等による
    PermissionErrorを、選択直後に分かりやすく検出するため。"""
    try:
        with open(path, 'rb') as f:
            f.read(4)
        return True, None
    except PermissionError:
        return False, _PERMISSION_HINT
    except OSError as e:
        return False, f'ファイルを開けません:\n{e}'


def _pick_file_checked(title, filetypes):
    """ファイル選択+読み取り可能チェック。読めない場合は原因と対処を
    表示して再試行できる。キャンセル時はNoneを返す。"""
    while True:
        path = _pick_file(title, filetypes)
        if not path:
            return None
        ok, err = _try_open_check(path)
        if ok:
            return path
        if not messagebox.askyesno(
                '読み取りエラー',
                f'{err}\n\n対処後、同じファイルを選び直しますか?\n'
                '(「いいえ」でこのファイルの選択をキャンセル)'):
            return None


def _month_sheets_of(path):
    wb = openpyxl.load_workbook(path, read_only=True)
    sheets = sorted(s for s in wb.sheetnames if re.fullmatch(r'\d{6}', s))
    wb.close()
    return sheets


def pick_files_interactively():
    """ダイアログを順に出してbuild_month_sheet用のconfigを組み立てる。
    必須ファイルがキャンセルされたらエラー表示してNoneを返す。"""
    root = tk.Tk()
    root.withdraw()

    # ① ベースとなるServiceMetricsワークブック
    # 前回、店舗数チェックに合格したファイルがあれば、まずそれを
    # 続けて使うか聞く(古い248店舗版へ後戻りする事故を防ぐため)。
    base_path = None
    last_good = _load_last_good_file()
    if last_good:
        if messagebox.askyesno(
                '①前回のファイルを使用しますか?',
                f'前回、店舗数チェックに合格したファイルを記憶しています。\n\n'
                f'{last_good}\n\n'
                '「はい」→ これを続けて使う(推奨。取り違え防止)\n'
                '「いいえ」→ 別のファイルを選び直す'):
            base_path = last_good
    if base_path is None:
        base_path = _pick_file_checked(
            '①ServiceMetricsワークブックを選択'
            '(前回の出力・前任担当者の系譜ファイル等。データ提供者から届いた'
            'ファイルは選ばないこと。MediaCaptureは⑥で別途取り込みます)',
            _XLSX_TYPES)
    if not base_path:
        messagebox.showerror('中断', 'ServiceMetricsワークブックが未選択のため中断します。')
        return None

    month_sheets = _month_sheets_of(base_path)
    if not month_sheets:
        messagebox.showerror('エラー', '「YYYYMM」形式のシートが見つかりません。ファイルを確認してください。')
        return None
    latest_sheet = month_sheets[-1]

    # ①選択直後に店舗数を確認する。ここで気づければ、残り7つの
    # ファイル選択を無駄にすることなく、早い段階でやり直せる。
    # (誰から届いたファイルか、ではなく中身の店舗数だけで判定する。
    #  データ提供者はMediaCapture分を追記するだけなので、彼が最後に触った
    #  ファイル=全店舗分とは限らない。中身を都度確認するのが確実)
    try:
        _wb0 = openpyxl.load_workbook(base_path, read_only=True)
        _ws0 = _wb0[latest_sheet]
        _cnt0 = sum(1 for r in range(2, _ws0.max_row + 1)
                    if _ws0[f'B{r}'].value is not None and _ws0[f'C{r}'].value != 'TTL')
        _wb0.close()
    except Exception:
        _cnt0 = None
    if _cnt0 is not None and _cnt0 < 270:
        if not messagebox.askyesno(
                '①の店舗数が少なめです',
                f'選んだファイルの{latest_sheet}シートは{_cnt0}店舗分しかありません\n'
                '(全店舗版は276店舗程度のはずです)。\n\n'
                '一部の店舗しか無いファイル(例: MediaCapture更新用に抜粋した'
                'もの等)を選んでいないか確認してください。\n\n'
                'このまま続行しますか?(残り7つの選択が無駄になるより、\n'
                '今ここでファイルを選び直すことを推奨します)'):
            messagebox.showinfo('中断', 'ファイルの選び直しのため中断しました。')
            return None

    # ②モード選択(既存月の更新か、新規月の作成か)
    overwrite_existing = messagebox.askyesno(
        '処理内容の確認',
        f'このファイルの最新シートは {latest_sheet} です。\n\n'
        f'「はい」→ {latest_sheet} を更新します(未完成分の入力/H列の追加)\n'
        f'「いいえ」→ 次の月({_month_after(latest_sheet)})のシートを新規作成します')

    h_only = False
    template_sheet = None
    if overwrite_existing:
        write_sheet = latest_sheet
        targets_sheet_default = _month_before(latest_sheet)
        # ※以前はここで「H列だけ追加」モードを選ばせていたが、ReceptionMetricsは
        # 今後毎回入手できる前提のため、常に全データ更新にして質問を廃止した。
        # ReceptionMetricsは後段⑧で他の任意ファイルと同じ扱いで選択する。
    else:
        write_sheet = _month_after(latest_sheet)
        template_sheet = latest_sheet
        targets_sheet_default = latest_sheet
        if not messagebox.askyesno(
                '新規シートの確認',
                f'新規作成するシート: {write_sheet}\n'
                f'コピー元(テンプレート): {template_sheet}\n\nこの内容でよろしいですか?'):
            typed = simpledialog.askstring(
                '新規シート名を入力', '作成する月を6桁(例: 202608)で入力してください:')
            if not typed or not re.fullmatch(r'\d{6}', typed):
                messagebox.showerror('中断', '入力が不正なため中断します。')
                return None
            if typed in month_sheets:
                if not messagebox.askyesno(
                        '上書き確認',
                        f'シート {typed} は既に存在します。作り直しますか?\n'
                        '(既存の入力内容は失われます)'):
                    messagebox.showerror('中断', '処理を中断しました。')
                    return None
            write_sheet = typed
            targets_sheet_default = _month_before(typed)

    write_month_num = int(write_sheet[4:6])

    config = {
        'base_workbook': base_path,
        'write_sheet': write_sheet,
        'write_month_num': write_month_num,
        'overwrite_existing': overwrite_existing,
        'h_only': h_only,
        'template_sheet': template_sheet,
        'targets_path': base_path,
        'targets_sheet': targets_sheet_default,
        'ppso_paths': [],
        'elsapro_path': None,
        'elsa2go_path': None,
        'service_cam_path': None,
        'service_cam_sheet': None,
        'haga_file_path': None,
        'haga_file_sheet': None,
        'msqp_path': None,
        'ro_dr_path': None,
        'ro_dr_sheet': None,
    }

    # ② UsagePortal(VW) ③ UsagePortal(商用車)
    ppso_vw = _pick_file_checked('②UsagePortal Usage Report(VW用)を選択', _XLSX_TYPES)
    if not ppso_vw:
        messagebox.showerror('中断', 'UsagePortal(VW)が未選択のため中断します。')
        return None
    ppso_cv = _pick_file_checked('③UsagePortal Usage Report(商用車用)を選択', _XLSX_TYPES)
    if not ppso_cv:
        messagebox.showerror('中断', 'UsagePortal(商用車)が未選択のため中断します。')
        return None
    config['ppso_paths'] = [ppso_vw, ppso_cv]

    # ④ LegacyServiceTool ⑤ NewServiceTool
    elsapro = _pick_file_checked('④LegacyServiceTool Maintenance Tablesを選択', _XLSX_TYPES)
    if not elsapro:
        messagebox.showerror('中断', 'LegacyServiceToolファイルが未選択のため中断します。')
        return None
    elsa2go = _pick_file_checked('⑤NewServiceToolReport(CSV)を選択', _CSV_TYPES)
    if not elsa2go:
        messagebox.showerror('中断', 'NewServiceToolReportが未選択のため中断します。')
        return None
    config['elsapro_path'] = elsapro
    config['elsa2go_path'] = elsa2go

    # ⑥ MediaCapture: ①にはデータ提供者のファイルを使わない(店舗リストの基準は
    # 前任担当者/担当者の系譜ファイルに統一する)。そのため、MediaCaptureの
    # 値だけはデータ提供者から届く別ファイルから、店舗コードで都度取り込む。
    if messagebox.askyesno(
            'MediaCapture(データ提供者のファイル)',
            'MediaCaptureデータ(I・J列)を、データ提供者から届いた\n'
            'ファイルから取り込みますか?\n\n'
            '(①には使わず、ここで選ぶファイルから店舗コードで\n'
            '該当する値だけを取り込みます。①の店舗リストが基準のまま\n'
            '変わることはありません)'):
        haga_path = _pick_file_checked('⑥データ提供者から届いたファイルを選択', _XLSX_TYPES)
        if haga_path:
            config['haga_file_path'] = haga_path
            config['haga_file_sheet'] = write_sheet

    # ⑦ QualityMetrics実績
    msqp = _pick_file_checked('⑦QualityMetrics実績を選択', _XLSX_TYPES)
    if not msqp:
        messagebox.showerror('中断', 'QualityMetrics実績が未選択のため中断します。')
        return None
    config['msqp_path'] = msqp

    # ⑧ ReceptionMetrics(任意。月初時点でまだ入手できていない場合は「いいえ」でよい)
    if messagebox.askyesno(
            'ダイアログレセプション',
            'H列(ダイアログレセプション)の元データは入手済みですか?\n\n'
            '次のどちらでも構いません:\n'
            f'  ・ReceptionMetrics {write_sheet}.xlsx (従来)\n'
            f'  ・{write_sheet}_T_Jisseki_su_sum_2.xlsx (Access出力・直読)\n\n'
            'Access出力を選べば、ReceptionMetricsブックへの貼り付けは不要です。'):
        ro_dr_path = _pick_file_checked(
            f'⑧ReceptionMetrics {write_sheet}.xlsx または {write_sheet}_T_Jisseki_su_sum_2.xlsx を選択',
            _XLSX_TYPES)
        if ro_dr_path:
            config['ro_dr_path'] = ro_dr_path
            config['ro_dr_sheet'] = f'{write_sheet}DR'

    # ディーラーマスタ(任意)。閉店・改称をこの時点で検知して反映する
    config['dealer_master_path'] = None
    if messagebox.askyesno(
            'ディーラーマスタ(任意)',
            '最新のディーラーマスタ(店舗マスター・サテライトマスター)は\n'
            'ありますか?\n\n'
            '「はい」を選ぶと、マスタに無い店舗コード(閉店の疑い)を\n'
            'まとめて確認・削除できます。店名の差分・新規店舗候補は\n'
            '一覧としてレポートに残ります(自動反映はしません)。'):
        dm_path = _pick_file_checked('ディーラーマスタファイルを選択', _XLSX_TYPES)
        if dm_path:
            config['dealer_master_path'] = dm_path

    # 目標値(K・M)の引き継ぎ元チェック。前月に値が無ければ参照ファイルを選ばせる
    try:
        targets = load_targets(base_path, targets_sheet_default)
    except Exception:
        targets = {}
    if len(targets) < 50:
        if messagebox.askyesno(
                '目標値(K・M列)の確認',
                f'①のファイルの {targets_sheet_default} シートに目標値(K・M列)が'
                f'ほとんど見つかりません(検出: {len(targets)}店舗)。\n\n'
                '目標値が入った別ファイル(例: 前任担当者の完成版)から'
                '取り込みますか?\n(「いいえ」の場合、K・M列は空欄になります)'):
            ref_path = _pick_file_checked('目標値の参照ファイルを選択', _XLSX_TYPES)
            if ref_path:
                try:
                    ref_sheets = _month_sheets_of(ref_path)
                except Exception as e:
                    messagebox.showerror(
                        '参照ファイルエラー',
                        f'目標値の参照ファイルを読み込めませんでした:\n{e}\n\n'
                        'K・M列は空欄のまま処理を続けます。')
                    ref_sheets = []
                if write_sheet in ref_sheets:
                    ref_sheet = write_sheet
                elif targets_sheet_default in ref_sheets:
                    ref_sheet = targets_sheet_default
                elif ref_sheets:
                    ref_sheet = ref_sheets[-1]
                else:
                    ref_sheet = None
                if ref_sheet:
                    config['targets_path'] = ref_path
                    config['targets_sheet'] = ref_sheet

    # 実行前の最終確認(月整合チェック+店舗数チェック)。
    # 書き込み先シートとソースの月が食い違うと誤った月に数字が入るため、
    # 実行前に必ず照合結果を表示し、不整合があれば中断を推奨する。
    # 店舗数も表示する(過去に248店舗版と276店舗版を取り違える事故が
    # 複数回発生したため、実行前に気づけるようにする)。
    try:
        check_lines, has_mismatch = validate_month_consistency(config)
    except Exception as e:
        check_lines = [f'(月チェックを実行できませんでした: {e})']
        has_mismatch = False

    try:
        _wb_cnt = openpyxl.load_workbook(base_path, read_only=True)
        _sheet_for_count = write_sheet if write_sheet in _wb_cnt.sheetnames else targets_sheet_default
        _ws_cnt = _wb_cnt[_sheet_for_count]
        dealer_count = sum(1 for r in range(2, _ws_cnt.max_row + 1)
                            if _ws_cnt[f'B{r}'].value is not None
                            and _ws_cnt[f'C{r}'].value != 'TTL')
        _wb_cnt.close()
    except Exception:
        dealer_count = None

    mode_label = '既存シートの全データ更新' if overwrite_existing else '新規シート作成'
    summary = (f'【実行前の最終確認】\n\n'
               f'処理: {mode_label}\n'
               f'書き込み先シート: {write_sheet}\n'
               f'①ファイルの対象店舗数: {dealer_count}店舗\n\n'
               f'[ソースの対象月チェック]\n' + '\n'.join(check_lines))
    if dealer_count is not None and dealer_count < 270:
        summary += (f'\n\n★店舗数が{dealer_count}件と少なめです。276店舗版と'
                    '248店舗版を取り違えていないか確認してください。')
        has_mismatch = True
    if has_mismatch:
        summary += ('\n\n★警告があります。このまま実行すると想定と違う\n'
                    '結果になる可能性があります。中断して①のファイル・\n'
                    'モード選択・ソースファイルを確認することを推奨します。'
                    '\n\n本当に続行しますか?')
    else:
        summary += '\n\nこの内容で実行しますか?'
    if not messagebox.askyesno('実行前の最終確認', summary):
        messagebox.showinfo('中断', '処理を中断しました。ファイルは変更されていません。')
        return None

    # ⑨ 保存先
    output_path = filedialog.asksaveasfilename(
        title='⑨保存先を指定',
        defaultextension='.xlsx',
        initialfile=f'ServiceMetrics_{write_sheet}.xlsx',
        filetypes=_XLSX_TYPES)
    if not output_path:
        messagebox.showerror('中断', '保存先が未指定のため中断します。')
        return None
    if os.path.abspath(output_path) == os.path.abspath(base_path):
        messagebox.showerror(
            '中断',
            '保存先が①の元ファイルと同じです。\n'
            '元ファイル保護のため、別の名前・場所を指定してください。')
        return None
    config['output_path'] = output_path

    root.destroy()
    return config



# ============================================================
# ReceptionMetricsブック生成(Access出力 → ReceptionMetrics [YYYYMM].xlsx へ当月シートを追加)
# ------------------------------------------------------------
# ServiceMetricsのH列はAccess出力を直読できるようになったため、
# この機能は「ReceptionMetricsブック自体を参照している人が他にいる場合」に
# 従来どおりのブックを維持するためのもの。
# 貼り付け位置・見出し・列順は202601〜202607の実物と同一形式で書き出す。
# ============================================================

# 書き出し先の固定レイアウト(実物のReceptionMetrics 202601〜202607と一致)
_DR_LAYOUT = {
    'start_col': 2,          # B列から
    'header_row': 2,
    'headers': ['DEALER_CODE', '店舗名の先頭', 'K', '*', '合計', 'YYYYMM'],
    # 書き出す見出し → Access出力側の見出し候補(先に見つかったものを採用)
    'source': {
        'DEALER_CODE': ['DEALER_CODE'],
        '店舗名の先頭': ['店舗名の先頭'],
        'K': ['CHECKOUT_YMDのカウントK', 'K'],
        '*': ['CHECKOUT_YMDのカウント*', '*'],
        '合計': ['CHECKOUT_YMDのカウント計', '合計'],
        'YYYYMM': ['YYYYMM'],
    },
    'notes': {},
}

_RO_LAYOUT = {
    'start_col': 2,
    'header_row': 2,
    'headers': ['DEALER_CODE', '店舗名の先頭', 'LINE_TYPE', 'CHECKOUT_YMDのカウント',
                '数値のカウント', 'XYZのカウント', '数値以外のカウント', 'YYYYMM'],
    'source': {
        'DEALER_CODE': ['DEALER_CODE'],
        '店舗名の先頭': ['店舗名の先頭'],
        'LINE_TYPE': ['LINE_TYPE'],
        'CHECKOUT_YMDのカウント': ['CHECKOUT_YMDのカウント'],
        '数値のカウント': ['数値のカウント'],
        'XYZのカウント': ['XYZのカウント'],
        '数値以外のカウント': ['数値以外のカウント'],
        'YYYYMM': ['YYYYMM'],
    },
    # 1行目の手書き見出し(前任担当者が付けていたもの)。列は書き出し後の位置で指定
    'notes': {'CHECKOUT_YMDのカウント': 'リペアオペレーション（分母）',
              '数値のカウント': 'リペアオペレーション（分子）'},
}


def _read_access_export(path):
    """Access出力(1行目=見出し)を読み、(見出しリスト, 行データ)を返す。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        headers = []
        for c in range(1, ws.max_column + 1):
            v = ws.cell(1, c).value
            headers.append(str(v).strip() if v is not None else None)
        rows = []
        for r in range(2, ws.max_row + 1):
            vals = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
            if all(v is None for v in vals):
                continue
            rows.append(dict(zip(headers, vals)))
        return headers, rows
    finally:
        wb.close()


def write_ro_dr_sheet(wb, sheet_name, layout, src_path, yyyymm):
    """Access出力1本から、ReceptionMetricsブックへシート1枚を書き出す。
    見出し名で対応付けるため、Access側の列順が入れ替わっても正しく並ぶ。
    """
    headers, rows = _read_access_export(src_path)
    if not rows:
        raise ValueError(f'{os.path.basename(src_path)} にデータ行がありません。')

    # 対象月の照合(単月ファイルなので月の取り違えが最大の事故要因)
    months = {str(r.get('YYYYMM')).strip() for r in rows if r.get('YYYYMM') is not None}
    if months and str(yyyymm) not in months:
        raise ValueError(
            f'{os.path.basename(src_path)} のYYYYMMは {"/".join(sorted(months))} です。\n'
            f'作成しようとしている月 {yyyymm} と一致しません。')

    # 必要な列がAccess出力に揃っているかを先に確認(揃わなければ静かに空にせず止める)
    colmap = {}
    for out_name in layout['headers']:
        for cand in layout['source'][out_name]:
            if cand in headers:
                colmap[out_name] = cand
                break
        else:
            raise ValueError(
                f'{os.path.basename(src_path)} に「{out_name}」に対応する列が'
                f'見つかりません。\n実際の見出し: {[h for h in headers if h]}')

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    sc = layout['start_col']
    hr = layout['header_row']
    for i, name in enumerate(layout['headers']):
        ws.cell(hr, sc + i, name)
        note = layout['notes'].get(name)
        if note:
            ws.cell(hr - 1, sc + i, note)
    for j, row in enumerate(rows):
        for i, name in enumerate(layout['headers']):
            ws.cell(hr + 1 + j, sc + i, row.get(colmap[name]))

    # 直前月の同種シートがあれば列幅を引き継ぐ(見た目を揃えるだけ)
    prev = _find_sheet(wb, f'{_month_before(yyyymm)}{sheet_name[6:]}')
    if prev:
        for key, dim in wb[prev].column_dimensions.items():
            if dim.width:
                ws.column_dimensions[key].width = dim.width
    return len(rows)


def run_ro_dr_builder():
    """ReceptionMetricsブックに当月のDR/ROシートを追加するモード。"""
    book = _pick_file_checked(
        '①ベースにするReceptionMetricsブック(ReceptionMetrics [前月].xlsx)を選択', _XLSX_TYPES)
    if not book:
        return
    wb = openpyxl.load_workbook(book)
    months = sorted({n.strip()[:6] for n in wb.sheetnames
                     if n.strip()[:6].isdigit()})
    default = _month_after(months[-1]) if months else ''
    yyyymm = simpledialog.askstring(
        '作成する月',
        f'ブック内の最新月: {months[-1] if months else "なし"}\n\n'
        '追加する月をYYYYMM形式で入力してください。',
        initialvalue=default)
    if not yyyymm or not (yyyymm.strip().isdigit() and len(yyyymm.strip()) == 6):
        messagebox.showerror('中断', 'YYYYMM(6桁)を入力してください。')
        return
    yyyymm = yyyymm.strip()
    if any(n.strip() == f'{yyyymm}DR' for n in wb.sheetnames) and not messagebox.askyesno(
            '確認', f'{yyyymm}DR は既に存在します。作り直しますか?'):
        return

    dr_src = _pick_file_checked(
        f'②DRシートの元データ {yyyymm}_T_Jisseki_su_sum_2.xlsx を選択', _XLSX_TYPES)
    if not dr_src:
        messagebox.showerror('中断', 'DRの元データが未選択のため中断します。')
        return
    ro_src = None
    if messagebox.askyesno(
            'ROシート',
            'ROシートも作成しますか?\n\n'
            '「はい」の場合、ROシートの元データ\n'
            f'({yyyymm}_T_Jisseki_su_sum_K_all.xlsx 相当)を次に選択します。\n\n'
            '※ServiceMetricsのH列はDRシートしか使いません。\n'
            '　ROシートを他で参照している人がいなければ「いいえ」で構いません。'):
        ro_src = _pick_file_checked('③ROシートの元データを選択', _XLSX_TYPES)

    try:
        n_dr = write_ro_dr_sheet(wb, f'{yyyymm}DR', _DR_LAYOUT, dr_src, yyyymm)
        n_ro = write_ro_dr_sheet(wb, f'{yyyymm}RO', _RO_LAYOUT, ro_src, yyyymm) if ro_src else 0
    except Exception as e:
        messagebox.showerror('エラー', str(e))
        return

    out = filedialog.asksaveasfilename(
        title='保存先を指定', defaultextension='.xlsx',
        initialfile=f'ReceptionMetrics {yyyymm}.xlsx', filetypes=_XLSX_TYPES)
    if not out:
        return
    try:
        wb.save(out)
    except PermissionError as e:
        messagebox.showerror('エラー(アクセス拒否)', f'{_PERMISSION_HINT}\n\n詳細: {e}')
        return
    finally:
        wb.close()

    msg = (f'保存先: {out}\n\n'
           f'{yyyymm}DR: {n_dr}店舗\n'
           + (f'{yyyymm}RO: {n_ro}店舗\n' if ro_src else 'ROシート: 作成せず\n')
           + '\n※シート名に余分な空白は入っていません。')
    messagebox.showinfo('完了', msg)
    print(msg)


def main():
    # 起動時にモードを選ぶ。通常はServiceMetrics作成。
    if messagebox.askyesno(
            'モード選択',
            'ReceptionMetricsブックに当月シート(DR/RO)を追加しますか?\n\n'
            '「はい」 → ReceptionMetricsブック作成モード\n'
            '「いいえ」→ ServiceMetrics作成(通常)\n\n'
            '※ServiceMetricsのH列はAccess出力を直読できるため、\n'
            '　ReceptionMetricsブックは他の参照者がいる場合のみ維持すれば足ります。'):
        run_ro_dr_builder()
        return

    config = pick_files_interactively()
    if config is None:
        return
    try:
        report = build_month_sheet(config)
    except PermissionError as e:
        messagebox.showerror(
            'エラー(アクセス拒否)',
            f'{_PERMISSION_HINT}\n\n'
            '※保存先に既存ファイルを指定した場合、そのファイルを\n'
            '　Excelで開いたままだと保存できません。\n\n詳細: {0}'.format(e))
        return
    except Exception as e:
        messagebox.showerror('エラー', f'処理中にエラーが発生しました:\n{e}')
        raise

    lines = [f'保存先: {config["output_path"]}', '', '=== 反映結果 ===']
    for k, v in report.items():
        lines.append(f'{k}: {v}')
    result_text = '\n'.join(lines)

    # 店舗数が十分(270以上)なら、次回①のデフォルト候補として記憶する。
    # 248店舗版のような店舗数が少ないファイルは記憶を更新しない
    # (取り違えた古いファイルが「正」として定着してしまうのを防ぐため)。
    dealer_cnt = report.get('対象店舗数')
    if isinstance(dealer_cnt, int) and dealer_cnt >= 270:
        _save_last_good_file(config['output_path'])
    print(result_text)
    messagebox.showinfo('完了', result_text[:1500])


if __name__ == '__main__':
    main()
