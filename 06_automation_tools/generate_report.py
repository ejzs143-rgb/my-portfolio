"""
==============================================================
業務報告レポート 生成ツール
==============================================================
使い方:
  python generate_report.py                          # 対話モード（GUIダイアログ）
  python generate_report.py 業務時間_Analyst.xlsx    # ファイル指定
  python generate_report.py 業務時間.xlsx 2026-02-16 # 期間も指定

出力: 業務報告_承認担当者向け_MMDD.xlsx（スクリプトと同じフォルダ）
==============================================================
"""

import sys
import io

# Windows CMD での文字化け防止
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
        
import os
import re
import warnings
import tkinter as tk
from tkinter import ttk, filedialog
from datetime import datetime, timedelta
from collections import defaultdict

import numpy as np
import pandas as pd
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════
# 設定：個人業務リスト（担当者担当固有の業務）
# ══════════════════════════════════════════════════════════
PERSONAL_TASKS = {
    "BrandA ProgramB改善",
    "BRANDBクレームチェック",
    "ProgramA管理",
    "予算実績管理",
    "重点店舗改善",
    "オントレ未受講者撲滅",
    "サプライヤー契約更改",
    "MO+RE情報共有",
    "InternalBusinessAppユーザー登録",
    "販売店業務改善",
    "ワランティマイスターコミッティー",
    "ローカルグッドウィル予算計算",
    "TOEIC学習",
    "Sample Vendor 契約/RFA/SOW",
    "オンライントレーニング刷新＆受講促進活動",
    "ProgramA MO+RE",
    "コアプロ関係",
    "RFP RFA DASS PO 契約関係",
    "AI関連",
    "コアプロ集計",
    "コアプロ集計・分析",
    "オントレ刷新",
}

# ══════════════════════════════════════════════════════════
# 設定：タスク命名辞書（備考のキーワード→タスク名）
# ══════════════════════════════════════════════════════════
LABEL_DICT = {
    "ServiceDesk":       "ServiceDesk問合せ対応",
    "Sample Vendor":      "Sample Vendor 契約/RFA/SOW",
    "ObjectiveReview":       "ObjectiveReview対応",
    "mbo":       "ObjectiveReview対応",
    "ObjectiveReview":       "ObjectiveReview対応",
    "判ミ":      "判定MTG",
    "課内会議":  "課内会議",
    "FBT":       "FBT定例",
    "KD2":       "KD2 MTG",
    "エラマネ":  "エラー分析MTG",
    "Sample Project":    "Sample Project案件",
    "VendorC":      "VendorC支払管理",
    "オントレ":  "オンライントレーニング刷新＆受講促進活動",
    "ATP":       "オンライントレーニング刷新＆受講促進活動",
    "ProgramA":       "ProgramA管理",
    "ProgramB":       "BrandA ProgramB改善",
    "BRANDB":      "BRANDBクレームチェック",
    "ExampleBrand B":      "BRANDBクレームチェック",
    "予算":      "予算実績管理",
    "TOEIC":     "TOEIC学習",
    "Vendor A":  "Vendor A/Vendor B対応",
    "Vendor B":    "Vendor A/Vendor B対応",
    "コアプロ":  "コアプロ関係",
    "Sample Project":      "コアプロ関係",
    "InternalBusinessApp":      "InternalBusinessAppユーザー登録",
    "InternalNotice":      "InternalNotice書簡",
    "拒絶":      "拒絶対応",
    "DASS":      "DASS対応",
    "StM":       "StM MTG",
    "Stm":       "StM MTG",
    "アカデミー":"アカデミー",
    "業務整理":  "業務整理",
    "AI":        "AI関連",
    "水没":      "水没車対応",
    "書簡":      "InternalNotice書簡",
}
SKIP_WORDS = {
    "さん","対応","確認","含む","関係","関連","業務","整理","準備",
    "nan","こと","する","いる","ある","です","ます","など","等",
    "検討","処理","実施","依頼","報告","作成","展開","修正",
}


# ══════════════════════════════════════════════════════════
# データ読み込み・クラスタリング
# ══════════════════════════════════════════════════════════

def load_and_cluster(excel_path):
    df = pd.read_excel(excel_path, sheet_name="入力", header=None)
    d  = df[df[1].notna() & (df[1] != "業務内容")].copy()
    d.columns = ["date","cat","start","end","dur","hours","memo"]
    d["date"]  = pd.to_datetime(d["date"].ffill())
    d          = d[d["date"].notna()].copy()
    d["hours"] = pd.to_numeric(d["hours"], errors="coerce").fillna(0)
    d          = d[d["hours"] > 0].copy().reset_index(drop=True)
    d["memo_clean"] = d["memo"].fillna("").astype(str).str.strip().replace("nan", "")
    return _cluster(d)


def _label_cluster(texts, cat):
    combined = " ".join(texts)
    for kw, lab in LABEL_DICT.items():
        if kw in combined:
            return lab
    tokens = re.findall(r'[^\s　、。・\[\]（）\d【】「」]+', combined)
    freq = defaultdict(int)
    for t in tokens:
        if len(t) >= 2 and t not in SKIP_WORDS:
            freq[t] += 1
    top = sorted(freq.items(), key=lambda x: -x[1])[:2]
    return f"{cat}（{'・'.join(t[0][:8] for t in top)}）" if top else cat


def _cluster(df):
    df = df.copy(); df["task"] = ""
    for cat, grp in df.groupby("cat"):
        idxs  = grp.index.tolist()
        memos = grp["memo_clean"].tolist()
        has   = [bool(m) for m in memos]

        for i, h in zip(idxs, has):
            if not h: df.at[i, "task"] = cat

        midx   = [i for i, h in zip(idxs, has) if h]
        mtexts = [memos[grp.index.tolist().index(i)] for i in midx]
        if not midx:
            continue

        if len(midx) < 3:
            for i, txt in zip(midx, mtexts):
                matched = False
                for kw, lab in LABEL_DICT.items():
                    if kw in txt:
                        df.at[i, "task"] = lab
                        matched = True
                        break
                if not matched:
                    df.at[i, "task"] = f"{cat}（{txt[:12].strip()}）" if txt else cat
            continue

        try:
            from sklearn.feature_extraction.text import TfidfVectorizer
            from sklearn.cluster import KMeans
            from sklearn.preprocessing import normalize
            
            X = normalize(TfidfVectorizer(
                analyzer="char", ngram_range=(2, 4),
                max_features=300, min_df=1
            ).fit_transform(mtexts))
            n_cl = max(2, min(10, int(np.sqrt(len(midx)))))
            labels = KMeans(n_clusters=n_cl, random_state=42, n_init=10).fit_predict(X)
            ct = defaultdict(list); ci = defaultdict(list)
            for ei, (l, t) in enumerate(zip(labels, mtexts)):
                ct[l].append(t); ci[l].append(midx[ei])
            for l, ts in ct.items():
                lab = _label_cluster(ts, cat)
                for i in ci[l]:
                    df.at[i, "task"] = lab
        except ImportError:
            for i, txt in zip(midx, mtexts):
                matched = False
                for kw, lab in LABEL_DICT.items():
                    if kw in txt:
                        df.at[i, "task"] = lab
                        matched = True
                        break
                if not matched:
                    df.at[i, "task"] = cat
        except Exception:
            for i, txt in zip(midx, mtexts):
                matched = False
                for kw, lab in LABEL_DICT.items():
                    if kw in txt:
                        df.at[i, "task"] = lab
                        matched = True
                        break
                if not matched:
                    df.at[i, "task"] = cat
    return df


# ══════════════════════════════════════════════════════════
# タスク集計（備考の自動抽出機能を追加）
# ══════════════════════════════════════════════════════════

def aggregate(focus_df, all_df):
    # ★追加：タスクごとに備考欄のテキストを集め、頻出順に最大3件を抽出する内部関数
    def extract_memos(series):
        valid = [str(m).strip() for m in series if str(m).strip()]
        if not valid: return "-"
        # 頻出する備考を上位3つ抽出し、改行で繋ぐ
        return " \n ".join(pd.Series(valid).value_counts().head(3).index)

    g = (focus_df.groupby("task")
                 .agg(時間=("hours","sum"), 
                      日数=("date","nunique"),
                      主な作業=("memo_clean", extract_memos)) # ★ここで備考を統合
                 .sort_values("時間", ascending=False)
                 .reset_index())
    g["時間"] = g["時間"].round(1)
    total = focus_df["hours"].sum()

    all_task  = all_df.groupby("task")["hours"].sum()
    all_weeks = all_df["date"].dt.to_period("W").nunique()
    foc_weeks = max(focus_df["date"].dt.to_period("W").nunique(), 1)

    def is_new(task, h):
        all_h = float(all_task.get(task, 0))
        if all_h <= 0: return True
        return (h / foc_weeks) >= (all_h / max(all_weeks, 1)) * 1.5

    g["新規"] = g.apply(lambda r: is_new(r["task"], r["時間"]), axis=1)
    g["個人"] = g["task"].apply(_is_personal)
    return g, round(total, 1)


def _is_personal(task_name):
    if task_name in PERSONAL_TASKS: return True
    for p in PERSONAL_TASKS:
        if p in task_name or task_name in p: return True
    return False


# ══════════════════════════════════════════════════════════
# スタイル定数
# ══════════════════════════════════════════════════════════

FN = "Meiryo UI"

COL = {
    "title":      "1F3864",
    "hdr":        "2E75B6",
    "ans_bg":     "EBF3FB",
    "ans_fg":     "1F3864",
    "personal":   "C9473F",   
    "common":     "1F497D",   
    "p_row":      ["FCE4D6", "FDEBD0", "FEF5E7"],  
    "c_row":      ["D6E4F0", "EBF3FB", "F0F8FF"],  
    "p_badge_bg": "FCE4D6",
    "c_badge_bg": "D6E4F0",
    "even":       "FFFFFF",
    "odd":        "F5F5F5",
    "note":       "F5F5F5",
}

thin_s  = Side(style="thin",   color="CCCCCC")
thick_s = Side(style="medium", color="1F3864")


def BG(h):  return PatternFill("solid", start_color=h, fgColor=h)
def BD():   return Border(left=thin_s,  right=thin_s,  top=thin_s,  bottom=thin_s)
def BDB():  return Border(left=thick_s, right=thick_s, top=thick_s, bottom=thick_s)


def MR(ws, row, c1, c2, val, bold=False, sz=11,
       fg="FFFFFF", bg="1F3864", ha="left", h=30, wrap=False):
    ws.merge_cells(f"{get_column_letter(c1)}{row}:{get_column_letter(c2)}{row}")
    c = ws.cell(row=row, column=c1, value=val)
    c.font      = Font(name=FN, bold=bold, size=sz, color=fg)
    c.fill      = BG(bg)
    c.alignment = Alignment(horizontal=ha, vertical="center", wrap_text=wrap)
    ws.row_dimensions[row].height = h


def CL(ws, row, col, val=None, bold=False, sz=10, fg="000000",
       bg=None, ha="center", wrap=False, fmt=None, thick=False):
    c = ws.cell(row=row, column=col, value=val)
    c.font      = Font(name=FN, bold=bold, size=sz, color=fg)
    c.alignment = Alignment(horizontal=ha, vertical="center", wrap_text=wrap)
    if bg:   c.fill = BG(bg)
    if fmt:  c.number_format = fmt
    c.border = BDB() if thick else BD()
    return c


# ══════════════════════════════════════════════════════════
# レポート生成
# ══════════════════════════════════════════════════════════

def generate(excel_path, focus_start_str, output_path):
    print("  データ読み込み・クラスタリング中...")
    all_df      = load_and_cluster(excel_path)
    focus_start = pd.to_datetime(focus_start_str)
    focus_end   = all_df["date"].max()
    focus_df    = all_df[all_df["date"] >= focus_start].copy()

    n_days  = focus_df["date"].dt.date.nunique()
    period  = (f"{focus_start.strftime('%Y/%m/%d')} 〜 "
               f"{focus_end.strftime('%Y/%m/%d')}（{n_days}日間）")

    print("  集計中...")
    g, total = aggregate(focus_df, all_df)
    n_rows   = min(12, len(g))

    top3 = g.head(3)
    top3_text = "　".join([
        f"① {top3.iloc[0]['task']}（{top3.iloc[0]['時間']:.0f}h）",
        f"② {top3.iloc[1]['task']}（{top3.iloc[1]['時間']:.0f}h）",
        f"③ {top3.iloc[2]['task']}（{top3.iloc[2]['時間']:.0f}h）",
    ])
    top3pt = int(round(top3["時間"].sum() / total * 100))

    wb = Workbook()
    ws = wb.active
    ws.title = "業務時間ランキング"
    ws.sheet_view.showGridLines = False

    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 30   # 業務名
    ws.column_dimensions["C"].width = 14   # 時間
    ws.column_dimensions["D"].width = 11   # 割合
    ws.column_dimensions["E"].width = 12   # 日数
    ws.column_dimensions["F"].width = 14   # 担当区分
    ws.column_dimensions["G"].width = 45   # ★新規：主な作業内容（備考）

    # ── 行1, 行2: タイトル・結論（マージ幅をG列まで拡張） ──
    MR(ws, 1, 2, 7,
       f"業務時間ランキング　{period}",
       bold=True, sz=14, h=35, wrap=True)

    MR(ws, 2, 2, 7,
       f"結論：{top3_text}\n─ 上位3件で全体の{top3pt}%",
       bold=True, sz=11, fg=COL["ans_fg"], bg=COL["ans_bg"], h=50, wrap=True)

    ws.row_dimensions[3].height = 8

    # ── 行4: ヘッダー（G列追加） ──
    HDR = 4
    for col, txt in [(2,"業務の大分類"), (3,"合計時間\n（時間）"),
                     (4,"全体の\n割合"), (5,"発生した\n延べ日数"), 
                     (6,"担当区分"), (7,"主な作業内容\n（備考欄より自動抽出）")]:
        c = ws.cell(row=HDR, column=col, value=txt)
        c.font      = Font(name=FN, bold=True, size=10, color="FFFFFF")
        c.fill      = BG(COL["hdr"])
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border    = BDB()
    ws.row_dimensions[HDR].height = 35

    TBL_START = 5
    TOT = TBL_START + n_rows  
    p_idx = c_idx = 0

    # オートフィルターをG列まで適用
    ws.auto_filter.ref = f"B{HDR}:G{TOT-1}"

    for ri in range(n_rows):
        row = g.iloc[ri]
        R   = TBL_START + ri
        h   = row["時間"]
        personal = row["個人"]
        new_flag = row["新規"]
        fw = ri < 3

        if personal:
            bg      = COL["p_row"][p_idx % 3]; p_idx += 1
            fg_name = COL["personal"]
            z_val   = "個人業務"
            z_bg    = COL["p_badge_bg"]
            z_fg    = COL["personal"]
        else:
            bg      = COL["c_row"][c_idx % 3]; c_idx += 1
            fg_name = COL["common"]
            z_val   = "課内共通"
            z_bg    = COL["c_badge_bg"]
            z_fg    = COL["common"]

        name = row["task"]
        if len(name) > 22: name = name[:22] + "…"
        mark = "  ★" if new_flag else ""

        CL(ws, R, 2, name + mark, bold=fw, sz=10, fg=fg_name, bg=bg, ha="left", wrap=True)
        CL(ws, R, 3, h,           bold=fw, sz=12, fg=fg_name, bg=bg, fmt='0.0" h"')
        
        fg_ratio = "C00000" if h/total >= 0.15 else "000000"
        CL(ws, R, 4, f"=C{R}/$C${TOT}", bold=fw, sz=10, bg=bg, fmt="0%", fg=fg_ratio)
        
        CL(ws, R, 5, int(row["日数"]), sz=10, bg=bg, fmt='0" 日"')
        CL(ws, R, 6, z_val, bold=True, sz=9, fg=z_fg, bg=z_bg, ha="center")
        
        # ★新規：主な作業内容（備考）を書き込む
        memo_text = row["主な作業"]
        CL(ws, R, 7, memo_text, sz=9, fg=fg_name, bg=bg, ha="left", wrap=True)
        
        # 備考の行数に合わせてセルの高さを自動調整（文字が途切れないようにする）
        lines = memo_text.count('\n') + 1
        ws.row_dimensions[R].height = max(26, lines * 16)

    # ── 合計行（G列まで拡張） ──
    CL(ws, TOT, 2, "合　計", bold=True, sz=10, fg="FFFFFF", bg=COL["title"], ha="left", thick=True)
    CL(ws, TOT, 3, f"=SUBTOTAL(9, C{TBL_START}:C{TOT-1})", bold=True, sz=12, fg="FFFFFF", bg=COL["title"], fmt='0.0" h"', thick=True)
    CL(ws, TOT, 4, f"=SUBTOTAL(9, D{TBL_START}:D{TOT-1})", bold=True, sz=10, fg="FFFFFF", bg=COL["title"], fmt="0%", thick=True)
    CL(ws, TOT, 5, f"=SUBTOTAL(9, E{TBL_START}:E{TOT-1})", bold=True, sz=10, fg="FFFFFF", bg=COL["title"], fmt='0" 日"', thick=True)
    CL(ws, TOT, 6, "",       bold=True, sz=10, fg="FFFFFF", bg=COL["title"], thick=True)
    CL(ws, TOT, 7, "",       bold=True, sz=10, fg="FFFFFF", bg=COL["title"], thick=True)
    ws.row_dimensions[TOT].height = 26

    # ── 凡例（マージ幅をG列まで拡張） ──
    NOTE = TOT + 2
    ws.row_dimensions[NOTE - 1].height = 6
    MR(ws, NOTE, 2, 7,
       "🔴 個人業務 = 重点目標（BrandA ProgramB・BRANDBチェック・ProgramA等）　　🔵 課内共通 = チームの定常業務\n"
       "★ = 全期間と比べてこの期間に突出して増えた業務",
       bold=False, sz=9, fg="595959", bg=COL["note"], h=40, wrap=True)

    # ── グラフ見出し（マージ幅をG列まで拡張） ──
    GRAPH_HDR = NOTE + 2
    ws.row_dimensions[GRAPH_HDR - 1].height = 6
    MR(ws, GRAPH_HDR, 2, 7,
       "▼ 個人業務（重点目標）に注力できているかの確認グラフ",
       bold=True, sz=10, bg=COL["hdr"], h=24)

    # ══════════════════════════════════════════════════
    # グラフ作成（前回の修正のまま維持）
    # ══════════════════════════════════════════════════
    print("  グラフ作成中...")
    GRAPH_ROW = GRAPH_HDR + 1

    DATA_COL = 27
    ws.cell(row=HDR, column=DATA_COL,     value="業務名")
    ws.cell(row=HDR, column=DATA_COL + 1, value="時間")

    for ri in range(n_rows):
        row  = g.iloc[ri]
        R    = TBL_START + ri
        short_name = row["task"] if len(row["task"]) <= 15 else row["task"][:14] + "…"
        ws.cell(row=R, column=DATA_COL,     value=short_name)
        ws.cell(row=R, column=DATA_COL + 1, value=row["時間"])

    ch = BarChart()
    ch.type    = "bar"
    ch.legend  = None
    ch.title   = None
    ch.style   = 10
    ch.width   = 24
    ch.height  = 16

    ch.y_axis.scaling.orientation = "maxMin"

    ch.dataLabels = DataLabelList()
    ch.dataLabels.showVal = True

    dr = Reference(ws, min_col=DATA_COL + 1, max_col=DATA_COL + 1, min_row=HDR, max_row=HDR + n_rows)
    cr = Reference(ws, min_col=DATA_COL,     max_col=DATA_COL,     min_row=HDR + 1, max_row=HDR + n_rows)
    
    ch.add_data(dr, titles_from_data=True)
    ch.set_categories(cr)

    for ri in range(n_rows):
        personal = g.iloc[ri]["個人"]
        color    = COL["personal"] if personal else COL["common"]
        dp = DataPoint(idx=ri)
        dp.spPr.solidFill = color
        ch.series[0].dPt.append(dp)

    ch.series[0].graphicalProperties.solidFill = COL["common"]

    ws.add_chart(ch, f"B{GRAPH_ROW}")

    wb.save(output_path)
    return g, total, n_days, period, top3_text, top3pt


# ══════════════════════════════════════════════════════════
# GUIダイアログ関連
# ══════════════════════════════════════════════════════════

def select_file_gui():
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    file_path = filedialog.askopenfilename(
        title="日報Excelファイルを選択してください",
        filetypes=[("Excelファイル", "*.xlsx")]
    )
    return file_path

def select_date_gui():
    root = tk.Tk()
    root.title("集計期間の選択")
    
    window_width = 380
    window_height = 200
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    center_x = int(screen_width/2 - window_width / 2)
    center_y = int(screen_height/2 - window_height / 2)
    root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    
    root.attributes('-topmost', True)
    root.update()
    root.attributes('-topmost', False)

    today = datetime.today()
    w4 = (today - timedelta(weeks=4)).strftime('%Y-%m-%d')
    w2 = (today - timedelta(weeks=2)).strftime('%Y-%m-%d')
    w1 = (today - timedelta(weeks=1)).strftime('%Y-%m-%d')
    this_mon = (today - timedelta(days=today.weekday())).strftime('%Y-%m-%d')
    this_month_1 = today.replace(day=1).strftime('%Y-%m-%d')
    last_month_1 = (today.replace(day=1) - timedelta(days=1)).replace(day=1).strftime('%Y-%m-%d')

    options_dict = {
        f"過去4週間 ({w4})": w4,
        f"過去2週間 ({w2})": w2,
        f"過去1週間 ({w1})": w1,
        f"今週月曜 ({this_mon})": this_mon,
        f"当月1日 ({this_month_1})": this_month_1,
        f"前月1日 ({last_month_1})": last_month_1,
        "全期間 (2000-01-01)": "2000-01-01"
    }

    result = [w4] 

    def on_ok():
        selected = combo.get()
        if selected in options_dict:
            result[0] = options_dict[selected]
        else:
            result[0] = w4
        root.destroy()
        
    def on_closing():
        result[0] = w4
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)

    frame = ttk.Frame(root, padding=20)
    frame.pack(fill=tk.BOTH, expand=True)

    lbl = ttk.Label(frame, text="集計を開始する日付を選んでください：", font=("Meiryo UI", 10))
    lbl.pack(pady=(0, 15))

    combo = ttk.Combobox(frame, values=list(options_dict.keys()), state="readonly", width=35, font=("Meiryo UI", 10))
    combo.current(0)
    combo.pack(pady=5)

    btn = ttk.Button(frame, text="OK", command=on_ok)
    btn.pack(pady=20)

    root.mainloop()
    return result[0]


# ══════════════════════════════════════════════════════════
# メイン（対話 or 引数）
# ══════════════════════════════════════════════════════════

def main():
    TODAY = datetime.today()

    if len(sys.argv) >= 2:
        excel_path = sys.argv[1]
    else:
        candidates = [f for f in os.listdir(".") if f.endswith(".xlsx") and "業務時間" in f and not f.startswith("~")]
        if len(candidates) == 1:
            excel_path = candidates[0]
            print(f"入力ファイルを自動検出: {excel_path}")
        else:
            print("ファイル選択ダイアログを開いています...")
            excel_path = select_file_gui()
            if not excel_path:
                print("キャンセルされました。処理を終了します。")
                sys.exit(0)

    if not os.path.exists(excel_path):
        print(f"エラー: ファイルが見つかりません → {excel_path}")
        sys.exit(1)

    if len(sys.argv) >= 3:
        focus_start_str = sys.argv[2]
    else:
        print("日付選択ダイアログを開いています...")
        focus_start_str = select_date_gui()

    out_dir  = os.path.dirname(os.path.abspath(excel_path))
    out_name = f"業務報告_承認担当者向け_{TODAY.strftime('%m%d')}.xlsx"
    output_path = os.path.join(out_dir, out_name)

    print()
    print(f"  入力: {excel_path}")
    print(f"  期間: {focus_start_str} 〜")
    print(f"  出力: {output_path}")
    print()

    g, total, n_days, period, top3_text, top3pt = generate(
        excel_path, focus_start_str, output_path
    )

    print("=" * 60)
    print(f"  ✅ 完成: {out_name}")
    print()
    print(f"  期間: {period}")
    print(f"  結論: {top3_text}")
    print(f"        上位3件で全体の {top3pt}%")
    print()
    print("  【個人業務】")
    for _, row in g.head(12).iterrows():
        if row["個人"]:
            new = "★" if row["新規"] else " "
            print(f"    🔴{new} {row['task']:<32} {row['時間']:5.1f}h  "
                  f"{row['時間']/total*100:.0f}%  {row['日数']:.0f}日")
    print()
    print("  【課内共通】")
    for _, row in g.head(12).iterrows():
        if not row["個人"]:
            new = "★" if row["新規"] else " "
            print(f"    🔵{new} {row['task']:<32} {row['時間']:5.1f}h  "
                  f"{row['時間']/total*100:.0f}%  {row['日数']:.0f}日")
    print("=" * 60)
    print()
    print("  Excelファイルを開きます...")
    import subprocess, platform
    if platform.system() == "Windows":
        os.startfile(output_path)
    elif platform.system() == "Darwin":
        subprocess.run(["open", output_path])


if __name__ == "__main__":
    main()