# =============================================================
# SupportDesk 返信ドラフト自動生成ツール  v3.3 (徹底精査モード)
# Example Company 保証課
# =============================================================

import os, re, json, hashlib, datetime, sys, queue, threading, time
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from tkcalendar import DateEntry

import win32com.client
import pythoncom
import pdfplumber
from google import genai

# =============================================================
# ★★★ 設定セクション（ここだけ変更）★★★
# =============================================================

GEMINI_API_KEY  = os.environ.get("GEMINI_API_KEY", "")
PDF_FOLDER      = r"C:\Users\Public\Desktop\SupportDesk\マニュアル"
OUTPUT_FOLDER   = r"C:\Users\Public\Desktop\SupportDesk\Drafts"
CACHE_FOLDER    = r"C:\Users\Public\Desktop\SupportDesk\cache"
SupportDesk_FOLDER_NAME = "SupportDesk"

MY_ADDRESS      = "user@example.com"
PAST_MAIL_DAYS  = 730
MAX_PAST_MAILS  = 8
MAX_PDF_CHARS   = 200000  # ←【変更】マニュアルを最後まで全読みさせる（20万文字）
MAX_BODY_CHARS  = 2000

TRIGGER_NAMES   = ["ご担当者様", "伊澤様", "河合様", "秀介様", "担当者様"]

MAIL_CACHE_FILE = os.path.join(CACHE_FOLDER, "mail_index.json")
PDF_CACHE_FILE  = os.path.join(CACHE_FOLDER, "pdf_cache.json")

# =============================================================
# ログキュー
# =============================================================

_log_queue = queue.Queue()

def log(msg):
    print(msg)
    _log_queue.put(msg)

# =============================================================
# GUI: 日時範囲ピッカー
# =============================================================

def ask_datetime_range_gui(root):
    result = {"start": None, "end": None, "ok": False}
    win    = tk.Toplevel(root)
    win.title("SupportDesk ドラフト生成ツール v3.3")
    win.resizable(False, False)
    win.grab_set()

    BG = "#f5f5f5"; ACCENT = "#1a6faf"
    FONT_H = ("Meiryo UI", 11, "bold")
    FONT_N = ("Meiryo UI", 10)
    FONT_S = ("Meiryo UI",  9)
    win.configure(bg=BG)

    hdr = tk.Frame(win, bg=ACCENT, padx=16, pady=10)
    hdr.pack(fill="x")
    tk.Label(hdr, text="SupportDesk 返信ドラフト自動生成ツール", bg=ACCENT, fg="white", font=("Meiryo UI", 13, "bold")).pack(anchor="w")
    tk.Label(hdr, text="Example Company  保証課", bg=ACCENT, fg="#cde4f7", font=FONT_S).pack(anchor="w")

    body = tk.Frame(win, bg=BG, padx=20, pady=14)
    body.pack(fill="both")
    tk.Label(body, text="処理対象メールの受信日時範囲を指定してください", bg=BG, fg="#333", font=FONT_H).grid(row=0, column=0, columnspan=7, sticky="w", pady=(0, 12))

    def make_row(parent, row, label):
        tk.Label(parent, text=label, bg=BG, font=FONT_N, width=6, anchor="e").grid(row=row, column=0, sticky="e", padx=(0, 8))
        cal = DateEntry(parent, width=12, background=ACCENT, foreground="white", borderwidth=2, font=FONT_N, date_pattern="yyyy-mm-dd", locale="ja_JP")
        cal.grid(row=row, column=1, padx=(0, 8))
        tk.Label(parent, text="時", bg=BG, font=FONT_N).grid(row=row, column=2)
        hv = tk.StringVar(value="00")
        hcb = ttk.Combobox(parent, textvariable=hv, width=4, values=[f"{h:02d}" for h in range(24)], state="readonly", font=FONT_N)
        hcb.grid(row=row, column=3, padx=4)
        tk.Label(parent, text="分", bg=BG, font=FONT_N).grid(row=row, column=4)
        mv = tk.StringVar(value="00")
        mcb = ttk.Combobox(parent, textvariable=mv, width=4, values=["00","15","30","45"], state="readonly", font=FONT_N)
        mcb.grid(row=row, column=5, padx=4)
        nolim = tk.BooleanVar(value=False)
        def toggle(c=cal, h=hcb, m=mcb, v=nolim):
            s = "disabled" if v.get() else "normal"
            c.config(state=s)
            h.config(state="disabled" if v.get() else "readonly")
            m.config(state="disabled" if v.get() else "readonly")
        tk.Checkbutton(parent, text="制限なし", variable=nolim, bg=BG, font=FONT_S, command=toggle).grid(row=row, column=6, padx=(10, 0))
        return cal, hv, mv, nolim

    sc, sh, sm, sn = make_row(body, 1, "開始")
    tk.Frame(body, bg=BG, height=8).grid(row=2)
    ec, eh, em, en = make_row(body, 3, "終了")

    btn_frame = tk.Frame(win, bg=BG, padx=20, pady=12)
    btn_frame.pack(fill="x")

    def on_run():
        try:
            result["start"] = None if sn.get() else datetime.datetime(*sc.get_date().timetuple()[:3], int(sh.get()), int(sm.get()))
            result["end"]   = None if en.get() else datetime.datetime(*ec.get_date().timetuple()[:3], int(eh.get()), int(em.get()))
            if result["start"] and result["end"] and result["start"] > result["end"]:
                messagebox.showerror("入力エラー", "開始日時が終了日時より後になっています。")
                return
            result["ok"] = True
            win.destroy()
        except Exception as ex:
            messagebox.showerror("エラー", str(ex))

    tk.Button(btn_frame, text="　実　行　", bg=ACCENT, fg="white", font=FONT_H, relief="flat", padx=12, pady=6, cursor="hand2", command=on_run).pack(side="right", padx=(8,0))
    tk.Button(btn_frame, text="キャンセル", bg="#ccc", fg="#333", font=FONT_N, relief="flat", padx=10, pady=6, cursor="hand2", command=win.destroy).pack(side="right")

    root.wait_window(win)
    if not result["ok"]: return None, None
    return result["start"], result["end"]

# =============================================================
# GUI: リアルタイムログウィンドウ
# =============================================================

def show_log_window(root, title="SupportDesk 処理中..."):
    win = tk.Toplevel(root)
    win.title(title)
    win.geometry("700x450")
    win.resizable(True, True)

    BG = "#1e1e1e"; FG = "#d4d4d4"; ACCENT = "#1a6faf"
    FONT_H = ("Meiryo UI", 11, "bold")
    FONT_M = ("Consolas", 9)
    win.configure(bg=BG)

    hdr = tk.Frame(win, bg=ACCENT, padx=12, pady=8)
    hdr.pack(fill="x")
    tk.Label(hdr, text="SupportDesk 処理ログ", bg=ACCENT, fg="white", font=FONT_H).pack(side="left")
    status_var = tk.StringVar(value="● 処理中...")
    status_lbl = tk.Label(hdr, textvariable=status_var, bg=ACCENT, fg="#ffe066", font=FONT_H)
    status_lbl.pack(side="right")

    txt = scrolledtext.ScrolledText(win, bg=BG, fg=FG, font=FONT_M, wrap="word", state="disabled", insertbackground=FG)
    txt.pack(fill="both", expand=True, padx=8, pady=8)

    txt.tag_config("info",    foreground="#9cdcfe")
    txt.tag_config("ok",      foreground="#4ec9b0")
    txt.tag_config("warn",    foreground="#ce9178")
    txt.tag_config("error",   foreground="#f44747")
    txt.tag_config("section", foreground="#dcdcaa", font=("Consolas", 9, "bold"))

    btn_frame = tk.Frame(win, bg=BG, pady=6)
    btn_frame.pack(fill="x")
    close_btn = tk.Button(btn_frame, text="閉じる", bg="#555", fg="white", font=FONT_H, relief="flat", padx=16, pady=4, cursor="hand2", state="disabled", command=root.quit)
    close_btn.pack(side="right", padx=10)

    def append(msg):
        txt.config(state="normal")
        tag = "info"
        if "[エラー]" in msg or "[ERROR]" in msg:  tag = "error"
        elif "[警告]" in msg:                       tag = "warn"
        elif "完了" in msg or "保存" in msg or "→" in msg: tag = "ok"
        elif msg.startswith("[") or msg.startswith("="): tag = "section"
        txt.insert("end", msg + "\n", tag)
        txt.see("end")
        txt.config(state="disabled")

    def poll():
        try:
            while True:
                msg = _log_queue.get_nowait()
                if msg == "__DONE__":
                    status_var.set("✔ 完了")
                    status_lbl.config(fg="#4ec9b0")
                    hdr.config(bg="#2d7a4f")
                    status_lbl.config(bg="#2d7a4f")
                    close_btn.config(state="normal", bg=ACCENT)
                    append("=" * 50)
                    append("  処理が完了しました。「閉じる」を押してください。")
                    append("=" * 50)
                    return
                elif msg == "__ERROR__":
                    status_var.set("✘ エラー")
                    close_btn.config(state="normal", bg="#c0392b")
                    return
                else:
                    append(msg)
        except queue.Empty:
            pass
        win.after(100, poll)
    win.after(100, poll)

# =============================================================
# キャッシュ: PDF
# =============================================================

def md5(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""): h.update(chunk)
    return h.hexdigest()

def load_pdf_manuals():
    os.makedirs(CACHE_FOLDER, exist_ok=True)
    cache = {}
    if os.path.exists(PDF_CACHE_FILE):
        try:
            with open(PDF_CACHE_FILE, "r", encoding="utf-8") as f: cache = json.load(f)
        except Exception: cache = {}

    manuals = []
    updated = False

    if not os.path.exists(PDF_FOLDER):
        log(f"  [警告] PDFフォルダなし: {PDF_FOLDER}")
        return manuals

    for fname in [f for f in os.listdir(PDF_FOLDER) if f.lower().endswith(".pdf")]:
        fpath   = os.path.join(PDF_FOLDER, fname)
        cur_md5 = md5(fpath)
        if fname in cache and cache[fname]["hash"] == cur_md5:
            log(f"  [PDF cached] {fname}")
            manuals.append({"filename": fname, "text": cache[fname]["text"]})
            continue

        try:
            parts = []
            with pdfplumber.open(fpath) as pdf:
                for page in pdf.pages:
                    t = page.extract_text()
                    if t: parts.append(t)
            full = "\n".join(parts)
            text = full[:MAX_PDF_CHARS]
            cache[fname] = {"hash": cur_md5, "text": text}
            manuals.append({"filename": fname, "text": text})
            updated = True
            log(f"  [PDF new]    {fname} ({len(full):,}文字完了)")
        except Exception as e:
            log(f"  [PDF ERROR]  {fname}: {e}")

    if updated:
        with open(PDF_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
    return manuals

# =============================================================
# キャッシュ: メールインデックス
# =============================================================

def _mail_to_dict(item):
    try:
        return {
            "subject":  item.Subject or "",
            "sender":   item.SenderName or "",
            "received": normalize_dt(item.ReceivedTime).isoformat(),
            "snippet":  (item.Body or "")[:600]
        }
    except Exception: return None

def load_mail_cache():
    if os.path.exists(MAIL_CACHE_FILE):
        try:
            with open(MAIL_CACHE_FILE, "r", encoding="utf-8") as f: return json.load(f)
        except Exception: pass
    return {"last_scan": None, "mails": []}

def save_mail_cache(cache):
    os.makedirs(CACHE_FOLDER, exist_ok=True)
    with open(MAIL_CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

def update_mail_index(inbox):
    cache     = load_mail_cache()
    last_scan = cache["last_scan"]
    now       = datetime.datetime.now()
    cutoff    = now - datetime.timedelta(days=PAST_MAIL_DAYS)

    if last_scan:
        since = datetime.datetime.fromisoformat(last_scan)
        log(f"  メールキャッシュ更新: {since.strftime('%Y-%m-%d %H:%M')} 以降の差分のみ取得")
    else:
        since = cutoff
        log(f"  メールキャッシュ初回構築: 過去{PAST_MAIL_DAYS}日分をスキャン（初回のみ時間がかかります）")

    new_mails = []
    def scan(folder, depth=0):
        if depth > 4: return
        try:
            for item in folder.Items:
                if item.Class != 43: continue
                try:
                    received = normalize_dt(item.ReceivedTime)
                    if received < since: continue
                    if received > now:   continue
                    d = _mail_to_dict(item)
                    if d: new_mails.append(d)
                except Exception: pass
        except Exception: pass
        try:
            for sub in folder.Folders: scan(sub, depth+1)
        except Exception: pass

    scan(inbox)
    log(f"  → 新着 {len(new_mails)} 件取得")

    existing = [m for m in cache["mails"] if datetime.datetime.fromisoformat(m["received"]) >= cutoff]
    merged   = existing + new_mails
    seen, dedup = set(), []
    for m in merged:
        key = (m["subject"], m["received"])
        if key not in seen:
            seen.add(key)
            dedup.append(m)

    cache["mails"]     = dedup
    cache["last_scan"] = now.isoformat()
    save_mail_cache(cache)
    log(f"  → キャッシュ合計: {len(dedup)} 件")
    return cache["mails"]

# =============================================================
# Outlook操作
# =============================================================

def connect_outlook():
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns      = outlook.GetNamespace("MAPI")
    return ns.GetDefaultFolder(6), ns.GetDefaultFolder(5)

def find_whd_folder(inbox):
    for folder in inbox.Folders:
        if folder.Name == SupportDesk_FOLDER_NAME: return folder
    return None

def normalize_dt(dt):
    try: return dt.replace(tzinfo=None)
    except Exception: return dt

def is_trigger_mail(mail, start_dt, end_dt):
    try:
        if mail.Class != 43: return False
        received = normalize_dt(mail.ReceivedTime)
        if start_dt and received < start_dt: return False
        if end_dt   and received > end_dt:   return False
        
        # 【変更】検索範囲を200文字に戻す（空白無視の機能は維持）
        body_text = (mail.Body or "")[:200]
        body_clean = body_text.replace(" ", "").replace("　", "")
        
        return any(name in body_clean for name in TRIGGER_NAMES)
    except Exception: return False

def get_trigger_mails(whd_folder, start_dt, end_dt):
    result = [item for item in whd_folder.Items if is_trigger_mail(item, start_dt, end_dt)]
    result.sort(key=lambda m: normalize_dt(m.ReceivedTime), reverse=True)
    return result

def extract_sender_firstname(mail):
    sender = mail.SenderName or ""
    name   = sender.split()[0] if sender.split() else sender
    name   = re.sub(r'[,，、]', '', name)
    name   = re.sub(r'(様|さん)$', '', name)
    return name or "ご担当者"

# =============================================================
# 過去メール検索
# =============================================================

def extract_keywords(subject):
    cleaned = re.sub(r'^(Re:|Fw:|FW:|RE:|【[^】]*】|\[[^\]]*\])\s*', '', subject, flags=re.IGNORECASE).strip()
    return re.findall(r'[^\s\u3000\u3001\u3002\uff0c\uff0e\u300c\u300d\u3010\u3011\(\)\[\]]{2,}', cleaned)[:6]

def search_similar_mails_from_cache(mail_index, subject):
    keywords = extract_keywords(subject)
    if not keywords: return []
    scored = []
    for m in mail_index:
        score = sum(1 for kw in keywords if kw in m["subject"] or kw in m["snippet"])
        if score > 0: scored.append((score, m))
    scored.sort(key=lambda x: x[0], reverse=True)
    return [m for _, m in scored[:MAX_PAST_MAILS]]

# =============================================================
# Gemini API 呼び出し（プロンプト強化版）
# =============================================================

def build_prompt(mail, past_mails, manuals, firstname):
    # 【変更】プロンプトを徹底精査モードに強化
    lines = [
        "あなたはExample Company保証課の経験豊富なSupportDesk担当者です。",
        "以下の【参照マニュアル】の規程内容を隅々まで徹底的に精査し、【過去類似案件】の対応履歴も踏まえた上で、ディーラー担当者への正確な回答メール本文を作成してください。",
        "※推測での回答は厳禁です。必ずマニュアルの該当ルールに則って論理的に記載してください。",
        "",
        "【出力フォーマット厳守】",
        f"{firstname}さん", "",
        "お世話になります。",
        "（回答本文：マニュアルの根拠を示しつつ、簡潔・丁寧・事実ベースで記載）", "",
        "よろしくお願いします。", "",
        "■参照情報",
        "（使用した類似案件の件名・日付、および根拠としたマニュアルのファイル名と該当箇所を具体的に記載）", "",
        "■確認事項（なければ省略）",
        "（ディーラーに追加確認が必要な不足情報があれば記載）", "",
        "---",
        "件名・署名・宛先は不要。本文のみ出力。",
        "=" * 50,
        "【対応メール】",
        f"件名: {mail.Subject}",
        f"差出人: {mail.SenderName}  受信: {mail.ReceivedTime}",
        "本文:", (mail.Body or "")[:MAX_BODY_CHARS], "",
        "=" * 50,
        f"【過去類似案件】（{len(past_mails)}件）",
    ]
    if past_mails:
        for i, m in enumerate(past_mails, 1):
            lines += [f"\n--- 類似{i} ---", f"件名: {m['subject']}  日時: {m['received']}", f"本文抜粋: {m['snippet'][:500]}"]
    else: lines.append("（なし）")

    lines += ["", "=" * 50, f"【参照マニュアル】（全文データ）"]
    if manuals:
        for m in manuals:
            lines += [f"\n--- {m['filename']} ---", m["text"]]
    else: lines.append("（なし）")
    return "\n".join(lines)

def call_gemini(prompt):
    client   = genai.Client(api_key=GEMINI_API_KEY)
    # 【変更】Lite版から標準モデル（より賢い推論力）へ変更
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt
    )
    return response.text

# =============================================================
# TXT出力・削除処理
# =============================================================

def sanitize(text):
    return re.sub(r'[\\/:*?"<>|\r\n\t]', '_', text).strip()

def save_draft(mail, draft_text, past_mails, manuals):
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    base = sanitize(mail.Subject or "無題")
    path = os.path.join(OUTPUT_FOLDER, f"{base}.txt")
    n = 1
    while os.path.exists(path):
        path = os.path.join(OUTPUT_FOLDER, f"{base}_{n}.txt"); n += 1
    header = "\n".join([
        "=" * 60, "  SupportDesk 返信ドラフト（自動生成）", "=" * 60,
        f"生成日時      : {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"対象件名      : {mail.Subject}",
        f"差出人        : {mail.SenderName}",
        f"受信日時      : {mail.ReceivedTime}",
        f"類似案件参照  : {len(past_mails)}件",
        f"PDFマニュアル : {len(manuals)}件",
        "=" * 60, "",
    ])
    with open(path, "w", encoding="utf-8") as f: f.write(header + draft_text)
    return path

def cleanup_replied_drafts(sent_box):
    if not os.path.exists(OUTPUT_FOLDER): return
    deleted = 0
    for fname in list(os.listdir(OUTPUT_FOLDER)):
        if not fname.endswith(".txt"): continue
        subject_guess = re.sub(r'_\d+$', '', fname[:-4])
        keywords = extract_keywords(subject_guess)
        if not keywords: continue
        try:
            for item in sent_box.Items:
                if item.Class != 43: continue
                addr = getattr(item, "SenderEmailAddress", "") or ""
                if MY_ADDRESS.lower() not in addr.lower(): continue
                if any(kw in (item.Subject or "") for kw in keywords):
                    os.remove(os.path.join(OUTPUT_FOLDER, fname))
                    log(f"  [削除] 返信済み検出 → {fname}")
                    deleted += 1; break
        except Exception: pass
    if deleted == 0: log("  → 削除対象なし")

# =============================================================
# メイン処理
# =============================================================

def run(start_dt, end_dt, auto_mode=False):
    pythoncom.CoInitialize()
    try:
        log(f"  Mode: {'AUTO' if auto_mode else 'MANUAL'}")
        log(f"  開始: {start_dt.strftime('%Y-%m-%d %H:%M') if start_dt else '制限なし'}")
        log(f"  終了: {end_dt.strftime('%Y-%m-%d %H:%M') if end_dt else '制限なし'}")

        log("\n[1/5] PDFマニュアル読み込み中（キャッシュ対応）...")
        manuals = load_pdf_manuals()
        log(f"  → {len(manuals)}件完了")

        log("\n[2/5] Outlook接続中...")
        inbox, sent_box = connect_outlook()
        whd_folder = find_whd_folder(inbox)
        if whd_folder is None:
            log(f"  [エラー] '{SupportDesk_FOLDER_NAME}' フォルダが見つかりません。")
            _log_queue.put("__ERROR__")
            return
        log("  → 接続成功")

        log("\n[3/5] メールインデックス更新中（差分のみ）...")
        mail_index = update_mail_index(inbox)

        log("\n[4/5] 返信済みドラフト削除チェック...")
        cleanup_replied_drafts(sent_box)

        log("\n[5/5] トリガーメール確認中...")
        trigger_mails = get_trigger_mails(whd_folder, start_dt, end_dt)
        log(f"  → 対象: {len(trigger_mails)}件")

        if not trigger_mails:
            log("\n処理対象メールはありませんでした。")
            _log_queue.put("__DONE__")
            return

        success = 0
        for i, mail in enumerate(trigger_mails, 1):
            log(f"\n--- {i}/{len(trigger_mails)}件目 ---")
            log(f"  件名: {mail.Subject}")
            try:
                firstname  = extract_sender_firstname(mail)
                past_mails = search_similar_mails_from_cache(mail_index, mail.Subject)
                log(f"  → 類似 {len(past_mails)}件 / 宛名: {firstname}さん")
                log("  → Gemini呼び出し中（徹底精査中...少し時間がかかります）...")
                
                draft = call_gemini(build_prompt(mail, past_mails, manuals, firstname))
                path  = save_draft(mail, draft, past_mails, manuals)
                log(f"  → 保存完了: {path}")
                success += 1
                
                # 【変更】大量のPDFデータを送るため、API制限を回避する10秒のインターバルを追加
                if i < len(trigger_mails):
                    log("  → API制限回避のため10秒待機します...")
                    time.sleep(10)

            except Exception as e:
                log(f"  [エラー] {e}")

        log(f"\n{'='*50}")
        log(f"完了: {success}/{len(trigger_mails)}件処理")
        log(f"出力先: {OUTPUT_FOLDER}")
        log("=" * 50)
        _log_queue.put("__DONE__")

    except Exception as e:
        log(f"\n[致命的エラー] {e}")
        _log_queue.put("__ERROR__")

# =============================================================
# エントリーポイント
# =============================================================

def main():
    print("=" * 60)
    print("  SupportDesk 返信ドラフト自動生成ツール  v3.3 (徹底精査モード)")
    print("=" * 60)

    root = tk.Tk()
    root.withdraw()

    log("  [手動モード] 日時範囲をGUIで指定してください")
    start_dt, end_dt = ask_datetime_range_gui(root)

    show_log_window(root, "SupportDesk Manual (徹底精査)")
    t = threading.Thread(target=run, args=(start_dt, end_dt, False), daemon=True)
    t.start()
    root.mainloop()

if __name__ == "__main__":
    main()
