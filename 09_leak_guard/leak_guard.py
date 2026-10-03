#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
leak_guard.py - 公開リポジトリへの機密情報混入を、コミット／プッシュの前に止めるスキャナ

検出するもの
  1. 組み込みルール : メールアドレス、Windowsのユーザーフォルダ、UNCパス、社内IP、
                      APIキー／トークン／秘密鍵、パスワード直書き
  2. 禁止語リスト   : 社名・人名・社内システム名・取引先コードなど、利用者が登録した語
                      （リポジトリ外のファイルに置く。リスト自体を公開しないため）
  3. 中身を検査できない文書ファイル（Excel / Word / PowerPoint / PDF / メール等）

使い方
  python leak_guard.py --staged        # ステージ済みの変更を検査（pre-commit フック用）
  python leak_guard.py --pre-push      # プッシュされるコミットを検査（pre-push フック用。stdin を読む）
  python leak_guard.py --all           # 追跡中の全ファイルを検査
  python leak_guard.py --history       # 全コミット履歴の追加行を検査（過去の混入確認）
  python leak_guard.py --path DIR      # Git 管理外のフォルダを検査

禁止語リストの場所（上から優先）
  --denylist で指定したファイル / 環境変数 LEAK_GUARD_DENYLIST / ~/.leak_guard/denylist.txt
  書式: 1行1語（大文字小文字は区別しない）。"re:" で始まる行は正規表現。"#" はコメント。

例外の指定
  行内に "leak-guard: allow" と書いた行は検査しない。
  リポジトリ直下の .leak_guard_allow に書いたパス（glob）は文書ファイルの検査を除外する。

終了コード: 0 = 問題なし / 1 = 検出あり（コミット・プッシュを中止） / 2 = 実行エラー
"""
import argparse
import fnmatch
import os
import re
import subprocess
import sys
from typing import Iterable, List, Optional, Tuple

ZERO_SHA = "0" * 40
ALLOW_MARK = "leak-guard: allow"

# ---------------------------------------------------------------------------
# 組み込みルール（名前, 正規表現）
# ---------------------------------------------------------------------------
BUILTIN_RULES: List[Tuple[str, "re.Pattern[str]"]] = [
    ("メールアドレス", re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")),
    ("Windowsユーザーフォルダ", re.compile(r"[A-Za-z]:\\{1,2}Users\\{1,2}(?!Public\\|public\\|<|%|\{)[^\\\s\"'*?]+", re.I)),
    ("UNCパス", re.compile(r"(?<![\\\w])\\\\[A-Za-z0-9_.-]{2,}\\[^\s\"']+")),
    ("社内IPアドレス", re.compile(r"\b(?:10\.\d{1,3}|192\.168|172\.(?:1[6-9]|2\d|3[01]))\.\d{1,3}\.\d{1,3}\b")),
    ("秘密鍵", re.compile(r"-----BEGIN [A-Z ]*PRIVATE KEY-----")),
    ("APIキー/トークン", re.compile(
        r"(?:AKIA[0-9A-Z]{16}|ghp_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{30,}"
        r"|AIza[0-9A-Za-z_-]{30,}|sk-(?:ant-|proj-)?[A-Za-z0-9_-]{20,}|xox[abprs]-[A-Za-z0-9-]{10,})")),
    ("パスワード直書き", re.compile(r"(?i)(?:password|passwd|pwd|パスワード)\s*[:=]\s*[\"'][^\"']{4,}[\"']")),
]

# 公開して問題ないメールアドレス（ダミー・noreply）
SAFE_EMAIL = re.compile(
    r"(?i)@(?:example\.(?:com|org|net|jp)|users\.noreply\.github\.com)$|^noreply@|^no-reply@")

# 中身をテキストとして検査できない文書・データファイル
BLOCKED_EXTENSIONS = {
    ".xlsx", ".xlsm", ".xlsb", ".xls", ".docx", ".docm", ".doc", ".pptx", ".pptm", ".ppt",
    ".pdf", ".msg", ".eml", ".pst", ".ost", ".accdb", ".mdb", ".zip", ".7z", ".lzh",
}


# ---------------------------------------------------------------------------
# 禁止語リスト
# ---------------------------------------------------------------------------
def load_denylist(path_arg: Optional[str]) -> Tuple[List[Tuple[str, "re.Pattern[str]"]], Optional[str]]:
    candidates = [path_arg, os.environ.get("LEAK_GUARD_DENYLIST"),
                  os.path.join(os.path.expanduser("~"), ".leak_guard", "denylist.txt")]
    path = next((p for p in candidates if p and os.path.isfile(p)), None)
    if not path:
        return [], None
    rules = []
    text, _ = decode(open(path, "rb").read())
    for raw in text.splitlines():
        line = raw.strip()
        if not line or line.startswith("#"):
            continue
        if line.startswith("re:"):
            rules.append(("禁止語(正規表現)", re.compile(line[3:], re.I)))
        else:
            rules.append(("禁止語", re.compile(re.escape(line), re.I)))
    return rules, path


# ---------------------------------------------------------------------------
# 共通処理
# ---------------------------------------------------------------------------
def decode(data: bytes) -> Tuple[str, bool]:
    """テキストに変換する。NUL を含むものはバイナリとみなす。"""
    if b"\x00" in data[:8000]:
        return "", True
    for enc in ("utf-8-sig", "cp932"):
        try:
            return data.decode(enc), False
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", "replace"), False


def mask(s: str) -> str:
    """検出値を画面に出すときは中央を伏せる（ログ経由の二次流出を防ぐ）。"""
    if len(s) <= 4:
        return s[0] + "*" * (len(s) - 1)
    keep = max(1, len(s) // 4)
    return s[:keep] + "*" * (len(s) - keep * 2) + s[-keep:]


def git(*args: str, input_bytes: Optional[bytes] = None) -> bytes:
    # core.quotepath=false: 日本語ファイル名を "\346\227..." 形式にエスケープさせない
    return subprocess.run(["git", "-c", "core.quotepath=false", *args], input=input_bytes,
                          check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE).stdout


def has_object(sha: str) -> bool:
    return subprocess.run(["git", "cat-file", "-e", f"{sha}^{{commit}}"],
                          stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL).returncode == 0


def diff_path(header: str) -> Optional[str]:
    """'+++ b/path' 行からパスを取り出す。引用符付き（特殊文字を含む）形式にも対応。"""
    p = header[4:].rstrip("\t")
    if p == "/dev/null":
        return None
    if p.startswith('"') and p.endswith('"'):
        raw = p[1:-1].encode("latin-1", "backslashreplace").decode("unicode_escape")
        p = raw.encode("latin-1", "replace").decode("utf-8", "replace")
    return p[2:] if p.startswith("b/") else p


def load_allowlist(root: str) -> List[str]:
    p = os.path.join(root, ".leak_guard_allow")
    if not os.path.isfile(p):
        return []
    text, _ = decode(open(p, "rb").read())
    return [l.strip() for l in text.splitlines() if l.strip() and not l.startswith("#")]


class Scanner:
    def __init__(self, rules, allow_globs: List[str]):
        self.rules = rules
        self.allow_globs = allow_globs
        self.findings: List[Tuple[str, str, str, str]] = []  # (場所, 行, ルール, 値)

    def is_allowed_path(self, path: str) -> bool:
        p = path.replace("\\", "/")
        return any(fnmatch.fnmatch(p, g) for g in self.allow_globs)

    def check_path(self, where: str, path: str) -> bool:
        """ファイル種別で判定。検査を続けてよければ True。"""
        ext = os.path.splitext(path)[1].lower()
        if ext in BLOCKED_EXTENSIONS and not self.is_allowed_path(path):
            self.findings.append((where, "-", "文書ファイル（中身を検査できない）", os.path.basename(path)))
            return False
        return True

    def scan_lines(self, where: str, lines: Iterable[Tuple[int, str]]) -> None:
        for lineno, line in lines:
            if ALLOW_MARK in line:
                continue
            for name, pat in self.rules:
                for m in pat.finditer(line):
                    val = m.group(0)
                    if name == "メールアドレス" and SAFE_EMAIL.search(val):
                        continue
                    self.findings.append((where, str(lineno), name, val))

    def scan_text_blob(self, where: str, path: str, data: bytes) -> None:
        if not self.check_path(where, path):
            return
        text, is_binary = decode(data)
        if is_binary:
            return
        self.scan_lines(where, enumerate(text.splitlines(), 1))
        # ファイル名そのものも検査する（社名・人名がファイル名に入る事故が多い）
        self.scan_lines(where + " (ファイル名)", [(0, path)])

    def scan_diff(self, label: str, diff: bytes) -> None:
        """git の unified diff から追加行だけを検査する。"""
        text, _ = decode(diff)
        current = None
        lineno = 0
        for line in text.splitlines():
            if line.startswith("+++ "):
                current = diff_path(line)
                if current:
                    self.scan_lines(f"{label}:{current} (ファイル名)", [(0, current)])
                    self.check_path(f"{label}:{current}", current)
                continue
            if line.startswith("@@"):
                m = re.search(r"\+(\d+)", line)
                lineno = int(m.group(1)) if m else 0
                continue
            if current and line.startswith("+"):
                self.scan_lines(f"{label}:{current}", [(lineno, line[1:])])
                lineno += 1
            elif current and not line.startswith("-"):
                lineno += 1


# ---------------------------------------------------------------------------
# 各モード
# ---------------------------------------------------------------------------
def mode_staged(sc: Scanner) -> None:
    names = git("diff", "--cached", "--name-only", "--diff-filter=ACMR", "-z").split(b"\x00")
    for raw in filter(None, names):
        path = raw.decode("utf-8", "replace")
        sc.scan_text_blob(path, path, git("show", f":{path}"))


def mode_all(sc: Scanner) -> None:
    for raw in filter(None, git("ls-files", "-z").split(b"\x00")):
        path = raw.decode("utf-8", "replace")
        if os.path.isfile(path):
            sc.scan_text_blob(path, path, open(path, "rb").read())


def mode_path(sc: Scanner, root: str) -> None:
    for dp, dirs, files in os.walk(root):
        dirs[:] = [d for d in dirs if d != ".git"]
        for f in files:
            p = os.path.join(dp, f)
            rel = os.path.relpath(p, root)
            sc.scan_text_blob(rel, rel, open(p, "rb").read())


def scan_commits(sc: Scanner, rev_args: List[str]) -> None:
    shas = git("rev-list", *rev_args).decode().split()
    for sha in shas:
        diff = git("show", "--format=", "--no-color", "-U0", "--no-renames", sha)
        sc.scan_diff(sha[:8], diff)
        # バイナリは diff に本文が出ないため、追加・変更ファイル名で種別を判定する
        names = git("show", "--format=", "--name-only", "--diff-filter=AM", "-z", sha)
        for raw in filter(None, names.split(b"\x00")):
            path = raw.decode("utf-8", "replace").strip()
            if path:
                sc.check_path(f"{sha[:8]}:{path}", path)
        # コミットメッセージと作成者メールも公開されるので検査する
        meta = git("show", "-s", "--format=%an <%ae>%n%B", sha)
        text, _ = decode(meta)
        sc.scan_lines(f"{sha[:8]} (コミット情報)", enumerate(text.splitlines(), 1))


def mode_pre_push(sc: Scanner) -> None:
    for line in sys.stdin.read().splitlines():
        parts = line.split()
        if len(parts) != 4:
            continue
        _lref, lsha, _rref, rsha = parts
        if lsha == ZERO_SHA:          # ブランチ削除
            continue
        if rsha == ZERO_SHA or not has_object(rsha):
            # 新規ブランチ、または強制プッシュ等でリモート側のコミットが手元に無い場合:
            # リモート追跡ブランチに無いコミットをすべて検査する
            scan_commits(sc, [lsha, "--not", "--remotes"])
        else:
            scan_commits(sc, [f"{rsha}..{lsha}"])


# ---------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(description="公開前の機密情報チェック")
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--staged", action="store_true")
    g.add_argument("--pre-push", action="store_true")
    g.add_argument("--all", action="store_true")
    g.add_argument("--history", action="store_true")
    g.add_argument("--path")
    ap.add_argument("--denylist")
    ap.add_argument("--show-values", action="store_true",
                    help="検出値を伏字にせず表示する（手元確認用。ログを共有しないこと）")
    args = ap.parse_args()

    deny_rules, deny_path = load_denylist(args.denylist)
    if not deny_path:
        print("[leak-guard] 注意: 禁止語リストが見つかりません。組み込みルールのみで検査します。",
              file=sys.stderr)

    if args.path:
        root = args.path
    else:
        root = git("rev-parse", "--show-toplevel").decode().strip()
        os.chdir(root)  # git のパスはリポジトリ直下からの相対パスのため
    sc = Scanner(BUILTIN_RULES + deny_rules, load_allowlist(root))

    try:
        if args.staged:
            mode_staged(sc)
        elif args.pre_push:
            mode_pre_push(sc)
        elif args.all:
            mode_all(sc)
        elif args.history:
            scan_commits(sc, ["--all"])
        else:
            mode_path(sc, args.path)
    except subprocess.CalledProcessError as e:
        print(f"[leak-guard] git の実行に失敗しました: {e.stderr.decode(errors='replace')}", file=sys.stderr)
        return 2

    if not sc.findings:
        print(f"[leak-guard] OK（禁止語 {len(deny_rules)} 件 + 組み込みルールで検査）")
        return 0

    print(f"[leak-guard] 公開できない可能性がある内容を {len(sc.findings)} 件検出しました。中止します。\n")
    seen = set()
    for where, line, rule, val in sc.findings:
        key = (where, line, rule, val)
        if key in seen:
            continue
        seen.add(key)
        try:
            print(f"  {where}:{line}  [{rule}]  {val if args.show_values else mask(val)}")
        except BrokenPipeError:  # head 等にパイプした場合
            break
    print("\n対処: 該当箇所を一般化して再実行してください。"
          f"\n      問題ない行には「{ALLOW_MARK}」とコメントを付けると除外できます。")
    return 1


if __name__ == "__main__":
    sys.exit(main())
