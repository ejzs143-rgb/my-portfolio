#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
comment_grouper.py  v4
VBAから呼び出す日本語コメント類似グルーピングスクリプト

使い方: python comment_grouper.py <入力CSV> <出力CSV> [類似度閾値]
  閾値デフォルト: 0.35
"""

import csv
import sys
import os
import re
from difflib import SequenceMatcher

DEFAULT_THRESHOLD = 0.35
MAX_GROUP_NAME_LEN = 40
MIN_COMMENT_LEN = 3

W_DIFFLIB = 0.25
W_BIGRAM  = 0.25
W_KEYWORD = 0.50

SYNONYM_MAP = {
    '写真': ['画像', 'フォト', '写真', '動画'],
    '添付なし': ['未添付', '添付漏れ', '添付なし', '添付がない', '添付されていない',
                 'アップされていない', 'アップロードがされていない', '確認できません'],
    '交換': ['交換', '取替'],
    '超過': ['超過', '超えて', '過大'],
    '不良': ['不良', '不具合', '故障', '破損'],
    '違い': ['違い', '異なる', '不一致', '一致しません'],
}


def detect_encoding(filepath):
    """ファイルのエンコーディングを自動検出"""
    for enc in ['utf-8-sig', 'utf-8', 'cp932', 'shift_jis']:
        try:
            with open(filepath, 'r', encoding=enc) as f:
                f.read()
            return enc
        except (UnicodeDecodeError, UnicodeError):
            continue
    return 'utf-8'


def normalize_comment(text):
    if not text:
        return ""
    result = ""
    for ch in text:
        code = ord(ch)
        if 0xFF01 <= code <= 0xFF5E:
            result += chr(code - 0xFEE0)
        else:
            result += ch
    result = result.lower().strip()
    result = re.sub(r'\s+', ' ', result)
    return result


def get_bigrams(text):
    if len(text) < 2:
        return set()
    return {text[i:i+2] for i in range(len(text) - 1)}


def normalize_keyword_synonyms(keywords):
    normalized = set()
    for kw in keywords:
        matched = False
        for canon, synonyms in SYNONYM_MAP.items():
            for syn in synonyms:
                if syn in kw or kw in syn:
                    normalized.add(canon)
                    matched = True
                    break
            if matched:
                break
        if not matched:
            normalized.add(kw)
    return normalized


def extract_keywords(text):
    keywords = set()
    for m in re.finditer(r'[A-Za-z0-9]{2,}', text):
        keywords.add(m.group().upper())
    for m in re.finditer(r'[（(](.+?)[）)]', text):
        keywords.add(m.group(1))
    jp_keywords = [
        '工賃', '車種', '写真', '画像', '動画', '部品', '交換',
        '添付', '未添付', '漏れ', '不備', '超過', '基準',
        'オイル', 'ブレーキ', 'エンジン', 'ポンプ',
        '請求', '申請', '記載', '不一致', '対象外',
        '不良', '不具合', '摩耗', '破損', '故障',
        'アップロード', 'アップ',
    ]
    for kw in jp_keywords:
        if kw in text:
            keywords.add(kw)
    keywords = normalize_keyword_synonyms(keywords)
    return keywords


def jaccard_similarity(set1, set2):
    if not set1 or not set2:
        return 0.0
    intersection = len(set1 & set2)
    union = len(set1 | set2)
    return intersection / union if union > 0 else 0.0


def keyword_similarity(kw1, kw2):
    if not kw1 or not kw2:
        return 0.0
    base_sim = jaccard_similarity(kw1, kw2)
    alpha_kw1 = {k for k in kw1 if re.match(r'^[A-Z0-9]+$', k)}
    alpha_kw2 = {k for k in kw2 if re.match(r'^[A-Z0-9]+$', k)}
    partial_matches = 0
    total_alpha = max(len(alpha_kw1), len(alpha_kw2), 1)
    for a in alpha_kw1:
        for b in alpha_kw2:
            if a != b and (a in b or b in a):
                partial_matches += 1
    partial_bonus = 0.15 * partial_matches / total_alpha
    return min(base_sim + partial_bonus, 1.0)


def calc_similarity(norm1, norm2, bigrams1, bigrams2, kw1, kw2):
    str_sim = SequenceMatcher(None, norm1, norm2).ratio()
    bigram_sim = jaccard_similarity(bigrams1, bigrams2)
    kw_sim = keyword_similarity(kw1, kw2)
    total = W_DIFFLIB * str_sim + W_BIGRAM * bigram_sim + W_KEYWORD * kw_sim
    return min(total, 1.0)


def make_group_name(representative_comment):
    name = representative_comment.strip()
    # 改行をスペースに置換（VBAのLine Input対策）
    name = name.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
    if len(name) > MAX_GROUP_NAME_LEN:
        name = name[:MAX_GROUP_NAME_LEN] + "…"
    return name


def group_comments(rows, threshold):
    groups = {}
    row_group_map = {}

    for i, row in enumerate(rows):
        comment = row.get("追加コメント", "")

        if len(comment.strip()) < MIN_COMMENT_LEN:
            grp_name = "（コメントなし/短文）"
            if grp_name not in groups:
                groups[grp_name] = {
                    "representative": "", "norm": "",
                    "bigrams": set(), "orig_kw": set(),
                    "all_kw": set(), "members": []
                }
            groups[grp_name]["members"].append(i)
            row_group_map[i] = grp_name
            continue

        norm = normalize_comment(comment)
        bigrams = get_bigrams(norm)
        kw = extract_keywords(comment)

        best_group = None
        best_sim = 0.0

        for grp_name, grp_info in groups.items():
            if grp_name == "（コメントなし/短文）":
                continue
            sim = calc_similarity(
                norm, grp_info["norm"],
                bigrams, grp_info["bigrams"],
                kw, grp_info["orig_kw"]
            )
            if sim >= threshold and sim > best_sim:
                best_sim = sim
                best_group = grp_name

        if best_group:
            groups[best_group]["members"].append(i)
            groups[best_group]["all_kw"] |= kw
            row_group_map[i] = best_group
        else:
            grp_name = make_group_name(comment)
            if grp_name in groups:
                grp_name = grp_name + f" ({i})"
            groups[grp_name] = {
                "representative": comment, "norm": norm,
                "bigrams": bigrams, "orig_kw": kw.copy(),
                "all_kw": kw.copy(), "members": [i]
            }
            row_group_map[i] = grp_name

    # パス2: 小グループ同士のマージ
    merge_threshold = threshold * 0.85
    merged = True
    while merged:
        merged = False
        grp_names = [k for k in groups.keys() if k != "（コメントなし/短文）"]
        for a in range(len(grp_names)):
            if grp_names[a] not in groups:
                continue
            for b in range(a + 1, len(grp_names)):
                if grp_names[b] not in groups:
                    continue
                ga = groups[grp_names[a]]
                gb = groups[grp_names[b]]
                sim = calc_similarity(
                    ga["norm"], gb["norm"],
                    ga["bigrams"], gb["bigrams"],
                    ga["all_kw"], gb["all_kw"]
                )
                if sim >= merge_threshold:
                    if len(ga["members"]) >= len(gb["members"]):
                        keep, drop = grp_names[a], grp_names[b]
                    else:
                        keep, drop = grp_names[b], grp_names[a]
                    groups[keep]["members"].extend(groups[drop]["members"])
                    groups[keep]["all_kw"] |= groups[drop]["all_kw"]
                    for idx in groups[drop]["members"]:
                        row_group_map[idx] = keep
                    del groups[drop]
                    merged = True
                    break
            if merged:
                break

    return groups, row_group_map


def main():
    if len(sys.argv) < 3:
        print("使い方: python comment_grouper.py <入力CSV> <出力CSV> [類似度閾値]")
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]
    threshold = float(sys.argv[3]) if len(sys.argv) >= 4 else DEFAULT_THRESHOLD

    if not os.path.exists(input_path):
        print(f"エラー: 入力ファイルが見つかりません: {input_path}")
        sys.exit(1)

    # エンコーディング自動検出
    enc = detect_encoding(input_path)
    print(f"検出エンコーディング: {enc}")

    rows = []
    with open(input_path, 'r', encoding=enc) as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)

    if not rows:
        print("エラー: データが0件です。")
        sys.exit(1)

    print(f"読み込み: {len(rows)}件")
    print(f"類似度閾値: {threshold}")

    groups, row_group_map = group_comments(rows, threshold)

    print(f"\n{'='*60}")
    print(f"グルーピング結果: {len(groups)}グループ")
    print(f"{'='*60}")

    sorted_groups = sorted(groups.items(), key=lambda x: len(x[1]["members"]), reverse=True)

    for grp_name, grp_info in sorted_groups:
        count = len(grp_info["members"])
        pct = count / len(rows) * 100
        print(f"  [{count:3d}件 {pct:5.1f}%] {grp_name}")

    # 出力はcp932（VBAのLine Inputと互換性を保つ）
    out_enc = 'cp932'
    fieldnames = list(rows[0].keys()) + ["グループ名"]
    with open(output_path, 'w', encoding=out_enc, newline='', errors='replace') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for i, row in enumerate(rows):
            # 全フィールドの改行をスペースに置換（VBA Line Input対策）
            sanitized = {}
            for k, v in row.items():
                sanitized[k] = str(v).replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
            sanitized["グループ名"] = row_group_map.get(i, "未分類")
            writer.writerow(sanitized)
    print(f"\n出力完了: {output_path}")

    summary_path = output_path.replace(".csv", "_summary.csv")
    with open(summary_path, 'w', encoding=out_enc, newline='', errors='replace') as f:
        writer = csv.writer(f)
        writer.writerow(["グループ名", "件数", "構成比"])
        for grp_name, grp_info in sorted_groups:
            count = len(grp_info["members"])
            pct = f"{count / len(rows) * 100:.1f}%"
            clean_name = grp_name.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
            writer.writerow([clean_name, count, pct])
    print(f"集計出力: {summary_path}")


if __name__ == "__main__":
    main()
