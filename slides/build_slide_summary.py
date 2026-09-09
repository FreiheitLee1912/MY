# -*- coding: utf-8 -*-
"""P.2 本日のサマリー（新設）.

Usage:
    python build_slide_summary.py --template 2026_Standard_Template_WideScreenEN.pptx

版面・配色は build_slides_mgmt.py と共通（濃紺パネル＋クリーム帯、Meiryo）。
"""
import argparse

from pptx import Presentation
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from titus_kit import BLACK, WHITE, say, rich, cell
from daicel_kit import L, R, W, MUTED, RED, base_from_template, page
from build_slides_mgmt import NAVY, NAVY_TXT, CREAM, PALE, THEAD, RULE, YELLOW, box, bar

DATE = "2026/09/09"


def tbl_head(sl, x, y, cols, h=0.28, size=10):
    for t, w in cols:
        box(sl, x, y, w, h, t, size, True, NAVY_TXT, THEAD, align=PP_ALIGN.CENTER)
        x += w


def summary(prs, n=2):
    sl = page(prs, "本日のサマリー", None, (), n, title_size=22, date=DATE)

    # --- メッセージライン ---
    box(sl, L, 0.94, W, 0.48,
        "◇JIRA登録対象を「新規立上げ」「PCR」の2区分に限定し、登録タイミング・"
        "管理内容・役割分担について本日ご決定いただきたい",
        12.5, True, BLACK, CREAM, None, pad=0.20, space=1.25)

    # --- 本日ご決定いただきたい事項 ---
    dy = 1.52
    bar(sl, L, dy, W, 0.34, "本日ご決定いただきたい事項", 12.5)
    cols = [("No.", 0.70), ("決定事項", 3.60), ("提案内容", W - 4.30)]
    tbl_head(sl, L, dy + 0.34, cols)
    rows = [
        ("1", "JIRA登録対象の範囲",
         "生準が管理する4分類のうち、「新規立上げ」「PCR」の2区分に限定する"),
        ("2", "JIRA登録のタイミング",
         "［要確認：確定したトリガー］の時点で登録する（P.6）"),
        ("3", "管理内容・役割分担・更新頻度",
         "進捗管理と課題管理の2種を、拠点主体・月次更新で運用する（P.10・P.11）"),
    ]
    ry = dy + 0.62
    for no, what, how in rows:
        box(sl, L, ry, 0.70, 0.58, no, 11, True, NAVY_TXT, PALE, align=PP_ALIGN.CENTER)
        box(sl, L + 0.70, ry, 3.60, 0.58, what, 10.5, True, NAVY_TXT, PALE, space=1.25)
        box(sl, L + 4.30, ry, W - 4.30, 0.58, how, 10.5, space=1.25)
        ry += 0.58

    # --- 2区分に限定する理由 ---
    wy = 4.06
    bar(sl, L, wy, W, 0.32, "2区分に限定する理由", 12)
    reasons = [
        "「新規立上げ」「PCR」は、いずれも生準が主体的に工数を投下し、"
        "進捗と課題を関係者間で共有する必要がある案件である",
        "一方、ポテンシャル案件は受注確度が低く前提が変動するため、登録しても情報が"
        "更新されない。補給品は別プロジェクトで既に管理されている",
        "全案件を登録すると更新負荷が上がり、かえって情報が信用されなくなる",
    ]
    rw = (W - 0.40) / 3
    for i, txt in enumerate(reasons):
        x = L + i * (rw + 0.20)
        box(sl, x, wy + 0.36, 0.34, 0.80, str(i + 1), 10, True, WHITE, NAVY,
            align=PP_ALIGN.CENTER)
        box(sl, x + 0.34, wy + 0.36, rw - 0.34, 0.80, txt, 9, space=1.3, pad=0.12)

    # --- 本日ご判断が必要な未決事項 ---
    oy = 5.42
    bar(sl, L, oy, W, 0.32, "本日ご判断が必要な未決事項（詳細はP.15）", 12, fill=RED)
    opens = ["JIRAを直接利用できない拠点が存在する場合の情報連携方式",
             "主要マイルストーンの最終決定者",
             "例外管理①の適用について合意を要する部門"]
    ow = (W - 0.40) / 3
    for i, txt in enumerate(opens):
        x = L + i * (ow + 0.20)
        box(sl, x, oy + 0.36, 0.34, 0.66, str(i + 1), 10, True, BLACK, YELLOW,
            border=BLACK, align=PP_ALIGN.CENTER)
        box(sl, x + 0.34, oy + 0.36, ow - 0.34, 0.66, txt, 9.5, False, BLACK,
            YELLOW, BLACK, space=1.3, pad=0.12)


def build(template, output):
    prs = Presentation(base_from_template(template))
    summary(prs, 2)
    prs.save(output)
    print("saved", output, "-", len(prs.slides._sldIdLst), "slides")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("-o", "--output", default="summary_slide.pptx")
    a = ap.parse_args()
    build(a.template, a.output)
