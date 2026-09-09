# -*- coding: utf-8 -*-
"""マイナーチェンジ管理ルールの設定について（ご決定依頼）— 2枚.

Usage:
    python build_slides_minorchange.py --template 2026_Standard_Template_WideScreenEN.pptx

レビュー修正版の内容をそのまま載せる。原文にない事実は補わず、
未確定箇所は［要確認：…］としてスライド上に残す。
"""
import argparse

from pptx import Presentation
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from titus_kit import C, BLACK, WHITE, style, say, rich, cell
from daicel_kit import L, R, W, MUTED, RED, base_from_template, page
from build_slides_mgmt import NAVY, NAVY_TXT, CREAM, PALE, THEAD, RULE, YELLOW, box, bar

DATE = "2026/09/09"


def tbl_head(sl, x, y, cols, h=0.28, size=10):
    for t, w in cols:
        box(sl, x, y, w, h, t, size, True, NAVY_TXT, THEAD, align=PP_ALIGN.CENTER)
        x += w


# ============================================================== P1 ===========
def slide_1(prs, n):
    sl = page(prs, "マイナーチェンジのJIRA管理要否について、2点のご決定をお願いしたい",
              None, (), n, title_size=20, date=DATE)

    box(sl, L, 0.94, W, 0.44,
        "いずれも管理の網羅性と管理工数の抑制の両立を狙うもの。判断の考え方は自動車業界の"
        "変更管理実務に準拠している。［要確認：出典］",
        11, False, BLACK, CREAM, None, pad=0.20)

    # --- 決定依頼 ---
    dy = 1.44
    bar(sl, L, dy, W, 0.32, "本日ご決定いただきたい事項（2点）", 12.5)
    cols = [("No.", 0.60), ("決定事項", 6.40), ("提案", W - 7.00)]
    tbl_head(sl, L, dy + 0.32, cols, h=0.26)
    decisions = [
        ("1", "バーコード変更のみのPGMを、JIRAで個別管理する対象から除外してよいか",
         "除外し、変更記録のみを残す運用とする"),
        ("2", "同時期に立ち上がる複数PGMを、代表PGM1件に集約して管理してよいか",
         "集約する。ただし差分発生時は独立管理へ切り替える"),
    ]
    ry = dy + 0.58
    for no, q, a in decisions:
        box(sl, L, ry, 0.60, 0.46, no, 10.5, True, NAVY_TXT, PALE, align=PP_ALIGN.CENTER)
        box(sl, L + 0.60, ry, 6.40, 0.46, q, 10, space=1.2)
        box(sl, L + 7.00, ry, W - 7.00, 0.46, a, 10, True, NAVY_TXT, PALE, space=1.2)
        ry += 0.46

    # --- 左：提案1 判定基準 ---
    cy = 3.06
    lw = 6.60
    bar(sl, L, cy, lw, 0.32, "提案1｜バーコード変更のみのPGM ― 判定基準", 12)
    cols = [("No.", 0.55), ("判定項目", 4.05), ("該当する場合", 2.00)]
    tbl_head(sl, L, cy + 0.32, cols, h=0.26, size=9.5)
    checks = [("1", "工程・設備への影響（新規工程・設備変更を伴うか）", "JIRAで個別管理"),
              ("2", "品質・機能への影響（製品仕様・性能に影響しうるか）", "JIRAで個別管理"),
              ("3", "顧客への届出義務（顧客承認・PPAP再提出が必要か）", "JIRAで個別管理")]
    ry = cy + 0.58
    for no, item, res in checks:
        box(sl, L, ry, 0.55, 0.42, no, 10, True, NAVY_TXT, PALE, align=PP_ALIGN.CENTER)
        box(sl, L + 0.55, ry, 4.05, 0.42, item, 9.5, space=1.2)
        box(sl, L + 4.60, ry, 2.00, 0.42, res, 9.5, align=PP_ALIGN.CENTER)
        ry += 0.42
    box(sl, L, ry + 0.08, lw, 0.50,
        "3項目すべてが「該当なし」の場合に限り、JIRAでの個別起票を不要とし、"
        "変更記録のみを残す。",
        10.5, True, NAVY_TXT, PALE, None, align=PP_ALIGN.CENTER, space=1.25)

    # --- 右：運用上の未決事項 ---
    x2 = L + lw + 0.20
    rw = W - lw - 0.20
    bar(sl, x2, cy, rw, 0.32, "運用上の未決事項｜本会議で方向性をご指示願いたい", 11)
    opens = [("判定者", "PGM担当者の自己判定とするか、第三者確認を要するか［要確認］"),
             ("承認者", "［要確認］"),
             ("判定タイミング", "変更情報の受領時か、着手判断時か［要確認］"),
             ("記録先", "既存の変更履歴台帳が存在するか。無い場合の代替手段［要確認］"),
             ("記録項目", "［要確認］")]
    ry2 = cy + 0.32
    for k, v in opens:
        box(sl, x2, ry2, 1.30, 0.42, k, 9.5, True, NAVY_TXT, PALE, pad=0.08)
        box(sl, x2 + 1.30, ry2, rw - 1.30, 0.42, v, 8.5, space=1.2, pad=0.10)
        ry2 += 0.42

    # --- 前提条件 ---
    ny = 5.66
    box(sl, L, ny, W, 0.56, "", fill=YELLOW, border=BLACK)
    rich(sl, L + 0.20, ny + 0.06, W - 0.40, 0.46,
         [("決定の前提条件　", 11, True, RED),
          ("対象範囲（表示変更のみ）の定義について、品質保証部門との合意が前提となる。"
           "［要確認：合意は取得済みか。未取得の場合、本提案の決定は合意取得を条件とする］",
           10, False, BLACK)], anchor=MSO_ANCHOR.MIDDLE, space=1.25)

    say(sl, L, 6.34, W, 0.24,
        "※ バーコード変更＝型式ラベル・刻印等、製品の形状・材料・機能に影響しない識別情報の変更。"
        "［要確認：部品番号・製品仕様の変更を伴わないという理解でよいか］",
        9, False, MUTED)


# ============================================================== P2 ===========
def slide_2(prs, n):
    sl = page(prs, "提案2：同時期立ち上げの複数PGMは、代表PGM1件に集約する",
              None, (), n, title_size=20, date=DATE)

    box(sl, L, 0.94, W, 0.44,
        "代表PGMをJIRAの主チケットとし、派生PGMはサブタスクで管理する。"
        "差分が発生した時点で独立チケットへ切り替える。",
        11, True, BLACK, CREAM, None, pad=0.20)

    # --- 管理方式 ---
    my = 1.44
    bar(sl, L, my, W, 0.32, "提案する管理方式", 12.5)
    cols = [("対象", 3.20), ("管理方式", 4.00), ("管理項目", W - 7.20)]
    tbl_head(sl, L, my + 0.32, cols, h=0.26)
    rows = [("代表PGM", "JIRA主チケットで個別管理", "通常のプロジェクト管理項目", True),
            ("派生PGM（差分なし）", "代表チケットのサブタスクとして紐付け",
             "型式・差分内容・担当・納期の4項目", False),
            ("派生PGM（差分発生後）", "独立チケットへ切り替え", "通常のプロジェクト管理項目", False)]
    ry = my + 0.58
    for tgt, how, items, hot in rows:
        box(sl, L, ry, 3.20, 0.48, tgt, 10.5, True, NAVY_TXT, PALE)
        box(sl, L + 3.20, ry, 4.00, 0.48, how, 10, hot, BLACK, WHITE, space=1.2)
        box(sl, L + 7.20, ry, W - 7.20, 0.48, items, 10, space=1.2)
        ry += 0.48

    # --- 左：リスクと対策 ---
    cy = 3.60
    cw = (W - 0.20) / 2
    bar(sl, L, cy, cw, 0.32, "リスクと対策", 12)
    box(sl, L, cy + 0.32, cw, 0.44,
        "派生PGM固有の遅延・仕様差異が代表管理の陰に隠れ、検知が遅れること。",
        10, True, RED, PALE, None, space=1.25)
    measures = [("1", "派生PGMのサブタスクには、代表チケットとは独立した進捗ステータスを持たせる。"),
                ("2", "差分（工程・品質・納期・顧客承認のいずれか）が発生した時点で独立チケットへ"
                      "切り替える。［要確認：切り替えの閾値（例：納期乖離◯日以上）］"),
                ("3", "［要確認：切り替え判断を誰が行い、誰が確認するか］")]
    ry2 = cy + 0.84
    for no, m in measures:
        box(sl, L, ry2, 0.42, 0.50, no, 10, True, WHITE, NAVY, align=PP_ALIGN.CENTER)
        box(sl, L + 0.42, ry2, cw - 0.42, 0.50, m, 9, space=1.2, pad=0.10)
        ry2 += 0.50

    # --- 右：次のアクション ---
    x2 = L + cw + 0.20
    bar(sl, x2, cy, cw, 0.32, "次のアクション", 12)
    cols2 = [("#", 0.42), ("実施事項", 3.23), ("担当", 1.00), ("期限", 1.00)]
    tbl_head(sl, x2, cy + 0.32, cols2, h=0.26, size=9)
    acts = [("1", "本提案の決定（提案1・提案2）", "［要確認］", "本会議"),
            ("2", "未決事項（判定者・承認者・記録先・差分閾値）の確定", "［要確認］", "［要確認］"),
            ("3", "品質保証部門との合意取得", "［要確認］", "［要確認］"),
            ("4", "マイナーチェンジ管理ルールの文書化", "［要確認］", "［要確認］"),
            ("5", "JIRA運用ガイドへの反映", "［要確認］", "［要確認］")]
    ry3 = cy + 0.58
    for no, item, who, when in acts:
        box(sl, x2, ry3, 0.42, 0.32, no, 9, True, NAVY_TXT, PALE, align=PP_ALIGN.CENTER)
        box(sl, x2 + 0.42, ry3, 3.23, 0.32, item, 8.5, pad=0.10)
        box(sl, x2 + 3.65, ry3, 1.00, 0.32, who, 8.5, align=PP_ALIGN.CENTER, pad=0.04)
        box(sl, x2 + 4.65, ry3, 1.00, 0.32, when, 8.5, when == "本会議", NAVY_TXT,
            PALE if when == "本会議" else WHITE, align=PP_ALIGN.CENTER, pad=0.04)
        ry3 += 0.32

    # --- 適用範囲 ---
    ny = 6.04
    box(sl, L, ny, W, 0.52, "", fill=YELLOW, border=BLACK)
    rich(sl, L + 0.20, ny + 0.06, W - 0.40, 0.40,
         [("適用範囲　", 11, True, RED),
          ("［要確認：適用開始時期、および既に起票済みの案件への遡及適用の要否］",
           10, False, BLACK)], anchor=MSO_ANCHOR.MIDDLE, space=1.25)


def build(template, output):
    prs = Presentation(base_from_template(template))
    slide_1(prs, 1)
    slide_2(prs, 2)
    prs.save(output)
    print("saved", output, "-", len(prs.slides._sldIdLst), "slides")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("-o", "--output", default="minorchange_rule_slides.pptx")
    a = ap.parse_args()
    build(a.template, a.output)
