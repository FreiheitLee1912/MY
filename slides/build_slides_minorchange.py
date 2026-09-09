# -*- coding: utf-8 -*-
"""マイナーチェンジ管理ルールの設定について（ご決定依頼）— 3枚.

Usage:
    python build_slides_minorchange.py --template 2026_Standard_Template_WideScreenEN.pptx

  1枚目 ご決定依頼（決定事項2点／前提条件／次のアクション）
  2枚目 提案1（バーコード変更のみのPGM）
  3枚目 提案2（同時期立ち上げの複数PGM）

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


def banner(sl, text, y=0.94, h=0.42, size=11, bold=False):
    box(sl, L, y, W, h, text, size, bold, BLACK, CREAM, None, pad=0.20, space=1.25)


def caution(sl, y, h, label, body, size=10):
    box(sl, L, y, W, h, "", fill=YELLOW, border=BLACK)
    rich(sl, L + 0.20, y + 0.05, W - 0.40, h - 0.10,
         [(label + "　", 11, True, RED), (body, size, False, BLACK)],
         anchor=MSO_ANCHOR.MIDDLE, space=1.25)


def basis(sl, y, industry, internal, bar_h=0.30, body_h=0.90):
    """業界実務と自社の現状を左右に並べる（根拠の分離）。"""
    cw = (W - 0.20) / 2
    bar(sl, L, y, cw, bar_h, "判断の考え方（業界実務）", 11)
    box(sl, L, y + bar_h, cw, body_h, industry, 9, space=1.25, pad=0.14)
    x2 = L + cw + 0.20
    bar(sl, x2, y, cw, bar_h, "自社の現状", 11)
    box(sl, x2, y + bar_h, cw, body_h, internal, 9, space=1.25, pad=0.14)


# ====================================================== P1 ご決定依頼 ========
def slide_1(prs, n):
    sl = page(prs, "マイナーチェンジのJIRA管理要否について、2点のご決定をお願いしたい",
              None, (), n, title_size=20, date=DATE)
    banner(sl, "いずれも管理の網羅性と管理工数の抑制の両立を狙うもの。判断の考え方は"
               "自動車業界の変更管理実務に準拠している。［要確認：出典］")

    dy = 1.44
    bar(sl, L, dy, W, 0.34, "本日ご決定いただきたい事項（2点）", 12.5)
    cols = [("No.", 0.60), ("決定事項", 6.20), ("提案", W - 6.80)]
    tbl_head(sl, L, dy + 0.34, cols)
    decisions = [
        ("1", "バーコード変更のみのPGMを、JIRAで個別管理する対象から除外してよいか",
         "除外し、変更記録のみを残す運用とする（提案1）"),
        ("2", "同時期に立ち上がる複数PGMを、代表PGM1件に集約して管理してよいか",
         "集約する。ただし差分発生時は独立管理へ切り替える（提案2）"),
    ]
    ry = dy + 0.62
    for no, q, a in decisions:
        box(sl, L, ry, 0.60, 0.62, no, 11, True, NAVY_TXT, PALE, align=PP_ALIGN.CENTER)
        box(sl, L + 0.60, ry, 6.20, 0.62, q, 10.5, space=1.25)
        box(sl, L + 6.80, ry, W - 6.80, 0.62, a, 10.5, True, NAVY_TXT, PALE, space=1.25)
        ry += 0.62

    caution(sl, 3.46, 0.56, "決定の前提条件",
            "対象範囲（表示変更のみ）の定義について、品質保証部門との合意が前提となる。"
            "［要確認：合意は取得済みか。未取得の場合、本提案の決定は合意取得を条件とする］")

    ay = 4.20
    bar(sl, L, ay, W, 0.34, "次のアクション", 12.5)
    cols2 = [("No.", 0.60), ("実施事項", 5.20), ("成果物", 2.90), ("担当", 1.40), ("期限", 1.40)]
    tbl_head(sl, L, ay + 0.34, cols2)
    acts = [("1", "本提案の決定（提案1・提案2）", "会議議事録", "［要確認］", "本会議"),
            ("2", "未決事項（判定者・承認者・記録先・差分閾値）の確定",
             "判定基準表（確定版）", "［要確認］", "［要確認］"),
            ("3", "品質保証部門との合意取得", "合意記録", "［要確認］", "［要確認］"),
            ("4", "マイナーチェンジ管理ルールの文書化", "ルール文書", "［要確認］", "［要確認］"),
            ("5", "JIRA運用ガイドへの反映", "改訂版運用ガイド", "［要確認］", "［要確認］")]
    ry = ay + 0.62
    for no, item, out, who, when in acts:
        hot = when == "本会議"
        box(sl, L, ry, 0.60, 0.34, no, 9.5, True, NAVY_TXT, PALE, align=PP_ALIGN.CENTER)
        box(sl, L + 0.60, ry, 5.20, 0.34, item, 9.5, pad=0.12)
        box(sl, L + 5.80, ry, 2.90, 0.34, out, 9.5, pad=0.12)
        box(sl, L + 8.70, ry, 1.40, 0.34, who, 9, align=PP_ALIGN.CENTER, pad=0.04)
        box(sl, L + 10.10, ry, 1.40, 0.34, when, 9, hot, NAVY_TXT,
            PALE if hot else WHITE, align=PP_ALIGN.CENTER, pad=0.04)
        ry += 0.34


# ========================================================== P2 提案1 ========
def slide_2(prs, n):
    sl = page(prs, "提案1：工程・品質・顧客承認に影響しない変更は、JIRA個別管理から外す",
              None, (), n, title_size=20, date=DATE)
    banner(sl, "3項目すべてが「該当なし」の場合に限り、JIRAでの個別起票を不要とし、"
               "変更記録のみを残す。", bold=True)

    basis(sl, 1.44,
          "自動車業界では、量産開始後の変更を一律に扱わず、影響範囲に応じて管理レベルを"
          "階層化する。設計変更（形状・材料・機能に影響）と表示変更（機能に影響しない）を"
          "区別し、後者は変更履歴の記録で処理する。PPAP（生産部品承認プロセス）の再提出"
          "要否も、この区分を基準に判定される。［要確認：出典（社内標準／業界規格／"
          "他社事例のいずれか）］",
          "［要確認：直近のマイナーチェンジ案件のうち、表示変更のみの案件が何件あり、"
          "現在どのように処理されているか］")

    cy = 2.74
    bar(sl, L, cy, W, 0.34, "提案する判定基準", 12.5)
    cols = [("No.", 0.60), ("判定項目", 6.90), ("判定結果が「該当」の場合", W - 7.50)]
    tbl_head(sl, L, cy + 0.34, cols)
    checks = [("1", "工程・設備への影響（新規工程・設備変更を伴うか）", "JIRAで個別管理"),
              ("2", "品質・機能への影響（製品仕様・性能に影響しうるか）", "JIRAで個別管理"),
              ("3", "顧客への届出義務（顧客承認・PPAP再提出が必要か）", "JIRAで個別管理")]
    ry = cy + 0.62
    for no, item, res in checks:
        box(sl, L, ry, 0.60, 0.58, no, 11, True, NAVY_TXT, PALE, align=PP_ALIGN.CENTER)
        box(sl, L + 0.60, ry, 6.90, 0.58, item, 10.5, space=1.25)
        box(sl, L + 7.50, ry, W - 7.50, 0.58, res, 10.5, align=PP_ALIGN.CENTER)
        ry += 0.58

    ny = ry + 0.16
    bar(sl, L, ny, W, 0.32, "リスクと対策", 12)
    box(sl, L, ny + 0.32, W, 0.62,
        "変更自体を記録しないわけではなく、JIRAでの進捗管理を省略するにとどめる。"
        "ただし、対象範囲（表示変更のみ）の定義について品質保証部門との合意が前提となる。",
        10, False, BLACK, PALE, None, pad=0.20, space=1.3)

    say(sl, L, ny + 1.06, W, 0.26,
        "※ バーコード変更＝型式ラベル・バーコード・刻印等、製品の形状・材料・機能に影響しない"
        "識別情報の変更。［要確認：部品番号・製品仕様の変更を伴わないという理解でよいか］",
        9, False, MUTED)


# ========================================================== P3 提案2 ========
def slide_3(prs, n):
    sl = page(prs, "提案2：同時期立ち上げの複数PGMは、代表PGM1件に集約する",
              None, (), n, title_size=20, date=DATE)
    banner(sl, "代表PGMをJIRAの主チケットとし、派生PGMはサブタスクで管理する。"
               "差分が発生した時点で独立チケットへ切り替える。", bold=True)

    # --- 自社の現状（最上段） ---
    ty = 1.44
    cw = (W - 0.20) / 2
    bar(sl, L, ty, cw, 0.30, "自社の現状", 11)
    box(sl, L, ty + 0.30, cw, 0.62,
        "［要確認：同時期立ち上げの複数PGMが年間何件発生しているか、"
        "現在1件ずつ起票しているか］", 9, space=1.25, pad=0.14)
    x2 = L + cw + 0.20
    bar(sl, x2, ty, cw, 0.30, "判断の考え方（業界実務）", 11)
    box(sl, x2, ty + 0.30, cw, 0.62,
        "プラットフォーム共通化により、同一ベース設計から派生する複数型式が同時期に"
        "立ち上がることは常態化している。代表機種を主管理とし、派生機種は差分のみを"
        "管理項目として紐付ける方式が採られる。［要確認：出典］", 9, space=1.25, pad=0.14)

    # --- 提案する管理方式（中段） ---
    my = 2.50
    bar(sl, L, my, W, 0.32, "提案する管理方式", 12.5)
    cols = [("対象", 3.20), ("管理方式", 4.10), ("管理項目", W - 7.30)]
    tbl_head(sl, L, my + 0.32, cols, h=0.26)
    rows = [("代表PGM", "JIRA主チケットで個別管理", "通常のプロジェクト管理項目", True),
            ("派生PGM（差分なし）", "代表チケットのサブタスクとして紐付け",
             "型式・差分内容・担当・納期の4項目", False),
            ("派生PGM（差分発生後）", "独立チケットへ切り替え",
             "通常のプロジェクト管理項目", False)]
    ry = my + 0.58
    for tgt, how, items, hot in rows:
        box(sl, L, ry, 3.20, 0.50, tgt, 10.5, True, NAVY_TXT, PALE)
        box(sl, L + 3.20, ry, 4.10, 0.50, how, 10, hot, BLACK, WHITE, space=1.2)
        box(sl, L + 7.30, ry, W - 7.30, 0.50, items, 10, space=1.2)
        ry += 0.50

    # --- リスクと対策（下段） ---
    cy = ry + 0.16
    bar(sl, L, cy, W, 0.32, "リスクと対策", 12.5)
    box(sl, L, cy + 0.32, W, 0.32,
        "派生PGM固有の遅延・仕様差異が代表管理の陰に隠れ、検知が遅れること。",
        10, True, RED, PALE, None, pad=0.20)
    measures = [("1", "派生PGMのサブタスクには、代表チケットとは独立した進捗ステータスを"
                      "持たせる。"),
                ("2", "差分（工程・品質・納期・顧客承認のいずれか）が発生した時点で独立"
                      "チケットへ切り替える。［要確認：切り替えの閾値（例：納期乖離◯日以上）］"),
                ("3", "［要確認：切り替え判断を誰が行い、誰が確認するか］")]
    mw = (W - 0.40) / 3
    mgy = cy + 0.68
    for i, (no, m) in enumerate(measures):
        x = L + i * (mw + 0.20)
        box(sl, x, mgy, 0.34, 0.60, no, 10, True, WHITE, NAVY, align=PP_ALIGN.CENTER)
        box(sl, x + 0.34, mgy, mw - 0.34, 0.60, m, 8, space=1.25, pad=0.10)

    caution(sl, mgy + 0.70, 0.42, "適用範囲",
            "［要確認：適用開始時期、および既に起票済みの案件への遡及適用の要否］", size=9.5)
    say(sl, L, mgy + 1.18, W, 0.24,
        "※ 代表PGM＝主管理対象とする1件。派生PGM＝仕向地違い・ハンドル位置違い等の派生型式。"
        "差分＝工程・品質・納期・顧客承認のいずれかに関し代表PGMと相違が生じた状態。",
        8.5, False, MUTED)


def build(template, output):
    prs = Presentation(base_from_template(template))
    slide_1(prs, 1)
    slide_2(prs, 2)
    slide_3(prs, 3)
    prs.save(output)
    print("saved", output, "-", len(prs.slides._sldIdLst), "slides")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("-o", "--output", default="minorchange_rule_slides.pptx")
    a = ap.parse_args()
    build(a.template, a.output)
