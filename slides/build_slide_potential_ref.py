# -*- coding: utf-8 -*-
"""P.5 ［参考］JIRA対象外案件の取扱い方針（ポテンシャル案件）.

Usage:
    python build_slide_potential_ref.py --template 2026_Standard_Template_WideScreenEN.pptx

版面・配色は build_slides_mgmt.py と共通（濃紺パネル＋クリーム帯、Meiryo）。
"""
import argparse

from pptx import Presentation
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from titus_kit import BLACK, WHITE, say, rich, cell
from daicel_kit import L, R, W, MUTED, RED, base_from_template, page
from build_slides_mgmt import NAVY, NAVY_TXT, CREAM, PALE, THEAD, RULE, YELLOW, box, bar

DATE = "2026/09/09"
BLUE = NAVY          # 強調行のヘッダ色
HL = PALE            # C+行の地色


def tbl_head(sl, x, y, cols, h=0.26, size=9.5, pad=0.12):
    for t, w in cols:
        box(sl, x, y, w, h, t, size, True, NAVY_TXT, THEAD,
            align=PP_ALIGN.CENTER, pad=pad)
        x += w


def reference(prs, n=5):
    sl = page(prs, "［参考］JIRA対象外案件の取扱い方針（ポテンシャル案件）",
              None, (), n, title_size=22, date=DATE)

    # --- メッセージライン ---
    box(sl, L, 0.94, W, 0.42,
        "◇ポテンシャル案件はJIRA登録対象外とするが、C+案件のみkintoneでモニタリングし、"
        "生産準備への影響を早期に把握する",
        12.5, True, BLACK, CREAM, None, pad=0.20)

    # --- 位置づけ ---
    box(sl, L, 1.44, W, 0.48, "", fill=PALE, border=RULE)
    rich(sl, L + 0.20, 1.48, W - 0.40, 0.40,
         [("位置づけ　", 10, True, NAVY_TXT),
          ("本ページはJIRA登録対象外の案件に関する運用であり、P.2の決定事項とは別の"
           "管理体系である。目的は受注前の全案件を詳細管理することではなく、"
           "将来生産準備案件へ移行する可能性がある案件を早期に把握することにある。",
           9.5, False, BLACK)], anchor=MSO_ANCHOR.MIDDLE, space=1.28)

    # --- Potential Rank と生準の関与 ---
    ry0 = 2.00
    bar(sl, L, ry0, W, 0.30, "Potential Rank と生準の関与", 12)
    cols = [("Rank", 0.80), ("定義", 4.60), ("生準の関与", 3.10), ("管理先", W - 8.50)]
    tbl_head(sl, L, ry0 + 0.30, cols)
    ranks = [
        ("A", "現行商権／受注獲得済のプログラム", "本表の対象外（新規立上げへ移行済）",
         "JIRA", False, False),
        ("B", "現行商権の次期車／商権獲得が確定的なもの", "関与あり",
         "既存のAction Plan／MTP・Budget管理", False, False),
        ("C+", "未受注だが戦略的に獲得を前提として取り組んでおり、投資判断にも"
               "考慮されるターゲット案件", "関与あり（本ルールの対象）", "kintone", True, False),
        ("C", "営業としてのターゲットプログラムだが、Moduleとは合意前のもの", "関与なし",
         "［要確認：「一部情報確認のみ」の具体内容］", False, True),
        ("D", "受注ターゲットにできるか、案件として先ずはリストに載せたもの", "関与なし",
         "対象外", False, True),
        ("X", "現行商権の次期車でも商権を失ったもの", "関与なし", "対象外", False, True),
    ]
    ry = ry0 + 0.56
    for rank, dfn, role, dest, hot, off in ranks:
        fill = HL if hot else WHITE
        fg = MUTED if off else BLACK
        box(sl, L, ry, 0.80, 0.34, rank, 11, True, WHITE if hot else NAVY_TXT,
            BLUE if hot else PALE, align=PP_ALIGN.CENTER)
        box(sl, L + 0.80, ry, 4.60, 0.34, dfn, 8.5, False, fg, fill, space=1.18, pad=0.10)
        box(sl, L + 5.40, ry, 3.10, 0.34, role, 8.5, hot, NAVY_TXT if hot else fg,
            fill, space=1.18, pad=0.10)
        box(sl, L + 8.50, ry, W - 8.50, 0.34, dest, 8.5, hot, NAVY_TXT if hot else fg,
            fill, space=1.18, pad=0.10)
        ry += 0.34

    # --- 下段左：C+案件の管理（3つの判定Gate） ---
    cy = ry + 0.14
    gw = 7.20
    bar(sl, L, cy, gw, 0.28, "C+案件の管理（3つの判定Gate）", 11)
    gcols = [("Gate", 0.70), ("判定内容", 1.70), ("主な判定条件", 3.10), ("判定後の処理", 1.70)]
    tbl_head(sl, L, cy + 0.28, gcols, h=0.24, size=9, pad=0.06)
    gates = [
        ("01", "C+該当／管理開始",
         "Action PlanでC+と判定／想定SOP・顧客判断時期／［要確認：全体管理対象の確定］",
         "kintoneへ登録"),
        ("02", "生準影響度判定",
         "設備・金型／工程／能力／拠点／長納期品／日程・技術・品質／投資／"
         "想定SOP時期／新規品種か否か", "小＝簡易／中＝標準／大＝重点"),
        ("03", "正式案件化", "Nomination／採用内示／正式発注など、受注・採用の確定",
         "JIRAへ新規立上げ案件を登録"),
    ]
    gy = cy + 0.52
    for no, what, cond, res in gates:
        box(sl, L, gy, 0.70, 0.44, no, 10, True, WHITE, NAVY, align=PP_ALIGN.CENTER)
        box(sl, L + 0.70, gy, 1.70, 0.44, what, 8.5, True, NAVY_TXT, PALE,
            space=1.18, pad=0.08)
        box(sl, L + 2.40, gy, 3.10, 0.44, cond, 8, space=1.18, pad=0.10)
        box(sl, L + 5.50, gy, 1.70, 0.44, res, 8, True, NAVY_TXT, PALE,
            align=PP_ALIGN.CENTER, space=1.18, pad=0.06)
        gy += 0.44

    # --- 下段右：影響度別の管理方法 ---
    x2 = L + gw + 0.20
    iw = W - gw - 0.20
    bar(sl, x2, cy, iw, 0.28, "影響度別の管理方法", 11)
    icols = [("影響度", 0.75), ("管理方法", 1.35), ("更新頻度", 0.80), ("確認内容", iw - 2.90)]
    tbl_head(sl, x2, cy + 0.28, icols, h=0.24, size=9, pad=0.05)
    levels = [("小", "個別担当で管理（レビューなし）", "発生都度", "変化点のみ", False),
              ("中", "個別担当で管理・報告", "発生都度", "影響・リスク・アクション", False),
              ("大", "レビューを実施", "月次", "新規性・長納期品・日程", True)]
    iy = cy + 0.52
    for lvl, how, freq, what, hot in levels:
        box(sl, x2, iy, 0.75, 0.44, lvl, 11, True, WHITE if hot else NAVY_TXT,
            BLUE if hot else PALE, align=PP_ALIGN.CENTER)
        box(sl, x2 + 0.75, iy, 1.35, 0.44, how, 8, hot, BLACK, HL if hot else WHITE,
            space=1.18, pad=0.08)
        box(sl, x2 + 2.10, iy, 0.80, 0.44, freq, 8, hot, BLACK, HL if hot else WHITE,
            align=PP_ALIGN.CENTER, pad=0.04)
        box(sl, x2 + 2.90, iy, iw - 2.90, 0.44, what, 8, hot, BLACK,
            HL if hot else WHITE, space=1.18, pad=0.08)
        iy += 0.44


def build(template, output):
    prs = Presentation(base_from_template(template))
    reference(prs, 5)
    prs.save(output)
    print("saved", output, "-", len(prs.slides._sldIdLst), "slides")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("-o", "--output", default="potential_reference_slide.pptx")
    a = ap.parse_args()
    build(a.template, a.output)
