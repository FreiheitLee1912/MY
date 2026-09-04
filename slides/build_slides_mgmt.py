# -*- coding: utf-8 -*-
"""管理内容／本社と拠点役割分担 — レビュー修正版の2枚.

Usage:
    python build_slides_mgmt.py --template 2026_Standard_Template_WideScreenEN.pptx

マスターはダイセル標準テンプレート、書体はMeiryo。
本文の配色は元スライドに合わせた濃紺パネル＋クリーム帯。
"""
import argparse

from pptx import Presentation
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from titus_kit import C, BLACK, WHITE, style, say, rich, cell
from daicel_kit import (L, R, W, MUTED, RED, base_from_template, foot, page)

# ---- 元スライドから合わせた本文配色 -----------------------------------------
NAVY = C("1F3864")        # パネル見出し・下段バンド
NAVY_TXT = C("1F3864")
CREAM = C("FDF6E0")       # キーメッセージ帯
PALE = C("EDF1F7")        # サブボックス
THEAD = C("D9E2F3")       # 表ヘッダ（淡青／濃紺文字）
RULE = C("BFBFBF")
YELLOW = C("FFF2A8")      # 要確認タグ


def box(sl, x, y, w, h, text="", size=11, bold=False, fg=BLACK, fill=WHITE,
        border=RULE, align=PP_ALIGN.LEFT, pad=0.12, space=None, bw=0.75):
    return cell(sl, x, y, w, h, text, size, bold, fg, fill, border, bw,
                align=align, pad=pad, space=space)


def bar(sl, x, y, w, h, text, size=12.5, fill=NAVY, fg=WHITE,
        align=PP_ALIGN.LEFT, pad=0.16):
    return cell(sl, x, y, w, h, text, size, True, fg, fill, None,
                align=align, pad=pad)


def head(sl, title, deck=None, tag=None, page_no=None):
    """タイトル＋クリーム帯。tag を渡すと帯の右に要確認ボックスを置く。"""
    sl_ = page(sl, title, None, (), page_no)
    return sl_


# ============================================================== P11 ==========
def slide_p11(prs, n):
    sl = page(prs, "生準の日常管理は、進捗と課題の2つに絞る", None, (), n)
    # キーメッセージ帯
    box(sl, L, 0.98, W, 0.50,
        "JIRA上で計画・進捗・課題を関係者間で共有し、判断と承認につなげる。",
        13, True, BLACK, CREAM, None, pad=0.20)

    py = 1.66
    pw = (W - 0.30) / 2
    panels = [
        ("① 進捗管理", "拠点のスケジュールを基準に、案件全体の進み具合を把握する。",
         [("管理単位", "主要マイルストーン・タスク"),
          ("記録項目", "計画日程／現在の進捗／完了実績／遅延有無"),
          ("更新契機・更新者", "［要確認：更新頻度］／拠点担当"),
          ("遅延発生時", "遅延理由と今後の対応を登録し、［要確認：報告先の役割名］へ"
                        "報告・判断を依頼する")], 0.54),
        ("② 課題管理", "案件推進中に発生した課題をJIRA上で一元管理する。",
         [("管理単位", "課題1件"),
          ("記録項目", "課題内容（事象と案件への影響）／対応方針／"
                      "担当者・担当部門／対応期限"),
          ("更新契機・更新者", "状況が動いた都度／発生元の拠点"),
          ("ステータス", "未着手・対応中・完了"),
          ("意思決定の記録", "本社判断・支援の要否と、決定内容を追跡する")], 0.432),
    ]
    for i, (title, lead, rows, rh) in enumerate(panels):
        x = L + i * (pw + 0.30)
        bar(sl, x, py, pw, 0.40, title, 13)
        box(sl, x, py + 0.40, pw, 0.34, lead, 11, True, NAVY_TXT, PALE, pad=0.16)
        ry = py + 0.74
        for k, v in rows:
            box(sl, x, ry, 1.65, rh, k, 10, True, NAVY_TXT, WHITE, pad=0.12)
            box(sl, x + 1.65, ry, pw - 1.65, rh, v, 9.5, space=1.2)
            ry += rh

    # 下段バンド
    by = 4.76
    bar(sl, L, by, W, 0.36, "JIRA上で実現する状態", 12.5, align=PP_ALIGN.CENTER)
    states = [("計画の確認", "計画日程・進捗・実績をタスク単位で確認できる"),
              ("課題の検知", "遅延・課題・影響を一覧で把握できる"),
              ("責任の明確化", "トラブル発生時に、対応する部門と決定内容が記録として残る")]
    sw = (W - 0.40) / 3
    for i, (t, d) in enumerate(states):
        x = L + i * (sw + 0.20)
        box(sl, x, by + 0.40, sw, 0.32, t, 11.5, True, NAVY_TXT, PALE,
            align=PP_ALIGN.CENTER)
        box(sl, x, by + 0.72, sw, 0.54, d, 10, align=PP_ALIGN.CENTER, space=1.25)

    # 本日の論点
    ny = by + 1.36
    box(sl, L, ny, W, 0.62, "", fill=YELLOW, border=None)
    rich(sl, L + 0.20, ny + 0.08, W - 0.40, 0.46,
         [("本日の論点　", 11, True, RED),
          ("トラブル発生時に対応部門をどう決めるか。"
           "［要確認：本社が指示するのか、本社・拠点の協議で決めるのか。"
           "P12の協業モデルと整合させる必要がある］", 10, False, BLACK)],
         anchor=MSO_ANCHOR.MIDDLE, space=1.22)


# ============================================================== P12 ==========
def slide_p12(prs, n):
    sl = page(prs, "本社と拠点は上下関係ではなく、管理対象の違いで役割を分ける",
              None, (), n)
    # キーメッセージ帯（右に要確認タグ）
    bw_ = 8.10
    box(sl, L, 0.96, bw_, 0.54, "", fill=CREAM, border=None)
    say(sl, L + 0.20, 1.02, bw_ - 0.40, 0.22,
        "拠点＝案件単位の計画・実行　／　本社＝拠点横断の調整・標準化", 12, True, BLACK)
    rich(sl, L + 0.20, 1.26, bw_ - 0.40, 0.22,
         [("ご依頼　", 11, True, RED),
          ("本役割分担案のご承認をいただきたい。［要確認：承認者・承認期限］",
           11, False, BLACK)])
    box(sl, L + bw_ + 0.20, 0.96, W - bw_ - 0.20, 0.54, "", fill=YELLOW, border=BLACK)
    say(sl, L + bw_ + 0.34, 1.02, W - bw_ - 0.48, 0.44,
        "要確認：管理主体タスクレベル／\n判断が分かれた場合の裁定者／運用開始時期",
        10, True, BLACK, space=1.2)

    # 3つの原則
    gy = 1.62
    gw = (W - 0.40) / 3
    principles = [
        ("① 拠点が案件単位の計画・実行を担う", "営業・顧客に最も近く、顧客日程の一次情報を持つため"),
        ("② 本社が拠点横断の調整・標準化を担う", "全社案件の把握、拠点間調整、管理方法の標準化を行うため"),
        ("③ 両者は協業関係とする", "上下関係ではなく、管理対象の違いに基づく分担であるため"),
    ]
    for i, (t, why) in enumerate(principles):
        x = L + i * (gw + 0.20)
        bar(sl, x, gy, gw, 0.32, t, 11)
        box(sl, x, gy + 0.32, gw, 0.44, why, 9.5, fill=PALE, space=1.2)

    # 用語の定義
    ty = 2.50
    bar(sl, L, ty, 1.55, 0.44, "用語の定義", 11, align=PP_ALIGN.CENTER, pad=0.06)
    terms = [("主体", "意思決定と実行の責任を持つ"),
             ("支援", "依頼に応じて助言・作業を分担する"),
             ("共有", "結果の連携を受ける（作業責任は負わない）"),
             ("情報提供", "判断材料を提出する")]
    tw = (W - 1.55) / 4
    for i, (t, d) in enumerate(terms):
        x = L + 1.55 + i * tw
        box(sl, x, ty, tw, 0.44, "", fill=WHITE)
        rich(sl, x + 0.09, ty + 0.03, tw - 0.18, 0.38,
             [(t + "　", 9.5, True, NAVY_TXT), (d, 8.5, False, BLACK)],
             anchor=MSO_ANCHOR.MIDDLE, space=1.18)

    # 役割分担マトリックス
    my = 3.06
    cols = [("管理項目", 3.30), ("拠点", 1.15), ("本社", 1.15), ("関与内容", W - 5.60)]
    x = L
    for t, w in cols:
        box(sl, x, my, w, 0.28, t, 10, True, NAVY_TXT, THEAD, align=PP_ALIGN.CENTER)
        x += w

    groups = [
        ("［1］案件レベル｜拠点が計画・実行の主体（関与内容＝本社）",
         [("顧客日程の確認", "主体", "共有", "拠点が確認した一次情報の連携を受ける"),
          ("案件スケジュールの策定", "主体", "支援", "［要確認：支援の具体内容］"),
          ("主要マイルストーンの設定", "主体", "支援",
           "拠点横断の観点から意見を提示する。［要確認：最終決定者］"),
          ("日々の進捗更新", "主体", "共有", "JIRA上で状況の連携を受ける"),
          ("遅延・課題の登録", "主体", "共有", "発生元である拠点が登録し、本社は連携を受ける"),
          ("課題対応・担当者・期限設定", "主体", "支援",
           "拠点からの依頼に応じて対応を分担する")]),
        ("［2］拠点横断レベル｜本社が調整・標準化の主体（関与内容＝拠点）",
         [("拠点間のリソース・日程調整", "情報提供", "主体",
           "自拠点の状況を提出し、本社の調整結果に従う"),
          ("全社案件ポートフォリオ管理", "情報提供", "主体", "自拠点の案件情報を提出する"),
          ("管理基準・テンプレートの標準化", "運用", "主体",
           "本社が定めた共通ルールを運用する"),
          ("拠点横断課題の整理", "情報提供", "主体", "自拠点で検知した横断課題を提出する")]),
    ]
    ry = my + 0.28
    for gi, (label, rows) in enumerate(groups):
        box(sl, L, ry, W, 0.26, label, 10, True, WHITE, NAVY, pad=0.14)
        ry += 0.26
        for item, base, hq, how in rows:
            box(sl, L, ry, 3.30, 0.28, item, 9.5, pad=0.12)
            box(sl, L + 3.30, ry, 1.15, 0.28, base, 9.5, base == "主体", NAVY_TXT,
                PALE if base == "主体" else WHITE, align=PP_ALIGN.CENTER)
            box(sl, L + 4.45, ry, 1.15, 0.28, hq, 9.5, hq == "主体", NAVY_TXT,
                PALE if hq == "主体" else WHITE, align=PP_ALIGN.CENTER)
            box(sl, L + 5.60, ry, W - 5.60, 0.28, how, 9.5, pad=0.12)
            ry += 0.28


def build(template, output):
    prs = Presentation(base_from_template(template))
    slide_p11(prs, 11)
    slide_p12(prs, 12)
    prs.save(output)
    print("saved", output, "-", len(prs.slides._sldIdLst), "slides")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--template", required=True)
    ap.add_argument("-o", "--output", default="mgmt_role_slides.pptx")
    a = ap.parse_args()
    build(a.template, a.output)
