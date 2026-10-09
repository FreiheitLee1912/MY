# -*- coding: utf-8 -*-
"""Guest活用と5拠点展開の費用比較（2ページ）。

Usage:
    python build_deck_guest_cost.py [-o out.pptx]

P1: Guest活用の考え方と各拠点への依頼
P2: 必要人数の集約と、費用3ケースの比較の枠組み

数値・期限が未確定のところは ［要確認］ として琥珀色で残し、
回答が入ったらそのまま差し替えられるようにしている。
"""
import argparse
import math

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn


def C(h):
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


BLUE = C("0096D8")
DARK = C("00558C")
PALE = C("E5F4FC")
GREEN = C("00843D")
GREEN_PALE = C("E7F4EC")
AMBER = C("C55A11")
AMBER_FILL = C("FDF1E3")
INK = C("1F2937")
GREY = C("6B7280")
RULE = C("D0D5DD")
FAINT = C("F4F5F3")
WHITE = C("FFFFFF")

JP = "Meiryo"
L, R = 0.42, 12.92
W = R - L

EM = 0.158            # 11pt Meiryo の折り返し見積り（安全側）
LH = 0.20


def style(run, size, bold=False, color=INK, face=JP, spacing=None):
    f = run.font
    f.size, f.bold, f.name = Pt(size), bold, face
    f.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    if spacing is not None:
        rPr.set("spc", str(int(spacing * 100)))
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {}); rPr.append(el)
        el.set("typeface", face)


def say(sl, x, y, w, h, text, size, bold=False, color=INK, align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP, space=None, spacing=None):
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    for i, ln in enumerate(text.split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = align
        if space:
            p.line_spacing = space
        r = p.add_run(); r.text = ln
        style(r, size, bold, color, spacing=spacing)
    return tf


def rich(sl, x, y, w, h, parts, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP,
         space=1.25):
    tb = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]; p.alignment = align; p.line_spacing = space
    for text, size, bold, color in parts:
        r = p.add_run(); r.text = text
        style(r, size, bold, color)
    return tf


def rect(sl, x, y, w, h, fill, outline=None, ow=1.0, shape=MSO_SHAPE.RECTANGLE):
    sh = sl.shapes.add_shape(shape, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill is None:
        sh.fill.background()
    else:
        sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if outline is None:
        sh.line.fill.background()
    else:
        sh.line.color.rgb = outline; sh.line.width = Pt(ow)
    sh.shadow.inherit = False
    sh.text_frame.word_wrap = True
    return sh


def chip(sl, x, y, w, h, text, fill, size=11, fg=WHITE, outline=None, ow=1.0):
    sh = rect(sl, x, y, w, h, fill, outline, ow)
    tf = sh.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    style(r, size, True, fg)
    return sh


def lines(text, inner_w):
    return max(1, math.ceil(len(text) / max(6, int(inner_w / EM))))


def bullet(sl, x, y, w, text, marker=BLUE, color=INK):
    h = lines(text, w - 0.20) * LH
    rect(sl, x, y + 0.075, 0.07, 0.07, marker)
    say(sl, x + 0.20, y, w - 0.20, h, text, 11, False, color, space=1.25)
    return h


def header(sl, eyebrow, headline, sub):
    say(sl, L, 0.36, 10.5, 0.26, eyebrow, 11.5, True, BLUE, spacing=1.0)
    say(sl, L, 0.62, W, 0.56, headline, 20, True, INK,
        anchor=MSO_ANCHOR.MIDDLE)
    say(sl, L, 1.20, 12.3, 0.26, sub, 12.5, False, GREY)


def sowhat(sl, y, h, text):
    rect(sl, L, y, W, h, BLUE)
    chip(sl, L + 0.16, y + 0.07, 1.22, h - 0.14, "So What", DARK, 11.5)
    say(sl, L + 1.58, y, W - 1.78, h, text, 13, True, WHITE,
        anchor=MSO_ANCHOR.MIDDLE)


def foot(sl, y, label, page):
    say(sl, L, y, 7.5, 0.24, label, 11, False, GREY)
    say(sl, 12.20, y, 0.72, 0.24, page, 11, False, GREY, align=PP_ALIGN.RIGHT)


# ---------------------------------------------------------------- P1
CARDS = [
    ("01", "Guest活用のポイント",
     ["有償アカウント1名につき、最大5名を無料で招待できる",
      "利用範囲は、指定されたSpaceに限定される"],
     "※ 代理店の説明に基づく。適用条件・権限は要確認。", None),
    ("02", "各拠点でGuest候補を整理",
     ["閲覧中心・簡単な作業・一時的な参加を想定する利用者を抽出する",
      "その利用者が必要な権限をGuestで満たせるかを確認する"],
     None, None),
    ("03", "各拠点への依頼",
     ["今後の利用予定人数（Guest候補を含む総数）をHQへご連絡ください",
      "そのうちGuest利用候補の人数もあわせてご連絡ください"],
     None, "回答期限：［要確認：各拠点の回答期限］"),
]
STEPS = ["利用内容を確認", "Guest候補を整理", "必要な有償人数を確定"]

# ---------------------------------------------------------------- P2
SITES = [("HQ", "14", "要確認", "要確認"),
         ("DSSE", "要確認", "要確認", "要確認"),
         ("DSST", "0", "要確認", "要確認"),
         ("DSSC", "0", "要確認", "要確認"),
         ("DSSA", "0", "要確認", "要確認")]
SITE_COLS = [("拠点", 2.20), ("現在の有償アカウント数", 3.20),
             ("展開後の利用予定人数 ※", 3.40), ("Guest利用候補人数", 3.70)]

PRICE_ROWS = [("［人数帯］", "［　］", "［　］"),
              ("［人数帯］", "［　］", "［　］"),
              ("［人数帯］", "［　］", "［　］")]
PRICE_COLS = [("人数帯", 2.05), ("現行 月額単価", 2.04), ("改定後 月額単価", 2.04)]

CASES = [("現在人数 × 現行価格", "現在の費用水準"),
         ("展開後人数 × 改定価格", "全員を有償とした場合の費用"),
         ("Guest活用後 × 改定価格", "Guestへの振り分けによる削減余地")]


def page1(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    header(sl, "① Guest活用 ｜ 5拠点展開の費用",
           "Guest活用で、5拠点展開に伴う費用増加を抑える",
           "HQとして、利用内容に応じて有償アカウントとGuestを使い分け、"
           "必要な有償人数を最適化する。")

    cy, ch = 1.56, 3.34
    cw = (W - 0.32) / 3
    for i, (no, title, items, note, ask) in enumerate(CARDS):
        cx = L + i * (cw + 0.16)
        rect(sl, cx, cy, cw, ch, WHITE, RULE, 0.75)
        rect(sl, cx, cy, cw, 0.06, BLUE)
        rect(sl, cx, cy + 0.06, cw, 0.62, PALE)
        say(sl, cx + 0.18, cy + 0.16, 0.50, 0.30, no, 18, True, BLUE)
        say(sl, cx + 0.74, cy + 0.20, cw - 0.92, 0.26, title, 13, True, INK)

        y = cy + 0.86
        for t in items:
            y += bullet(sl, cx + 0.18, y, cw - 0.36, t) + 0.14
        if note:
            say(sl, cx + 0.18, y + 0.04, cw - 0.36, 0.44, note, 11, False, GREY,
                space=1.25)
        if ask:
            h = 0.48
            by = cy + ch - h - 0.18
            rect(sl, cx + 0.18, by, cw - 0.36, h, AMBER_FILL, AMBER, 0.75)
            say(sl, cx + 0.30, by, cw - 0.60, h, ask, 11, True, AMBER,
                anchor=MSO_ANCHOR.MIDDLE)

    # 進め方
    fy, fh = 5.06, 1.00
    rect(sl, L, fy, W, fh, FAINT)
    say(sl, L + 0.22, fy + 0.12, 2.0, 0.22, "進め方", 11.5, True, DARK)
    bw, gap = 3.70, 0.46
    bx0 = L + (W - (bw * 3 + gap * 2)) / 2
    for i, name in enumerate(STEPS):
        bx = bx0 + i * (bw + gap)
        ch2 = rect(sl, bx, fy + 0.42, bw, 0.46, DARK if i == 2 else WHITE,
                   DARK, 1.0 if i == 2 else 0.75, MSO_SHAPE.CHEVRON)
        ch2.adjustments[0] = 0.22
        say(sl, bx + 0.26, fy + 0.42, bw - 0.52, 0.46, name, 11.5, i == 2,
            WHITE if i == 2 else INK, align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE)

    sowhat(sl, 6.52, 0.52,
           "全員を有償にせず、利用内容で切り分ければ、拠点を増やしても有償人数は増えない")
    foot(sl, 7.18, "Guest活用 ｜ 5拠点展開の費用", "1")

    sl.notes_slide.notes_text_frame.text = (
        "5拠点への展開に向け、HQではGuestの活用によって費用増加を抑えたいと考えています。"
        "代理店の説明では、有償アカウント1名につき最大5名を無料で招待できますが、利用範囲は"
        "指定されたSpaceに限られます。各拠点では、閲覧中心の方や一時的に参加する方などを"
        "Guest候補として整理してください。必要な権限を確認したうえで、有償アカウントが"
        "必要な人数を取りまとめます。")


def page2(prs):
    sl = prs.slides.add_slide(prs.slide_layouts[6])
    header(sl, "② 費用比較 ｜ 必要人数の集約",
           "必要人数を集約し、5拠点展開の費用とGuest活用の効果を比較する",
           "各拠点の回答と公式価格表を基に、展開後の費用とGuest活用による削減余地を"
           "確認する。")

    # ① 人数表
    say(sl, L, 1.54, 6.0, 0.22, "① 現在と今後の利用人数", 11.5, True, DARK)
    ty, hh, rh = 1.82, 0.34, 0.34
    x = L
    for name, cwid in SITE_COLS:
        chip(sl, x, ty, cwid - 0.02, hh, name, DARK, 11)
        x += cwid
    y = ty + hh
    for i, row in enumerate(SITES):
        x = L
        for j, (val, (_, cwid)) in enumerate(zip(row, SITE_COLS)):
            tbd = val == "要確認"
            rect(sl, x, y, cwid - 0.02, rh,
                 AMBER_FILL if tbd else (PALE if j == 0 else WHITE),
                 RULE, 0.75)
            say(sl, x, y, cwid - 0.02, rh, val, 11, j == 0 or not tbd,
                AMBER if tbd else (DARK if j == 0 else INK),
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            x += cwid
        y += rh
    say(sl, L, y + 0.06, 12.0, 0.22,
        "※ Guest候補を含む総利用人数。Guestの適用確認後、有償人数を確定する。",
        11, False, GREY)

    # ②③ 下段
    py, ph = 4.24, 1.96
    pw = (W - 0.24) / 2
    px = [L, L + pw + 0.24]

    rect(sl, px[0], py, pw, ph, WHITE, RULE, 0.75)
    chip(sl, px[0], py, pw, 0.34, "② Jira Standard の参考価格（USD／人・月）",
         DARK, 11.5)
    x = px[0] + 0.14
    for name, cwid in PRICE_COLS:
        rect(sl, x, py + 0.44, cwid, 0.28, PALE, RULE, 0.75)
        say(sl, x, py + 0.44, cwid, 0.28, name, 11, True, DARK,
            align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        x += cwid
    ry = py + 0.72
    for row in PRICE_ROWS:
        x = px[0] + 0.14
        for val, (_, cwid) in zip(row, PRICE_COLS):
            rect(sl, x, ry, cwid, 0.26, WHITE, RULE, 0.75)
            say(sl, x, ry, cwid, 0.26, val, 11, False, AMBER,
                align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
            x += cwid
        ry += 0.26
    say(sl, px[0] + 0.14, ry + 0.08, pw - 0.28,
        0.42, "月額は人数帯ごとの累進計算。各人数帯の人数に単価を掛けて合算する。",
        11, False, GREY, space=1.2)

    rect(sl, px[1], py, pw, ph, WHITE, RULE, 0.75)
    chip(sl, px[1], py, pw, 0.34, "③ 人数回答後に比較する3ケース", DARK, 11.5)
    rect(sl, px[1] + 0.14, py + 0.44, 2.60, 0.28, PALE, RULE, 0.75)
    rect(sl, px[1] + 2.74, py + 0.44, pw - 2.88, 0.28, PALE, RULE, 0.75)
    say(sl, px[1] + 0.14, py + 0.44, 2.60, 0.28, "比較ケース", 11, True, DARK,
        align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    say(sl, px[1] + 2.74, py + 0.44, pw - 2.88, 0.28, "確認すること", 11, True,
        DARK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    ry = py + 0.72
    for i, (case, check) in enumerate(CASES):
        last = i == len(CASES) - 1
        rect(sl, px[1] + 0.14, ry, 2.60, 0.30,
             GREEN_PALE if last else WHITE, RULE, 0.75)
        rect(sl, px[1] + 2.74, ry, pw - 2.88, 0.30,
             GREEN_PALE if last else WHITE, RULE, 0.75)
        say(sl, px[1] + 0.22, ry, 2.44, 0.30, case, 11, True,
            GREEN if last else DARK, anchor=MSO_ANCHOR.MIDDLE)
        say(sl, px[1] + 2.84, ry, pw - 3.08, 0.30, check, 11, False, INK,
            anchor=MSO_ANCHOR.MIDDLE)
        ry += 0.30
    say(sl, px[1] + 0.14, ry + 0.08, pw - 0.28, 0.42,
        "比較金額は各拠点の回答後に記載する。", 11, False, GREY, space=1.2)

    rect(sl, L, 6.26, W, 0.30, FAINT)
    rich(sl, L + 0.16, 6.26, W - 0.32, 0.30,
         [("試算条件｜", 11, True, DARK),
          ("改定価格は2026年10月13日（太平洋時間）から適用。5拠点を1契約に集約する"
           "前提で試算し、Guestの適用条件と最終費用は代理店に確認する。", 11,
           False, GREY)], anchor=MSO_ANCHOR.MIDDLE)

    sowhat(sl, 6.66, 0.46,
           "人数が揃えば、3ケースの差額でGuest活用の効果をそのまま金額で示せる")
    foot(sl, 7.18, "費用比較 ｜ 必要人数の集約", "2")

    sl.notes_slide.notes_text_frame.text = (
        "現在はHQとDSSEで利用しており、今後はDSST・DSSC・DSSAへの展開を予定しています。"
        "この表で、各拠点の利用予定人数とGuest候補人数を集約します。回答後は、現在の費用、"
        "全員を有償とした場合の展開後費用、Guestを活用した場合の費用を比較します。"
        "公式価格を基に試算し、最終的な契約金額は代理店の見積で確認します。")


def build(output):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    page1(prs)
    page2(prs)
    prs.save(output)
    print("saved", output)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("-o", "--output", default="guest_cost_daicel.pptx")
    build(ap.parse_args().output)
