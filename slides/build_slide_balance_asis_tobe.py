# -*- coding: utf-8 -*-
"""現状と課題｜需給バランス — マーケを起点に、生準と拠点を左右で比べる。

Usage:
    python build_slide_balance_asis_tobe.py [--lang ja|en] [-o out.pptx]

販売計画を出すマーケを上に置き、そこから分岐して生準（左）と拠点（右）に配る。
受け取る2者を左右に並べ、各パネルの中で As is / To be の共有文書を上下に置くので、
「左右＝だれが受け取るか」「上下＝いまと、これから」の2軸で読める。

日英で版面は同じ。折り返しに効く寸法だけ言語ごとに持ち替える（METRICS）。
"""
import argparse

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
RED = C("C00000")
RED_FILL = C("FBEAEA")
INK = C("1F2937")
GREY = C("6B7280")
SLATE = C("7F7F7F")
FAINT = C("F4F5F3")
RULE = C("D0D5DD")
WHITE = C("FFFFFF")

L, R = 0.42, 12.92
W = R - L
MID = (L + R) / 2

BODY = "Meiryo"              # build() が言語に応じて差し替える
DISPLAY = "Meiryo"


def style(run, size, bold=False, color=INK, face=None, spacing=None,
          strike=False):
    face = face or BODY
    f = run.font
    f.size, f.bold, f.name = Pt(size), bold, face
    f.color.rgb = color
    rPr = run._r.get_or_add_rPr()
    if spacing is not None:
        rPr.set("spc", str(int(spacing * 100)))
    if strike:
        rPr.set("strike", "sngStrike")
    for tag in ("a:ea", "a:cs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = rPr.makeelement(qn(tag), {}); rPr.append(el)
        el.set("typeface", face)


def say(sl, x, y, w, h, text, size, bold=False, color=INK, align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP, space=None, spacing=None, face=None):
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
        style(r, size, bold, color, face, spacing=spacing)
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


def chip(sl, x, y, w, h, text, fill, size=11, fg=WHITE, outline=None, ow=1.0,
         strike=False):
    sh = rect(sl, x, y, w, h, fill, outline, ow)
    tf = sh.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    style(r, size, True, fg, strike=strike)
    return sh


# チップの色づかい（予算は共通、月次=遅い、日次=速い、拠点限定=警告）
BUDGET = (GREEN_PALE, GREEN, GREEN)
SLOW = (AMBER_FILL, AMBER, AMBER)
FAST = (PALE, DARK, BLUE)
WARN = (RED_FILL, RED, RED)

JA = dict(
    eyebrow="① 現状と課題 ｜ 需給バランス",
    headline="違いは「配る文書」だけ — 月次を日次に変え、拠点にも同じものを配る",
    sub="販売計画はHQ営業が日々更新しているが、生準・拠点が変化点を確認するまでに"
        "時間差がある。流れそのものは変わらない。",
    src_tag="情報の起点",
    src_main="マーケ（Marketing）",
    src_role="HQ営業が販売計画を日々更新し、生準と拠点に配る",
    asis_label="As is", asis_cap="いま受け取っている文書",
    tobe_label="To be", tobe_cap="これから受け取る文書",
    dropped_cap="なくなるもの",
    panels=[
        dict(main="生準（Production Prep.）", role="需給を見て供給可否を判断",
             asis=[("バランス（予算）", BUDGET), ("バランス（月次）", SLOW)],
             asis_note="月次更新のため、日々の変化は次の月次まで伝わらない",
             tobe=[("バランス（予算）", BUDGET), ("バランス（日次）", FAST)],
             tobe_note="日次更新で、変化がその日のうちに伝わる",
             dropped=None),
        dict(main="拠点（Site）", role="設備・投資の検討に落とす",
             asis=[("予算販売計画", WARN)],
             asis_note="予算レンジ内のみ。レンジ外の将来案件は見えない",
             tobe=[("バランス（予算）", BUDGET), ("バランス（日次）", FAST)],
             tobe_note="マーケ・生準と同じものを、同じ鮮度で見られる",
             dropped="予算販売計画"),
    ],
    sowhat="仕組みを作り替える話ではなく、配る文書を「日次バランス」に揃えるだけで"
           "時間差と拠点の情報格差は解消する",
    foot="現状と課題｜需給バランス",
    notes="需給バランスの As is / To be。販売計画を出すマーケを起点に置き、そこから"
          "配られる先の生準と拠点を左右で比べている。As is では生準が月次バランスを"
          "受け取るため、日々の変化は次の月次まで伝わらない。拠点は予算レンジ内の"
          "予算販売計画しか受け取らず、レンジ外の将来案件を投資検討に織り込めない。"
          "To be は生準が日次バランスに切り替わり、変化がその日のうちに伝わる。拠点は"
          "予算販売計画に代えて同じ日次バランスを同じ鮮度で受け取るので、予算レンジ外を"
          "含む将来需要を双方が確認でき、大型案件や数量変動を早期に把握して供給対応を"
          "検討できる。変えるのは配る文書だけで、情報の流れ自体は変わらない。",
)

EN = dict(
    eyebrow="1. CURRENT STATE & ISSUES ｜ SUPPLY-DEMAND BALANCE",
    headline="Only the distributed document changes — monthly becomes "
             "daily, and the site gets it",
    sub="HQ sales updates the plan daily, but production preparation and the "
        "sites see the change only later. The flow itself does not change.",
    src_tag="Source",
    src_main="Marketing",
    src_role="Updates the sales plan daily and sends it to both",
    asis_label="As is", asis_cap="What they receive today",
    tobe_label="To be", tobe_cap="What they will get",
    dropped_cap="Dropped",
    panels=[
        dict(main="Production Prep.", role="Judges whether supply can be met",
             asis=[("Balance (Budget)", BUDGET), ("Balance (Monthly)", SLOW)],
             asis_note="Monthly refresh — a daily change only lands at the "
                       "next monthly cycle",
             tobe=[("Balance (Budget)", BUDGET), ("Balance (Daily)", FAST)],
             tobe_note="Daily refresh — a change lands the same day",
             dropped=None),
        dict(main="Site", role="Turns it into equipment and capex",
             asis=[("Budget sales plan", WARN)],
             asis_note="Budget range only — out-of-range future cases stay "
                       "invisible",
             tobe=[("Balance (Budget)", BUDGET), ("Balance (Daily)", FAST)],
             tobe_note="The same file as Marketing and Prod. Prep., at the same "
                       "freshness",
             dropped="Budget sales plan"),
    ],
    sowhat="Not a system rebuild — one shared daily balance removes the lag "
           "and the site's information gap",
    foot="Current state & issues | Supply-demand balance",
    notes="As is / To be for the supply-demand balance. Marketing, which issues "
          "the sales plan, sits at the top, and the two parties it distributes "
          "to - production preparation and the site - are compared left and "
          "right. Today production preparation receives a monthly balance, so a "
          "daily change only lands at the next monthly cycle, and the site "
          "receives only a budget sales plan limited to the budget range, "
          "leaving out-of-range future cases out of capex decisions. In To be "
          "production preparation moves to a daily balance and the site receives "
          "that same file instead of the budget sales plan, at the same "
          "freshness, so both see future demand including out-of-range cases and "
          "can plan supply together early. Only the distributed document "
          "changes; the flow itself does not.",
)

METRICS = dict(
    ja=dict(head_size=20, node_size=14, drop_w=1.72),
    en=dict(head_size=18, node_size=14, drop_w=1.86),
)

# --- 版面 ---
PW = (W - 0.24) / 2             # 左右パネルの幅
PX = [L, L + PW + 0.24]
PC = [x + PW / 2 for x in PX]   # パネルの中心（分岐線の落とし先）

SRC_W, SRC_Y, SRC_H = 6.30, 1.56, 0.72
SRC_TX = 1.52                   # タグチップのぶんの左インデント
FORK_Y = SRC_Y + SRC_H          # 分岐の横棒
BAR_Y = FORK_Y + 0.22

PY, PH = 2.78, 3.76
HDR_H = 0.64
SEC_H = 1.34
SEC_Y = [PY + 0.78, PY + 2.24]
SW_Y, SW_H = 6.66, 0.50
FOOT_Y = 7.22


def section(sl, px, y, label, cap, fill, accent, chips, note,
            dropped=None, drop_cap=None, drop_w=1.72):
    rect(sl, px + 0.12, y, PW - 0.24, SEC_H, fill, RULE, 0.75)
    chip(sl, px + 0.24, y + 0.14, 0.84, 0.24, label, accent, 11)
    cap_w = PW - 1.40 - (drop_w + 1.10 if dropped else 0)
    say(sl, px + 1.16, y + 0.14, cap_w, 0.24, cap, 11, False, GREY,
        anchor=MSO_ANCHOR.MIDDLE)
    if dropped:
        # 置き換わってなくなる文書は、打ち消し線のグレーチップで残す
        dx = px + PW - 0.24 - drop_w
        say(sl, dx - 1.06, y + 0.14, 1.00, 0.24, drop_cap, 11, False, GREY,
            align=PP_ALIGN.RIGHT, anchor=MSO_ANCHOR.MIDDLE)
        chip(sl, dx, y + 0.14, drop_w, 0.24, dropped, FAINT, 11, GREY,
             outline=RULE, ow=0.75, strike=True)

    inner = PW - 0.48
    n = len(chips)
    cw = 2.55 if n > 1 else 3.20
    gap = (inner - cw * n) / (n + 1)
    cx = px + 0.24 + gap
    for text, (cf, fg, oc) in chips:
        chip(sl, cx, y + 0.50, cw, 0.32, text, cf, 11, fg, outline=oc, ow=1.0)
        cx += cw + gap
    say(sl, px + 0.24, y + 0.94, inner, 0.26, note, 11, False, GREY,
        align=PP_ALIGN.CENTER)


def build(output, lang):
    global BODY, DISPLAY
    T = JA if lang == "ja" else EN
    M = METRICS[lang]
    BODY = "Meiryo" if lang == "ja" else "Arial"
    DISPLAY = "Meiryo" if lang == "ja" else "Arial Black"

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    sl = prs.slides.add_slide(prs.slide_layouts[6])

    # --- ヘッダー ---
    say(sl, L, 0.36, 10.5, 0.26, T["eyebrow"], 11.5, True, BLUE, spacing=1.0)
    say(sl, L, 0.62, W, 0.56, T["headline"], M["head_size"], True, INK,
        anchor=MSO_ANCHOR.MIDDLE, face=DISPLAY)
    say(sl, L, 1.20, 12.2, 0.26, T["sub"], 12.5, False, GREY)

    # --- 起点：マーケ ---
    sx = MID - SRC_W / 2
    rect(sl, sx, SRC_Y, SRC_W, SRC_H, DARK)
    chip(sl, sx + 0.14, SRC_Y + 0.24, 1.24, 0.24, T["src_tag"], BLUE, 11)
    say(sl, sx + SRC_TX, SRC_Y + 0.10, SRC_W - SRC_TX - 0.14, 0.26,
        T["src_main"], M["node_size"], True, WHITE)
    say(sl, sx + SRC_TX, SRC_Y + 0.40, SRC_W - SRC_TX - 0.14, 0.22,
        T["src_role"], 11, False, PALE)

    # 分岐（起点から左右のパネルへ）
    rect(sl, MID - 0.015, FORK_Y, 0.03, 0.22, BLUE)
    rect(sl, PC[0], BAR_Y - 0.015, PC[1] - PC[0], 0.03, BLUE)
    for c in PC:
        rect(sl, c - 0.13, BAR_Y, 0.26, 0.24, BLUE,
             shape=MSO_SHAPE.DOWN_ARROW)

    # --- 受け取る2者 ---
    for p, px in zip(T["panels"], PX):
        rect(sl, px, PY, PW, PH, WHITE, RULE, 0.75)
        rect(sl, px, PY, PW, HDR_H, DARK)
        say(sl, px, PY + 0.10, PW, 0.26, p["main"], M["node_size"], True, WHITE,
            align=PP_ALIGN.CENTER)
        say(sl, px, PY + 0.38, PW, 0.22, p["role"], 11, False, PALE,
            align=PP_ALIGN.CENTER)

        section(sl, px, SEC_Y[0], T["asis_label"], T["asis_cap"], FAINT, SLATE,
                p["asis"], p["asis_note"])
        section(sl, px, SEC_Y[1], T["tobe_label"], T["tobe_cap"], PALE, DARK,
                p["tobe"], p["tobe_note"], p["dropped"], T["dropped_cap"],
                M["drop_w"])

    # --- So What ---
    rect(sl, L, SW_Y, W, SW_H, BLUE)
    chip(sl, L + 0.16, SW_Y + 0.07, 1.22, SW_H - 0.14, "So What", DARK, 11.5)
    say(sl, L + 1.58, SW_Y, W - 1.78, SW_H, T["sowhat"], 13, True, WHITE,
        anchor=MSO_ANCHOR.MIDDLE)

    say(sl, L, FOOT_Y, 7.5, 0.24, T["foot"], 11, False, GREY)
    say(sl, 12.20, FOOT_Y, 0.72, 0.24, "12", 11, False, GREY,
        align=PP_ALIGN.RIGHT)

    sl.notes_slide.notes_text_frame.text = T["notes"]
    prs.save(output)
    print("saved", output)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--lang", choices=("ja", "en"), default="ja")
    ap.add_argument("-o", "--output")
    a = ap.parse_args()
    build(a.output or "balance_asis_tobe_daicel%s.pptx" %
          ("" if a.lang == "ja" else "_en"), a.lang)
