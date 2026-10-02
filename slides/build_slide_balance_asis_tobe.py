# -*- coding: utf-8 -*-
"""現状と課題｜需給バランス — HQ（マーケ＋生準）と拠点を左右で比べる。

Usage:
    python build_slide_balance_asis_tobe.py [--lang ja|en] [-o out.pptx]

マーケと生準は同じ文書を見ているので左パネルに上下で積み、その右に拠点を置く。
各パネルの中で As is / To be の共有文書を上下に並べるので、
「左右＝だれが持っているか」「上下＝いまと、これから」の2軸で読める。

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

BODY = "Meiryo"              # build() が言語に応じて差し替える
DISPLAY = "Meiryo"


def style(run, size, bold=False, color=INK, face=None, spacing=None):
    face = face or BODY
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


def chip(sl, x, y, w, h, text, fill, size=11, fg=WHITE, outline=None, ow=1.0):
    sh = rect(sl, x, y, w, h, fill, outline, ow)
    tf = sh.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    style(r, size, True, fg)
    return sh


# チップの色づかい（予算は共通、月次=遅い、日次=速い、拠点限定=警告）
BUDGET = (GREEN_PALE, GREEN, GREEN)
SLOW = (AMBER_FILL, AMBER, AMBER)
FAST = (PALE, DARK, BLUE)
WARN = (RED_FILL, RED, RED)

JA = dict(
    eyebrow="① 現状と課題 ｜ 需給バランス",
    headline="違いは「共有する文書」だけ — 月次を日次に変え、拠点にも同じものを配る",
    sub="販売計画はHQ営業が日々更新しているが、生準・拠点が変化点を確認するまでに"
        "時間差がある。流れそのものは変わらない。",
    asis_label="As is", asis_cap="いま共有している文書",
    tobe_label="To be", tobe_cap="これから共有する文書",
    panels=[
        dict(title="マーケ ＋ 生準（HQ）｜2者は同じ文書を見ている",
             nodes=[("マーケ", "HQ営業が販売計画を日々更新"),
                    ("生準", "需給を見て供給可否を判断")],
             asis=[("バランス（予算）", BUDGET), ("バランス（月次）", SLOW)],
             asis_note="月次更新のため、日々の変化は次の月次まで伝わらない",
             tobe=[("バランス（予算）", BUDGET), ("バランス（日次）", FAST)],
             tobe_note="日次更新で、変化がその日のうちに伝わる"),
        dict(title="拠点（Site）｜HQとは別のものを見ている",
             nodes=[("拠点", "設備・投資の検討に落とす")],
             asis=[("予算販売計画", WARN)],
             asis_note="予算レンジ内のみ。レンジ外の将来案件は見えない",
             tobe=[("バランス（予算）", BUDGET), ("バランス（日次）", FAST)],
             tobe_note="HQと同じものを、同じ鮮度で見られる"),
    ],
    sowhat="仕組みを作り替える話ではなく、配る文書を「日次バランス」に揃えるだけで"
           "時間差と拠点の情報格差は解消する",
    foot="現状と課題｜需給バランス",
    notes="需給バランスの As is / To be。マーケと生準は同じ文書を見ているので左に上下で"
          "積み、その右に拠点を置いて左右で比べている。As is ではHQ側が月次バランスを"
          "共有し、日々の変化は次の月次まで伝わらない。拠点は予算レンジ内の予算販売計画"
          "しか持たず、レンジ外の将来案件を投資検討に織り込めない。To be はHQ側が日次"
          "バランスに切り替わり、変化がその日のうちに伝わる。拠点にも同じものが同じ鮮度で"
          "渡るので、予算レンジ外を含む将来需要を双方が確認でき、大型案件や数量変動を"
          "早期に把握して供給対応を検討できる。変えるのは配る文書だけで、情報の流れ自体は"
          "変わらない。",
)

EN = dict(
    eyebrow="1. CURRENT STATE & ISSUES ｜ SUPPLY-DEMAND BALANCE",
    headline="Only the shared document changes — monthly becomes daily, "
             "and the site gets it",
    sub="HQ sales updates the plan daily, but production preparation and the "
        "sites see the change only later. The flow itself does not change.",
    asis_label="As is", asis_cap="What they share today",
    tobe_label="To be", tobe_cap="What they will share",
    panels=[
        dict(title="Marketing + Production Prep. (HQ) | both see the same document",
             nodes=[("Marketing", "HQ sales updates the plan daily"),
                    ("Production Prep.", "Judges whether supply can be met")],
             asis=[("Balance (Budget)", BUDGET), ("Balance (Monthly)", SLOW)],
             asis_note="Monthly refresh — a daily change only lands at the "
                       "next monthly cycle",
             tobe=[("Balance (Budget)", BUDGET), ("Balance (Daily)", FAST)],
             tobe_note="Daily refresh — a change lands the same day"),
        dict(title="Site | sees something different from HQ",
             nodes=[("Site", "Turns it into equipment and capex")],
             asis=[("Budget sales plan", WARN)],
             asis_note="Budget range only — out-of-range future cases stay "
                       "invisible",
             tobe=[("Balance (Budget)", BUDGET), ("Balance (Daily)", FAST)],
             tobe_note="The same file as HQ, at the same freshness"),
    ],
    sowhat="Not a system rebuild — one shared daily balance removes the lag "
           "and the site's information gap",
    foot="Current state & issues | Supply-demand balance",
    notes="As is / To be for the supply-demand balance. Marketing and Production "
          "Preparation see the same document, so they are stacked on the left, "
          "with the site beside them for a left-right comparison. Today the HQ "
          "side shares a monthly balance, so a daily change only lands at the "
          "next monthly cycle, and the site holds only a budget sales plan "
          "limited to the budget range, leaving out-of-range future cases out of "
          "capex decisions. In To be the HQ side moves to a daily balance and the "
          "site receives the same file at the same freshness, so both see future "
          "demand including out-of-range cases and can plan supply together "
          "early. Only the shared document changes; the flow itself does not.",
)

METRICS = dict(
    ja=dict(head_size=20, title_size=11.5, node_size=14),
    en=dict(head_size=18, title_size=11, node_size=13.5),
)

# --- 版面 ---
PW = (W - 0.24) / 2             # 左右パネルの幅
PX = [L, L + PW + 0.24]
PY, PH = 1.54, 4.88             # パネルの天地
HDR_H = 0.46                    # パネル見出し帯
NODE_Y, NODE_H = 2.16, 1.40     # ノード置き場（左は2段、右は1段を同じ高さで）
BOX_H = 0.58
SEC = [(3.72, "asis"), (5.06, "tobe")]   # As is / To be セクションの上端
SEC_H = 1.20
SW_Y, SW_H = 6.56, 0.52
FOOT_Y = 7.16


def section(sl, px, y, label, cap, fill, accent, chips, note):
    rect(sl, px + 0.12, y, PW - 0.24, SEC_H, fill, RULE, 0.75)
    chip(sl, px + 0.24, y + 0.12, 0.84, 0.24, label, accent, 11)
    say(sl, px + 1.16, y + 0.12, PW - 1.40, 0.24, cap, 11, False, GREY,
        anchor=MSO_ANCHOR.MIDDLE)
    inner = PW - 0.48
    n = len(chips)
    cw = 2.55 if n > 1 else 3.20
    gap = (inner - cw * n) / (n + 1)
    cx = px + 0.24 + gap
    for text, (cf, fg, oc) in chips:
        chip(sl, cx, y + 0.44, cw, 0.30, text, cf, 11, fg, outline=oc, ow=1.0)
        cx += cw + gap
    say(sl, px + 0.24, y + 0.82, inner, 0.26, note, 11, False, GREY,
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

    for p, px in zip(T["panels"], PX):
        rect(sl, px, PY, PW, PH, WHITE, RULE, 0.75)
        rect(sl, px, PY, PW, HDR_H, DARK)
        say(sl, px + 0.22, PY, PW - 0.44, HDR_H, p["title"], M["title_size"],
            True, WHITE, anchor=MSO_ANCHOR.MIDDLE)

        # ノード：左は2段、右は同じ高さの1段
        nodes = p["nodes"]
        if len(nodes) == 2:
            ys = [NODE_Y, NODE_Y + NODE_H - BOX_H]
            hs = [BOX_H, BOX_H]
            rect(sl, px + PW / 2 - 0.13, NODE_Y + BOX_H + 0.02, 0.26, 0.18,
                 BLUE, shape=MSO_SHAPE.DOWN_ARROW)
        else:
            ys, hs = [NODE_Y + (NODE_H - 0.86) / 2], [0.86]
        for (main, role), ny, nh in zip(nodes, ys, hs):
            rect(sl, px + 0.12, ny, PW - 0.24, nh, DARK)
            say(sl, px + 0.12, ny + nh / 2 - 0.27, PW - 0.24, 0.26, main,
                M["node_size"], True, WHITE, align=PP_ALIGN.CENTER)
            say(sl, px + 0.12, ny + nh / 2 + 0.02, PW - 0.24, 0.22, role, 11,
                False, PALE, align=PP_ALIGN.CENTER)

        for (sy, key) in SEC:
            asis = key == "asis"
            section(sl, px, sy,
                    T["asis_label"] if asis else T["tobe_label"],
                    T["asis_cap"] if asis else T["tobe_cap"],
                    FAINT if asis else PALE, SLATE if asis else DARK,
                    p[key], p[key + "_note"])

    # HQ → 拠点（情報の流れ）
    rect(sl, PX[1] - 0.25, NODE_Y + NODE_H / 2 - 0.13, 0.26, 0.26, BLUE,
         shape=MSO_SHAPE.RIGHT_ARROW)

    # --- So What ---
    rect(sl, L, SW_Y, W, SW_H, BLUE)
    chip(sl, L + 0.16, SW_Y + 0.08, 1.22, SW_H - 0.16, "So What", DARK, 11.5)
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
