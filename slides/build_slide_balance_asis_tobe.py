# -*- coding: utf-8 -*-
"""現状と課題｜需給バランス — As is / To be を1つの流れで比較する。

Usage:
    python build_slide_balance_asis_tobe.py [--lang ja|en] [-o out.pptx]

As is と To be の違いは「3者が共有している文書」だけなので、
マーケ → 生準 → 拠点 の流れは1列だけ描き、その下に As is の共有文書、
To be の共有文書を重ねて差分を1か所で読ませる。

日英で版面は同じ。英文は折り返し幅が違うので、帯の高さと折り返し見積りだけを
言語ごとに持ち替える（METRICS）。
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
LH11 = 0.20                   # 11pt / 行送り1.25 の1行

BODY = "Meiryo"              # build() が言語に応じて差し替える
DISPLAY = "Meiryo"
EM_EST = 0.158               # 折り返し見積りの1文字幅（安全側）


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


def chip(sl, x, y, w, h, text, fill, size=11, fg=WHITE, outline=None, ow=1.0,
         bold=True):
    sh = rect(sl, x, y, w, h, fill, outline, ow)
    tf = sh.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    style(r, size, bold, fg)
    return sh


def bullet(sl, x, y, w, text, marker, color=INK):
    cpl = max(6, int((w - 0.20) / EM_EST))
    h = max(1, math.ceil(len(text) / cpl)) * LH11
    rect(sl, x, y + 0.075, 0.07, 0.07, marker)
    say(sl, x + 0.20, y, w - 0.20, h, text, 11, False, color, space=1.25)
    return h


# --- チップの色づかい（予算は共通、月次=遅い、日次=速い、拠点限定=警告） ---
BUDGET = (GREEN_PALE, GREEN, GREEN)
SLOW = (AMBER_FILL, AMBER, AMBER)
FAST = (PALE, DARK, BLUE)
WARN = (RED_FILL, RED, RED)

JA = dict(
    eyebrow="① 現状と課題 ｜ 需給バランス",
    headline="違いは「3者が共有する文書」だけ — 月次を日次に変え、拠点にも同じものを配る",
    sub="販売計画はHQ営業が日々更新しているが、生準・拠点が変化点を確認するまでに"
        "時間差がある。流れそのものは変わらない。",
    flow_label="情報の流れ\n（共通）",
    nodes=[("マーケ", "Marketing", "HQ営業が販売計画を日々更新"),
           ("生準", "Production Prep.", "需給を見て供給可否を判断"),
           ("拠点", "Site", "設備・投資の検討に落とす")],
    asis_label="As is", asis_cap="いま共有して\nいる文書",
    tobe_label="To be", tobe_cap="これから共有\nする文書",
    asis=[(0, 2, "マーケ・生準で共通",
           [("バランス（予算）", BUDGET), ("バランス（月次）", SLOW)], None),
          (2, 1, "拠点だけ別もの", [("予算販売計画", WARN)],
           "予算レンジ内のみ。レンジ外の将来案件は見えない")],
    tobe=[(0, 2, "マーケ・生準で共通",
           [("バランス（予算）", BUDGET), ("バランス（日次）", FAST)], None),
          (2, 1, "拠点も同じものに",
           [("バランス（予算）", BUDGET), ("バランス（日次）", FAST)], None)],
    issues_title="As is の課題",
    issues=["生準が販売計画の変化を把握するまでに時間差がある",
            "拠点はレンジ内の案件しか見えず、投資の検討に織り込めない"],
    gains_title="To be でできること",
    gains=["HQと拠点の両方が、予算レンジ外を含む将来需要を確認できる",
           "大型案件や数量変動を早期に把握し、HQ・拠点で連携して供給対応を検討できる"],
    sowhat="仕組みを作り替える話ではなく、配る文書を「日次バランス」に揃えるだけで"
           "時間差と拠点の情報格差は解消する",
    foot="現状と課題｜需給バランス",
    notes="需給バランスの As is / To be。マーケ→生準→拠点という情報の流れ自体は変わらず、"
          "差分は3者が共有している文書だけなので、流れは1列だけ描いて下に As is と To be の"
          "共有文書を重ねている。As is ではマーケと生準が月次バランスを共有し、拠点は予算"
          "レンジ内の予算販売計画しか持たない。そのため生準が販売計画の変化を把握するまでに"
          "時間差が生じ、拠点はレンジ外の将来案件を投資検討に織り込めない。To be は3者が同じ"
          "日次バランスを同じ鮮度で見る状態。HQと拠点の双方が予算レンジ外を含む将来需要を"
          "確認でき、大型案件や数量変動を早期に把握して供給対応を検討できる。",
)

EN = dict(
    eyebrow="1. CURRENT STATE & ISSUES ｜ SUPPLY-DEMAND BALANCE",
    headline="Only the shared document changes \u2014 monthly becomes daily, "
             "and the site gets it",
    sub="HQ sales updates the plan daily, but production preparation and the "
        "sites see the change only later. The flow itself does not change.",
    flow_label="Information flow\n(unchanged)",
    nodes=[("Marketing", "HQ Sales", "Updates the sales plan daily"),
           ("Production Prep.", "Supply planning", "Judges whether supply can be met"),
           ("Site", "Plant", "Turns it into equipment and capex")],
    asis_label="As is", asis_cap="What they\nshare today",
    tobe_label="To be", tobe_cap="What they\nwill share",
    asis=[(0, 2, "Shared by Marketing & Prod. Prep.",
           [("Balance (Budget)", BUDGET), ("Balance (Monthly)", SLOW)], None),
          (2, 1, "Site has something else", [("Budget sales plan", WARN)],
           "Budget range only \u2014 out-of-range cases stay invisible")],
    tobe=[(0, 2, "Shared by Marketing & Prod. Prep.",
           [("Balance (Budget)", BUDGET), ("Balance (Daily)", FAST)], None),
          (2, 1, "Site now gets the same",
           [("Balance (Budget)", BUDGET), ("Balance (Daily)", FAST)], None)],
    issues_title="Issues with As is",
    issues=["Production preparation sees sales-plan changes only after a lag",
            "Sites see only in-range cases and cannot factor them into capex"],
    gains_title="What To be enables",
    gains=["HQ and sites both see future demand, including out-of-range cases",
           "Volume swings are caught early for joint HQ-site supply planning"],
    sowhat="Not a system rebuild — one shared daily balance removes the lag and "
           "the site's information gap",
    foot="Current state & issues | Supply-demand balance",
    notes="As is / To be for the supply-demand balance. The information flow "
          "Marketing to Production Preparation to Site does not change; the only "
          "difference is the document the three share, so the flow is drawn once "
          "and the As is and To be documents are stacked beneath it. Today "
          "Marketing and Production Preparation share a monthly balance while the "
          "site holds only a budget sales plan limited to the budget range. "
          "Production preparation therefore sees sales-plan changes only after a "
          "lag, and sites cannot factor out-of-range future cases into capex "
          "decisions. In To be all three see the same daily balance at the same "
          "freshness, so HQ and sites both see future demand including "
          "out-of-range cases and can plan supply together early.",
)

# 版面は共通、折り返しに効く寸法だけ言語で持ち替える
METRICS = dict(
    ja=dict(em=0.158, blk_y=1.52, blk_h=3.90, node_y=1.72, node_h=0.60,
            node_size=15, head_size=20, as_y=2.60, as_h=1.38, to_y=4.04,
            to_h=1.24, exp_y=5.52, exp_h=1.10, sw_y=6.70, sw_h=0.46,
            foot_y=7.22),
    en=dict(em=0.085, blk_y=1.48, blk_h=3.94, node_y=1.68, node_h=0.60,
            node_size=14, head_size=18, as_y=2.58, as_h=1.40, to_y=4.04,
            to_h=1.24, exp_y=5.52, exp_h=1.10, sw_y=6.70, sw_h=0.46,
            foot_y=7.22),
)

GUT = 1.38                      # 行ラベルの左ガター
CX0 = L + GUT + 0.18
COLW = 3.30
CGAP = (R - CX0 - COLW * 3) / 2


def band(sl, y, h, label, caption, label_fill, fill, cells):
    """cells: (開始列, 列スパン, セル見出し, チップ, 注記) の並び。

    マーケと生準は同じ文書を見ているので、2列ぶんを1セルにまとめて描く。
    """
    rect(sl, L, y, W, h, fill)
    chip(sl, L + 0.22, y + 0.16, GUT - 0.20, 0.34, label, label_fill, 13)
    say(sl, L + 0.22, y + 0.56, GUT - 0.20, 0.42, caption, 11, False, GREY,
        align=PP_ALIGN.CENTER, space=1.15)
    for col, span, head, chips, note in cells:
        cx = CX0 + col * (COLW + CGAP)
        cellw = span * COLW + (span - 1) * CGAP
        rect(sl, cx, y + 0.10, cellw, h - 0.20, WHITE, RULE, 0.75)
        say(sl, cx, y + 0.18, cellw, 0.22, head, 11, True, DARK,
            align=PP_ALIGN.CENTER)
        cw = 3.60 if span > 1 else 2.90
        cy = y + 0.44
        for text, (cf, fg, oc) in chips:
            chip(sl, cx + (cellw - cw) / 2, cy, cw, 0.30, text, cf, 11, fg,
                 outline=oc, ow=1.0)
            cy += 0.38
        if note:
            say(sl, cx + 0.10, cy - 0.02, cellw - 0.20, 0.46, note, 11, False,
                GREY, align=PP_ALIGN.CENTER, space=1.15)


def build(output, lang):
    global BODY, DISPLAY, EM_EST
    T = JA if lang == "ja" else EN
    M = METRICS[lang]
    BODY = "Meiryo" if lang == "ja" else "Arial"
    DISPLAY = "Meiryo" if lang == "ja" else "Arial Black"
    EM_EST = M["em"]

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    sl = prs.slides.add_slide(prs.slide_layouts[6])

    # --- ヘッダー ---
    say(sl, L, 0.36, 10.5, 0.26, T["eyebrow"], 11.5, True, BLUE, spacing=1.0)
    say(sl, L, 0.62, W, 0.56, T["headline"], M["head_size"], True, INK,
        anchor=MSO_ANCHOR.MIDDLE, face=DISPLAY)
    say(sl, L, 1.20, 12.2, 0.26, T["sub"], 12.5, False, GREY)

    # --- 比較ブロック ---
    rect(sl, L, M["blk_y"], W, M["blk_h"], WHITE, RULE, 0.75)

    ny, nh = M["node_y"], M["node_h"]
    say(sl, L + 0.22, ny + 0.12, GUT - 0.20, 0.44, T["flow_label"], 11, True,
        DARK, align=PP_ALIGN.CENTER, space=1.15)
    for i, (main, sub, role) in enumerate(T["nodes"]):
        cx = CX0 + i * (COLW + CGAP)
        rect(sl, cx, ny, COLW, nh, DARK)
        say(sl, cx, ny + 0.06, COLW, 0.26, main, M["node_size"], True, WHITE,
            align=PP_ALIGN.CENTER)
        say(sl, cx, ny + nh - 0.28, COLW, 0.22, sub, 11, False, PALE,
            align=PP_ALIGN.CENTER)
        say(sl, cx, ny + nh + 0.05, COLW, 0.20, role, 11, False, GREY,
            align=PP_ALIGN.CENTER)
        if i < len(T["nodes"]) - 1:
            rect(sl, cx + COLW + 0.06, ny + (nh - 0.30) / 2, CGAP - 0.12, 0.30,
                 BLUE, shape=MSO_SHAPE.RIGHT_ARROW)

    band(sl, M["as_y"], M["as_h"], T["asis_label"], T["asis_cap"], SLATE,
         FAINT, T["asis"])
    band(sl, M["to_y"], M["to_h"], T["tobe_label"], T["tobe_cap"], DARK,
         PALE, T["tobe"])

    # --- 説明 ---
    ey, eh = M["exp_y"], M["exp_h"]
    bw = (W - 0.24) / 2
    for i, (title, items, accent) in enumerate(
            [(T["issues_title"], T["issues"], RED),
             (T["gains_title"], T["gains"], DARK)]):
        bx = L + i * (bw + 0.24)
        rect(sl, bx, ey, bw, eh, WHITE, RULE, 0.75)
        rect(sl, bx, ey, 0.07, eh, accent)
        say(sl, bx + 0.22, ey + 0.10, bw - 0.40, 0.22, title, 11.5, True, accent)
        iy = ey + 0.36
        for t in items:
            iy += bullet(sl, bx + 0.22, iy, bw - 0.40, t, accent) + 0.06

    # --- So What ---
    sy, sh = M["sw_y"], M["sw_h"]
    rect(sl, L, sy, W, sh, BLUE)
    chip(sl, L + 0.16, sy + 0.07, 1.22, sh - 0.14, "So What", DARK, 11.5)
    say(sl, L + 1.58, sy, W - 1.78, sh, T["sowhat"], 13, True, WHITE,
        anchor=MSO_ANCHOR.MIDDLE)

    fy = M["foot_y"]
    say(sl, L, fy, 7.5, 0.24, T["foot"], 11, False, GREY)
    say(sl, 12.20, fy, 0.72, 0.24, "12", 11, False, GREY, align=PP_ALIGN.RIGHT)

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
