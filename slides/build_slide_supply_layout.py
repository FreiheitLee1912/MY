# -*- coding: utf-8 -*-
"""現状と課題｜供給レイアウト — 供給案比較の自動化。

Usage:
    python build_slide_supply_layout.py [--lang ja|en] [-o out.pptx]

上段に検討プロセス（5工程、自動化対象の「供給案の作成・比較」を強調）、
下段に現状の課題と目指す姿を1行ずつ対で並べる。

日英で版面は同じ。折り返しに効く寸法だけ言語ごとに持ち替える（METRICS）。
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

BODY = "Meiryo"       # build() が言語に応じて差し替える
DISPLAY = "Meiryo"
EM11 = 0.158          # 11pt 本文の折り返し見積り（安全側・言語で持ち替え）
LH11 = 0.20           # 11pt / 行送り1.25 の1行


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
         shape=MSO_SHAPE.RECTANGLE):
    sh = rect(sl, x, y, w, h, fill, outline, ow, shape)
    tf = sh.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = text
    style(r, size, True, fg)
    return sh


def lines(text, inner_w):
    return max(1, math.ceil(len(text) / max(6, int(inner_w / EM11))))


FOCUS = 2                      # 自動化対象の工程

JA = dict(
    eyebrow="② 現状と課題 ｜ 供給レイアウト",
    headline="供給レイアウト候補の比較を自動化し、共通条件で複数案を迅速に評価して"
             "選定を支援する",
    sub="最新の需給情報を活用し、条件変更に伴う再計算・資料更新の手作業を削減することで、"
        "比較検討にかかる時間を短縮する。",
    proc="検討プロセス",
    steps=["需給状況の確認", "課題の特定", "供給案の作成・比較",
           "供給レイアウトの更新", "報告"],
    badge="比較作業の自動化対象",
    issues_head="現状の課題", goals_head="目指す姿",
    rows=[
        ("Excelで供給案を個別に作成・比較しており、情報の整理と比較に時間がかかる",
         "共通のデータと比較条件で、複数の供給案を評価できる"),
        ("条件変更のたびに、再計算と比較資料の更新が必要になる",
         "条件変更を比較結果に反映し、更新作業を削減できる"),
        ("手入力・手計算に伴うミスのリスクがある",
         "入力・計算の手作業を減らし、ミスのリスクを低減できる"),
        ("作業負荷により、比較できるパターンが限られる",
         "複数のシナリオを比較し、各案の差分を確認できる"),
    ],
    sowhat="条件変更のたびの再計算と資料更新をなくせば、同じ条件で複数案を並べて比べられる",
    foot="現状と課題｜供給レイアウト",
    notes="供給レイアウトの現状と課題。狙いは供給レイアウト候補の比較を自動化し、共通条件で"
          "複数案を迅速に評価して供給案の選定を支援すること。最新の需給情報を活用し、条件変更に"
          "伴う再計算・資料更新の手作業を削減することで、比較検討にかかる時間を短縮する。"
          "検討プロセスは需給状況の確認、課題の特定、供給案の作成・比較、供給レイアウトの更新、"
          "報告の5工程で、このうち「供給案の作成・比較」が比較作業の自動化対象。現状はExcelで"
          "供給案を個別に作成・比較しており情報の整理と比較に時間がかかる、条件変更のたびに"
          "再計算と比較資料の更新が必要、手入力・手計算のミスのリスクがある、作業負荷により"
          "比較できるパターンが限られる、の4点。目指す姿は共通のデータと比較条件で複数案を"
          "評価でき、条件変更を比較結果に反映して更新作業を削減し、手作業を減らしてミスの"
          "リスクを下げ、複数シナリオの差分を確認できる状態。",
)

EN = dict(
    eyebrow="2. CURRENT STATE & ISSUES ｜ SUPPLY ALLOCATION",
    headline="Compare supply-allocation options automatically, on one common "
             "set of conditions",
    sub="Use the latest supply-demand data to cut manual recalculation and deck "
        "updates, shortening the time comparison takes.",
    proc="Review process",
    steps=["Check supply & demand", "Identify the issue", "Build & compare "
           "options", "Update the allocation", "Report"],
    badge="Target of automation",
    issues_head="Issues today", goals_head="Target state",
    rows=[
        ("Supply options are built and compared one by one in Excel, so "
         "organising and comparing them takes time",
         "Evaluate several options on one data set and common conditions"),
        ("Every change of conditions requires a recalculation and an updated "
         "comparison deck",
         "Condition changes flow into the result, cutting update work"),
        ("Manual entry and manual calculation carry a risk of error",
         "Less manual entry and calculation, so less risk of error"),
        ("The workload limits how many patterns can be compared",
         "Compare several scenarios and see each option's difference"),
    ],
    sowhat="Remove the rework each condition change triggers and options can be "
           "compared on the same conditions",
    foot="Current state & issues | Supply allocation",
    notes="Current state and issues for supply allocation - which site "
          "supplies which demand. The aim is to automate the comparison of "
          "allocation options so that several can be evaluated quickly on common "
          "conditions, supporting the choice of a supply option. Using the latest supply-demand data cuts the manual "
          "recalculation and deck updates that every change of conditions "
          "triggers, shortening the time comparison takes. The review process "
          "runs check supply and demand, identify the issue, build and compare "
          "options, update the allocation, report; of these, building and "
          "comparing options is the target of automation. Today options are "
          "built and compared one by one in Excel so organising and comparing "
          "them takes time, every change of conditions requires a recalculation "
          "and an updated deck, manual entry and calculation carry a risk of "
          "error, and the workload limits how many patterns can be compared. The "
          "target state evaluates several options on one data set and common "
          "conditions, flows condition changes straight into the result, reduces "
          "manual work and the risk of error, and compares several scenarios to "
          "show each option's difference.",
)

METRICS = dict(
    ja=dict(em=0.158, head_size=20, chev_pad=0.30),
    en=dict(em=0.085, head_size=18, chev_pad=0.24),
)

BAND_FILL = C("DCEEF9")        # 自動化対象を括る帯

# --- 版面 ---
# ACNの矢羽根に倣い、両端の工程を大きいシェブロンで囲い、
# 自動化対象の工程だけハッチの帯で括る。
PROC_Y = 1.56
CHEV_Y, CHEV_H = 2.12, 0.56
OUTER_H = 0.80
PITCH, CHEV_W = 2.46, 2.56
BAND_Y, BAND_H = 1.86, 0.92

TBL_Y, HEAD_H, ROW_H = 3.00, 0.40, 0.72
CW, AW = 5.85, 0.80            # 課題／矢印／目指す姿の列幅
AX = L + CW
RX = AX + AW

SW_Y, SW_H = 6.66, 0.52
FOOT_Y = 7.24


def build(output, lang):
    global BODY, DISPLAY, EM11
    T = JA if lang == "ja" else EN
    M = METRICS[lang]
    BODY = "Meiryo" if lang == "ja" else "Arial"
    DISPLAY = "Meiryo" if lang == "ja" else "Arial Black"
    EM11 = M["em"]

    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    sl = prs.slides.add_slide(prs.slide_layouts[6])

    # --- ヘッダー ---
    say(sl, L, 0.36, 10.5, 0.26, T["eyebrow"], 11.5, True, BLUE, spacing=1.0)
    say(sl, L, 0.62, W, 0.56, T["headline"], M["head_size"], True, INK,
        anchor=MSO_ANCHOR.MIDDLE, face=DISPLAY)
    say(sl, L, 1.20, 12.3, 0.26, T["sub"], 12.5, False, GREY)

    # --- 検討プロセス（ACNの矢羽根に倣う） ---
    say(sl, L, PROC_Y, 3.0, 0.22, T["proc"], 11.5, True, DARK)

    last = len(T["steps"]) - 1
    fx = L + FOCUS * PITCH
    bx, bw = fx - 0.20, CHEV_W + 0.40
    rect(sl, bx, BAND_Y, bw, BAND_H, BAND_FILL, BLUE, 0.75)
    say(sl, bx + 0.06, BAND_Y + 0.05, bw - 0.12, 0.24, T["badge"], 11,
        True, DARK, align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    for i, name in enumerate(T["steps"]):
        x = L + i * PITCH
        focus = i == FOCUS
        outer = i in (0, last)
        h = OUTER_H if outer else CHEV_H
        y = CHEV_Y + (CHEV_H - h) / 2
        if outer:
            fill, line, lw, fg = DARK, None, 0.75, WHITE
        elif focus:
            fill, line, lw, fg = BLUE, DARK, 1.25, WHITE
        else:
            fill, line, lw, fg = WHITE, RULE, 0.75, INK
        ch = rect(sl, x, y, CHEV_W, h, fill, line, lw, MSO_SHAPE.CHEVRON)
        # 矢じりを浅くしないと11pt の工程名が入らない
        ch.adjustments[0] = 0.26
        pad = M["chev_pad"] + (0.08 if outer else 0)
        say(sl, x + pad, y, CHEV_W - pad * 2, h, name, 11.5 if outer else 11,
            True if outer or focus else False, fg, align=PP_ALIGN.CENTER,
            anchor=MSO_ANCHOR.MIDDLE, space=1.1)

    # --- 現状の課題 ／ 目指す姿 ---
    chip(sl, L, TBL_Y, CW, HEAD_H, T["issues_head"], RED, 12.5)
    chip(sl, RX, TBL_Y, CW, HEAD_H, T["goals_head"], DARK, 12.5)

    y = TBL_Y + HEAD_H + 0.06
    for i, (issue, goal) in enumerate(T["rows"], 1):
        rect(sl, L, y, CW, ROW_H, WHITE, RULE, 0.75)
        rect(sl, RX, y, CW, ROW_H, PALE, RULE, 0.75)
        chip(sl, L + 0.16, y + (ROW_H - 0.28) / 2, 0.30, 0.28, str(i), RED, 11)
        ih = lines(issue, CW - 0.80) * LH11
        say(sl, L + 0.58, y + (ROW_H - ih) / 2, CW - 0.80, ih, issue, 11, False,
            INK, space=1.25)
        gh = lines(goal, CW - 0.44) * LH11
        say(sl, RX + 0.22, y + (ROW_H - gh) / 2, CW - 0.44, gh, goal, 11, True,
            DARK, space=1.25)
        rect(sl, AX + (AW - 0.42) / 2, y + (ROW_H - 0.22) / 2, 0.42, 0.22, BLUE,
             shape=MSO_SHAPE.RIGHT_ARROW)
        y += ROW_H + 0.06

    # --- So What ---
    rect(sl, L, SW_Y, W, SW_H, BLUE)
    chip(sl, L + 0.16, SW_Y + 0.07, 1.22, SW_H - 0.14, "So What", DARK, 11.5)
    say(sl, L + 1.58, SW_Y, W - 1.78, SW_H, T["sowhat"], 13, True, WHITE,
        anchor=MSO_ANCHOR.MIDDLE)

    say(sl, L, FOOT_Y, 7.5, 0.24, T["foot"], 11, False, GREY)
    say(sl, 12.20, FOOT_Y, 0.72, 0.24, "13", 11, False, GREY,
        align=PP_ALIGN.RIGHT)

    sl.notes_slide.notes_text_frame.text = T["notes"]
    prs.save(output)
    print("saved", output)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--lang", choices=("ja", "en"), default="ja")
    ap.add_argument("-o", "--output")
    a = ap.parse_args()
    build(a.output or "supply_layout_daicel%s.pptx" %
          ("" if a.lang == "ja" else "_en"), a.lang)
