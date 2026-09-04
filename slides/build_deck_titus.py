# -*- coding: utf-8 -*-
"""JIRA文書管理規程 — Bジ／TITUS 資料と同じスタイルで作成.

Usage:
    python build_deck_titus.py [-o out.pptx]

版面は titus_kit（元資料からの実測値）に従う。
"""
import argparse

from pptx import Presentation
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from titus_kit import (BLACK, WHITE, NAVY, BLUE, PALE, GREY, MIDGREY, RED, GREEN,
                       TAG_DECIDED, TAG_OPEN, NOTE_FILL, WARN_FILL, FONT,
                       L, R, W, BODY_TOP, say, rich, line, cell, secbar, note,
                       frame, blank)

DOCCLASS = "(規程)生産準備部門 JIRA文書管理規程"
DECIDED = ("決定事項", TAG_DECIDED)
OPEN = ("継続議論", TAG_OPEN)


# ------------------------------------------------------------------ 1 表紙 ----
def cover(prs):
    sl = blank(prs)
    cell(sl, 0, 0, 13.333, 0.86, "", fill=GREY, border=None)
    say(sl, L, 0.24, 9.0, 0.38, "生産準備部門　運用規程", 18, True, BLACK)
    cell(sl, L, 1.30, W, 0.86, "JIRA文書管理規程", 30, True, WHITE, BLUE, None)
    say(sl, L, 2.44, 11.5, 0.62,
        "プロジェクト管理に関する記録をタスク単位で保管し、\n"
        "社外との合意事項を後日確認できる状態を維持することを目的とする。",
        16, False, BLACK, space=1.3)
    meta = [("文書番号", "－　※要記入", True),
            ("制定日", "2026年9月4日", False),
            ("版数", "第1.0版", False),
            ("管理部署", "－　※要記入", True),
            ("適用対象", "生準メンバ（今後、関連部署へ拡大）", False)]
    y = 3.55
    for k, v, tbd in meta:
        cell(sl, L, y, 2.40, 0.47, k, 13, True, BLACK, PALE, BLACK, 0.75)
        cell(sl, L + 2.40, y, 6.40, 0.47, v, 13, False,
             RED if tbd else BLACK, WHITE, BLACK, 0.75, align=PP_ALIGN.LEFT, pad=0.14)
        y += 0.47
    note(sl, L, y + 0.30, W, 0.56, "未決事項",
         "文書番号の採番、管理部署の確定、関連部署への公開時期（別途通知）の3点が未確定。")
    say(sl, 7.58, 7.24, 4.81, 0.20,
        "Copyright © 2021 Accenture. All rights reserved. Highly Confidential",
        9, False, BLACK, align=PP_ALIGN.RIGHT)
    say(sl, 12.70, 7.22, 0.30, 0.20, "1", 12, False, BLACK)


# ---------------------------------------------------------------- 2 全体像 ----
def overview(prs, n):
    sl = blank(prs)
    frame(sl, DOCCLASS, "本規程の全体像と条文の構成",
          "本規程は、記録の保管先と登録可否を定める第1条〜第7条で構成する。\n"
          "第3条・第6条が登録可否の判断軸、第5条が保管先の切り分けにあたる。",
          "（第1条〜第7条）", (DECIDED,), n)
    y = BODY_TOP
    heads = [("条", 1.10), ("表題", 3.20), ("要点", 7.70)]
    x = L
    for t, w in heads:
        secbar(sl, x, y, w, t, 0.36, 14)
        x += w
    rows = [("第1条", "目的", "記録は個人のメールボックスではなく、当該タスクのチケットに保管する"),
            ("第2条", "適用範囲および公開範囲", "生準メンバのみ → 関連部署へ段階的に公開する"),
            ("第3条", "保管対象", "社外とのやり取りを最優先で登録し、図面・個人情報は登録しない"),
            ("第4条", "社外コミュニケーションの記録", "メールは原本のまま保管し、口頭決定は確認メールを送って残す"),
            ("第5条", "JIRAおよびWinchillの役割分担", "図面・CADと版管理はWinchillに寄せ、JIRAには参照リンクのみ"),
            ("第6条", "情報区分と取扱い", "機密以上はJIRAに登録せず、参照リンクのみを記載する"),
            ("第7条", "登録前チェックリスト", "登録前に5点を確認してから登録する")]
    ry = y + 0.36
    for no, title, pt in rows:
        cell(sl, L, ry, 1.10, 0.54, no, 13, True, BLACK, PALE, BLACK, 0.75)
        cell(sl, L + 1.10, ry, 3.20, 0.54, title, 13, False, BLACK, WHITE, BLACK,
             0.75, align=PP_ALIGN.LEFT, pad=0.14)
        cell(sl, L + 4.30, ry, 7.70, 0.54, pt, 13, False, BLACK, WHITE, BLACK,
             0.75, align=PP_ALIGN.LEFT, pad=0.14)
        ry += 0.54
    note(sl, L, ry + 0.24, W, 0.62, "判断の軸",
         "「何を登録するか」＝第3条・第6条　／　「どこに保管するか」＝第5条　／　"
         "「どう残すか」＝第4条")


# ------------------------------------------------------------- 3 第1条 目的 ----
def purpose(prs, n):
    sl = blank(prs)
    frame(sl, DOCCLASS, "[第1条] 目的",
          "JIRAは進捗管理の手段にとどまらず、「誰が・いつ・何を決定したか」を後から確認できる\n"
          "記録の保管先として運用する。",
          "（第1条）", (DECIDED,), n)
    y = BODY_TOP
    secbar(sl, L, y, W, "記録を残す目的")
    goals = [("1", "認識の相違を防止する", "社外との協議・合意の経緯を\n確実に保全する。"),
             ("2", "追跡できる状態を保つ", "担当者の交代または不在時においても、\n第三者が経緯を追跡できる。"),
             ("3", "事実関係を再構成する", "問題発生時に、事実関係を速やかに\n再構成できるようにする。")]
    cw = (W - 0.40) / 3
    for i, (no, t, d) in enumerate(goals):
        x = L + i * (cw + 0.20)
        cell(sl, x, y + 0.60, cw, 0.42, f"({no})　{t}", 14, True, WHITE, NAVY, BLACK, 0.75)
        cell(sl, x, y + 1.02, cw, 1.62, d, 13, False, BLACK, WHITE, BLACK, 0.75,
             space=1.3)
    ny = y + 2.90
    secbar(sl, L, ny, W, "原則")
    cell(sl, L, ny + 0.39, W, 1.55,
         "記録は個人のメールボックスではなく、当該タスクのチケットに保管する。\n"
         "個人の受信箱のみに存在する情報は、組織の記録として扱わない。",
         15, True, BLACK, NOTE_FILL, BLACK, 0.75, space=1.4)


# --------------------------------------------------------- 4 第2条 公開範囲 ----
def scope(prs, n):
    sl = blank(prs)
    frame(sl, DOCCLASS, "[第2条] 適用範囲および公開範囲",
          "関連部署への公開を前提とするため、登録内容は当該範囲で閲覧されることを想定して\n"
          "記載する。",
          "（第2条）", (DECIDED, OPEN), n)
    y = BODY_TOP
    secbar(sl, L, y, W, "表1　JIRAの公開範囲")
    cols = [("区分", 1.60), ("対象", 4.40), ("備考", 6.00)]
    x = L
    for t, w in cols:
        cell(sl, x, y + 0.39, w, 0.36, t, 13, False, WHITE, NAVY, BLACK, 0.75)
        x += w
    rows = [("現行", "生準メンバのみ", "試行運用期間"),
            ("今後", "関連部署に公開する", "公開時期は別途通知する　※時期未定")]
    ry = y + 0.75
    for k, tgt, rem in rows:
        cell(sl, L, ry, 1.60, 0.47, k, 13, True, BLACK, PALE, BLACK, 0.75)
        cell(sl, L + 1.60, ry, 4.40, 0.47, tgt, 13, False, BLACK, WHITE, BLACK,
             0.75, align=PP_ALIGN.LEFT, pad=0.12)
        cell(sl, L + 6.00, ry, 6.00, 0.47, rem, 13, False, BLACK, WHITE, BLACK,
             0.75, align=PP_ALIGN.LEFT, pad=0.12)
        ry += 0.47
    ny = ry + 0.34
    secbar(sl, L, ny, W, "公開拡大に備えて徹底すること")
    prep = [("社外情報は原本で残す", "要約ではなく原本を保管し、\n閲覧者が経緯を追える状態にする。"),
            ("個人情報を含めない", "業務上不要な連絡先・私的な内容は\n登録前に取り除く。"),
            ("機密はリンクのみ", "機密以上はWinchill等に保管し、\nJIRAには参照先だけを記載する。")]
    cw = (W - 0.40) / 3
    for i, (t, d) in enumerate(prep):
        x = L + i * (cw + 0.20)
        cell(sl, x, ny + 0.51, cw, 0.40, t, 13, True, BLACK, PALE, BLACK, 0.75)
        cell(sl, x, ny + 0.91, cw, 1.15, d, 12.5, False, BLACK, WHITE, BLACK,
             0.75, space=1.3)
    note(sl, L, ny + 2.30, W, 0.62, "留意点",
         "関連部署に閲覧されて差し支えない表現か、登録の都度確認する（第7条のチェック項目に対応）。")


# --------------------------------------------------------- 5 第3条 保管対象 ----
def targets(prs, n):
    sl = blank(prs)
    frame(sl, DOCCLASS, "[第3条] 保管対象",
          "プロジェクト管理に関連する以下の情報を、該当するチケットに保管する。",
          "（第3条）", (DECIDED,), n)
    y = BODY_TOP
    cw = 5.90
    secbar(sl, L, y, cw, "登録するもの")
    rows = [("社外とのやり取り　※最重要", "顧客・サプライヤとの依頼、回答、条件調整、合意"),
            ("決定事項", "方針決定および重要な決議内容"),
            ("進捗報告", "定期報告、マイルストーンの達成状況"),
            ("要件確認", "仕様決定、変更依頼の承認"),
            ("議事録", "打合せの記録および議論の経過"),
            ("ステータス", "課題・リスク・対応状況")]
    ry = y + 0.39
    for i, (k, v) in enumerate(rows):
        cell(sl, L, ry, 2.55, 0.68, k, 12.5, True, RED if i == 0 else BLACK,
             PALE, BLACK, 0.75, align=PP_ALIGN.LEFT, pad=0.12, space=1.2)
        cell(sl, L + 2.55, ry, cw - 2.55, 0.68, v, 12, False, BLACK, WHITE, BLACK,
             0.75, align=PP_ALIGN.LEFT, pad=0.12, space=1.2)
        ry += 0.68
    x2 = L + cw + 0.20
    cw2 = W - cw - 0.20
    secbar(sl, x2, y, cw2, "登録しないもの")
    nos = [("図面・設計図・CADファイル", "Winchillにて保管する（第5条）"),
           ("業務上不要な個人情報", "個人携帯番号、自宅住所、私的な内容")]
    ry2 = y + 0.39
    for k, v in nos:
        cell(sl, x2, ry2, cw2, 0.42, k, 13, True, RED, PALE, BLACK, 0.75)
        cell(sl, x2, ry2 + 0.42, cw2, 0.47, v, 12.5, False, BLACK, WHITE, BLACK,
             0.75, align=PP_ALIGN.LEFT, pad=0.14)
        ry2 += 1.05
    note(sl, x2, ry2 + 0.14, cw2, 1.80, "判断に迷う場合",
         "登録前に第7条のチェックリストで確認する。機密区分に該当するものは、"
         "JIRAには参照リンクのみを記載する。", fill=WARN_FILL, size=12.5)


# ------------------------------------------ 6 第4条 社外コミュニケーション ----
def external(prs, n):
    sl = blank(prs)
    frame(sl, DOCCLASS, "[第4条] 社外コミュニケーションの記録",
          "社外とのメールは原本のまま保管する。要約のみの記載は認めない。",
          "（第4条）", (DECIDED,), n)
    y = BODY_TOP
    secbar(sl, L, y, W, "記録の原則")
    rules = [("(1)", "原本のまま保管する",
              "社外とのメールは、該当チケットに原本のまま保管する。要約のみの記載は認めない。"),
             ("(2)", "依頼と回答を対で保管する",
              "依頼と回答は対で保管し、合意内容を読み取れる状態とする。"),
             ("(3)", "訂正は追記で残す",
              "訂正が生じた場合、元の記録は削除せず、コメントにより追記する。")]
    ry = y + 0.39
    for no, t, d in rules:
        cell(sl, L, ry, 0.70, 0.78, no, 13, True, BLACK, PALE, BLACK, 0.75)
        cell(sl, L + 0.70, ry, 3.60, 0.78, t, 13, True, BLACK, WHITE, BLACK,
             0.75, align=PP_ALIGN.LEFT, pad=0.14)
        cell(sl, L + 4.30, ry, W - 4.30, 0.78, d, 12.5, False, BLACK, WHITE, BLACK,
             0.75, align=PP_ALIGN.LEFT, pad=0.14)
        ry += 0.78
    ny = ry + 0.34
    secbar(sl, L, ny, W, "電話・口頭で決定した場合")
    cell(sl, L, ny + 0.39, W, 1.40,
         "内容の確認メールを相手方へ送付し、当該メールをチケットに保管する。\n"
         "本項の徹底が、認識の相違に対する最も有効な予防措置となる。",
         15, True, BLACK, WARN_FILL, BLACK, 0.75, space=1.4)


# --------------------------------------------------- 7 第5条 JIRA/Winchill ----
def systems(prs, n):
    sl = blank(prs)
    frame(sl, DOCCLASS, "[第5条] JIRAおよびWinchillの役割分担",
          "JIRAでは版の履歴を都度確認できないため、バージョン管理の観点からWinchillにて\n"
          "一元管理する。",
          "（第5条）", (DECIDED,), n)
    y = BODY_TOP
    secbar(sl, L, y, W, "表2　システム別の保管対象")
    cols = [("システム", 1.70), ("保管対象", 4.30), ("用途", 3.30), ("アクセス", 2.70)]
    x = L
    for t, w in cols:
        cell(sl, x, y + 0.39, w, 0.36, t, 13, False, WHITE, NAVY, BLACK, 0.75)
        x += w
    rows = [("JIRA", "決定事項／進捗報告／議事録／\nステータス／要件確認",
             "プロジェクト管理\nコミュニケーション記録", "生準メンバ\n（今後、関連部署）"),
            ("Winchill", "図面・設計図／CADファイル／技術仕様書／\n製造情報／その他機密資料",
             "機密情報管理\n設計・技術情報の版管理", "権限者のみ\n（別途制御）")]
    ry = y + 0.75
    for name, keep, use, acc in rows:
        cell(sl, L, ry, 1.70, 1.35, name, 14, True, BLACK, PALE, BLACK, 0.75)
        cell(sl, L + 1.70, ry, 4.30, 1.35, keep, 12, False, BLACK, WHITE, BLACK,
             0.75, align=PP_ALIGN.LEFT, pad=0.14, space=1.35)
        cell(sl, L + 6.00, ry, 3.30, 1.35, use, 12, False, BLACK, WHITE, BLACK,
             0.75, align=PP_ALIGN.LEFT, pad=0.14, space=1.35)
        cell(sl, L + 9.30, ry, 2.70, 1.35, acc, 12, False, BLACK, WHITE, BLACK,
             0.75, align=PP_ALIGN.LEFT, pad=0.14, space=1.35)
        ry += 1.35
    note(sl, L, ry + 0.32, W, 1.10, "図面・CADファイルの取扱い",
         "JIRAには登録せず、Winchillにて保管する。JIRAには参照リンクまたは"
         "Winchill No.のみを記載する。JIRAでは版の履歴を都度確認できないため、"
         "バージョン管理はWinchillに一元化する。", size=13)


# ------------------------------------------------------- 8 第6条 情報区分 ----
def classes(prs, n):
    sl = blank(prs)
    frame(sl, DOCCLASS, "[第6条] 情報区分と取扱い",
          "機密区分の情報を参照する必要がある場合は、JIRAには外部リンクのみを記載する。\n"
          "添付ファイルとしては登録しない。",
          "（第6条）", (DECIDED,), n)
    y = BODY_TOP
    secbar(sl, L, y, W, "表3　情報区分別の登録可否")
    cols = [("情報区分", 3.40), ("JIRAへの登録", 2.60), ("保管先", 6.00)]
    x = L
    for t, w in cols:
        cell(sl, x, y + 0.39, w, 0.36, t, 13, False, WHITE, NAVY, BLACK, 0.75)
        x += w
    rows = [("一般（Public）", "可", GREEN, "JIRA"),
            ("社内（Internal）", "可", GREEN, "JIRA（アクセス権限を設定する）"),
            ("機密（Confidential）", "不可", RED, "Winchill／専用ストレージ　※参照リンクのみ"),
            ("極秘（Secret）", "不可", RED, "限定者のみアクセス可能な場所")]
    ry = y + 0.75
    for k, ok, col, dest in rows:
        cell(sl, L, ry, 3.40, 0.70, k, 13, True, BLACK, PALE, BLACK, 0.75,
             align=PP_ALIGN.LEFT, pad=0.16)
        cell(sl, L + 3.40, ry, 2.60, 0.70, ok, 14, True, col, WHITE, BLACK, 0.75)
        cell(sl, L + 6.00, ry, 6.00, 0.70, dest, 12.5, False, BLACK, WHITE, BLACK,
             0.75, align=PP_ALIGN.LEFT, pad=0.16)
        ry += 0.70
    note(sl, L, ry + 0.34, W, 0.95, "運用",
         "機密区分の情報を参照する必要がある場合は、JIRAには外部リンクのみを記載し、"
         "添付ファイルとしては登録しない。")


# ------------------------------------------------- 9 第7条 チェックリスト ----
def checklist(prs, n):
    sl = blank(prs)
    frame(sl, DOCCLASS, "[第7条] 登録前チェックリスト",
          "JIRAに文書を登録する前に、以下の5点を確認する。",
          "（第7条）", (DECIDED,), n)
    y = BODY_TOP
    secbar(sl, L, y, W, "登録前の確認事項")
    items = [("図面・CADファイルではないか", "Winchillで保管すべきものではないか"),
             ("顧客の個人情報・連絡先が含まれていないか", "個人携帯番号・自宅住所・私的な内容"),
             ("必要なアクセス権限が設定されているか", "社内区分は権限設定のうえ登録する"),
             ("ファイルサイズは50MB以下か", "超過する場合は分割または別保管を検討する"),
             ("関連部署に閲覧されても差し支えない内容か", "公開範囲の拡大を前提に判断する")]
    ry = y + 0.39
    for i, (q, rem) in enumerate(items, 1):
        cell(sl, L, ry, 0.60, 0.70, "☐", 15, False, BLACK, WHITE, BLACK, 0.75)
        cell(sl, L + 0.60, ry, 0.60, 0.70, str(i), 13, True, BLACK, PALE, BLACK, 0.75)
        cell(sl, L + 1.20, ry, 6.20, 0.70, q, 13, True, BLACK, WHITE, BLACK, 0.75,
             align=PP_ALIGN.LEFT, pad=0.16)
        cell(sl, L + 7.40, ry, W - 7.40, 0.70, rem, 12.5, False, BLACK, WHITE,
             BLACK, 0.75, align=PP_ALIGN.LEFT, pad=0.16)
        ry += 0.70
    note(sl, L, ry + 0.30, W, 0.80, "図面・設計図の場合",
         "Winchillに登録のうえ、JIRAには参照リンクまたはWinchill No.を記載する。",
         fill=WARN_FILL)


# ---------------------------------------------------------------- 10 まとめ ----
def closing(prs, n):
    sl = blank(prs)
    cell(sl, 0, 0, 13.333, 0.86, "", fill=GREY, border=None)
    say(sl, L, 0.24, 9.0, 0.38, "まとめ", 18, True, BLACK)
    say(sl, 12.70, 7.22, 0.30, 0.20, str(n), 12, False, BLACK)
    say(sl, 7.58, 7.24, 4.81, 0.20,
        "Copyright © 2021 Accenture. All rights reserved. Highly Confidential",
        9, False, BLACK, align=PP_ALIGN.RIGHT)
    cell(sl, L, 1.22, W, 0.80,
         "記録の所在を一つに定めることが、認識の相違と属人化を防ぐ",
         24, True, WHITE, BLUE, None)
    pts = [("残す場所", "個人のメールボックスではなく、\n当該タスクのチケットに保管する。"),
           ("残し方", "社外メールは原本のまま。口頭決定は\n確認メールを送り、それを保管する。"),
           ("分けるもの", "図面・CADと機密以上はWinchillへ。\nJIRAには参照リンクのみを記載する。")]
    cw = (W - 0.40) / 3
    for i, (t, d) in enumerate(pts):
        x = L + i * (cw + 0.20)
        cell(sl, x, 2.42, cw, 0.42, t, 14, True, WHITE, NAVY, BLACK, 0.75)
        cell(sl, x, 2.84, cw, 1.10, d, 13, False, BLACK, WHITE, BLACK, 0.75, space=1.35)
    secbar(sl, L, 4.34, W, "未決事項")
    opens = [("(1)", "文書番号の採番"), ("(2)", "管理部署の確定"),
             ("(3)", "関連部署への公開時期（別途通知）")]
    ry = 4.73
    for no, t in opens:
        cell(sl, L, ry, 0.70, 0.47, no, 13, True, BLACK, TAG_OPEN, BLACK, 0.75)
        cell(sl, L + 0.70, ry, W - 0.70, 0.47, t, 13, False, BLACK, WHITE, BLACK,
             0.75, align=PP_ALIGN.LEFT, pad=0.14)
        ry += 0.47
    say(sl, L, ry + 0.34, W, 0.30,
        "附則　本規程は制定日より実施する。改定は管理部署の承認を経て行う。",
        13, False, BLACK)


def build(output):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    cover(prs)
    for i, fn in enumerate([overview, purpose, scope, targets, external,
                            systems, classes, checklist], start=2):
        fn(prs, i)
    closing(prs, 10)
    prs.save(output)
    print("saved", output, "-", len(prs.slides._sldIdLst), "slides")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("-o", "--output", default="jira_document_policy_titus.pptx")
    build(ap.parse_args().output)
