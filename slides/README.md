# JIRA登録タイミング スライド

DAICEL版「JIRA登録タイミング」スライドを組み直したものです。
レイアウトは共通で、配色だけを 2 ブランド分切り替えられます。

## 生成方法

```bash
pip install python-pptx

# DAICEL配色（既定）
python build_slide.py --brand daicel

# Accenture配色（スターターパックの .potx が必要）
python build_slide.py --brand acn --template <Acc_PPT_IMP_StarterPack....potx>
```

## レイアウトの考え方

版面は Accenture 2020 スターターパックから採寸したものを両ブランドで共用しています。

- 左右マージン 0.42"、テキスト幅 12.50"
- 見出しは Arial Black 27pt を左上にベタ置き（**下線・上罫線は引かない**）
- ページ下部を「So What」バンドで締める
- 本文 Arial／和文 Meiryo

`--brand acn` はスターターパックの `Long Headline Only` レイアウト上に組むため、
`>` マークとフッターはマスターから継承されます。
`--brand daicel` はマスターを使いません。Accenture の `>` マークは他社の商標であり、
紫は青系配色と衝突するためです。ロゴはフッターの空き（左下）に差し込んでください。

## 配色

| | acn | daicel |
|---|---|---|
| 基調 | `A100FF` Accenture Purple | `0068B7` DAICEL Blue |
| 濃色 | `460073` | `1F4E79` |
| 淡色/地色 | `F4E9FF` / `E6E6DC` | `E8F1FA` / `E7E7E7` |
| SOP後 | パープル `A100FF`／`7500C0` | グリーン `4C9A2A`／`3D7C21` |
| 未決事項 | ウォームグレー `E6E6DC` | ペールアンバー `FBF4E0` |

ライフサイクルの色分けは「対象外／SOP前／SOP後」という元スライドの意味を保持しています。
`daicel` では元スライド同様に SOP 後をグリーンで表し、`acn` ではパープル・ランプに寄せています。

## 構成

- 見出し（結論）＋サブ見出し
- 登録判断バンド … INF／イニシそれぞれのトリガー定義
- プロジェクトライフサイクル … 9フェーズのシェブロン、SOPを破線で区切り
- スイムレーン … INF／イニシ別に「個別管理 → JIRAで管理 → PCR都度起票」
- 未決事項
- So What（結論バンド）

## 注意

生成した `.pptx` はリポジトリにコミットしていません（`.gitignore` で除外）。
`--brand acn` の生成物は Accenture のマスター・テーマ・ロゴを含むため、
配布・公開時はテンプレートの利用条件を確認してください。
