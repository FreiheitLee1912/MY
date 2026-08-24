# MY

このリポジトリは、Web公開用の小さなツールやページを管理するためのスペースです。

## daicel

`daicel` フォルダには、Jira から出力した CSV をブラウザ上で読み込み、ガントチャートとして表示・出力する Web ツールを配置しています。

公開ページ:

```text
https://freiheitlee1912.github.io/MY/daicel/
```

主な機能:

- Jira CSV の読み込み
- ガントチャートのブラウザ表示
- プロジェクト別フィルター
- Type / Parent / グループなしでの表示切替
- PNG 出力
- PPTX 出力

PPTX 出力は、GitHub Pages 上でも動作するようにブラウザ内で生成します。現在の PPTX は、スクリーンショット貼り付けではなく、表ヘッダー・月表示・タスクバー・マイルストーンなどを PowerPoint の図形として生成する形式です。

詳しい使い方は [daicel/README.md](daicel/README.md) を参照してください。

## docs

`docs` フォルダには、運用ルールや管理定義に関するドキュメントを配置しています。

- [JIRA における「案件管理」の定義 — 認識合わせドラフト](docs/jira-project-management-definition.md)
