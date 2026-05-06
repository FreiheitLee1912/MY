# Gantt Chart Generator

Jira から出力した CSV を読み込み、ブラウザ上でガントチャートを作成する Web ツールです。

## 公開ページ

GitHub Pages で以下の URL から利用できます。

```text
https://freiheitlee1912.github.io/MY/daicel/
```

## 主な機能

- Jira CSV ファイルの読み込み
- 複数 CSV の同時読み込み
- ガントチャートのブラウザ表示
- プロジェクト別フィルター
- Type / Parent / グループなしでの表示切替
- 表示期間の調整
- PNG 出力
- PPTX 出力

## 使い方

1. 公開ページを開きます。
2. Jira から出力した CSV ファイルをアップロードします。
3. 必要に応じてプロジェクト、グループ、表示期間を調整します。
4. `Export PNG` または `Export PPTX` を実行します。

## CSV 形式

Jira CSV の以下の列を主に利用します。

- `Key`
- `Type`
- `Summary`
- `Status`
- `Assignee`
- `Start date`
- `Due date`

開始日または終了日のどちらかが存在するタスクをガントチャートに表示します。

## PPTX 出力について

GitHub Pages のような静的ホスティング上でも動くように、PPTX はブラウザ内で生成します。

現在の PPTX 出力は、画面キャプチャを貼り付ける形式ではありません。PowerPoint の図形・テキストとして、以下の要素を生成します。

- 深い青のヘッダー
- 年・月のタイムライン
- 白背景のタスク行
- Type バッジ
- タスクバー
- マイルストーンの diamond 表示
- Today ライン

そのため、PowerPoint 上で各要素を編集しやすい形式になっています。

## ローカル実行

ローカルで Python 版サーバーを使う場合は、Python と `python-pptx` が必要です。

```bash
python server.py
```

起動後、以下をブラウザで開きます。

```text
http://localhost:8090/
```

ローカルサーバー経由の出力ファイルは `output/` フォルダに保存されます。

## ファイル構成

```text
daicel/
  index.html        Web 画面
  styles.css        画面スタイル
  app.js            CSV 読込、ガント描画、PNG/PPTX 出力処理
  sample_tasks.csv  動作確認用サンプル
  vendor/           PPTX 出力用ライブラリ
```
