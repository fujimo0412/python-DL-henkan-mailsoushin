# Kaggle データ自動取得・メール送信ツール

## 概要

KaggleのAPIでデータを取得し、Outlookで自動メール送信するまでの一連の流れをPythonで自動化したツールです。

外部APIからのデータ取得・加工・連携という実務でよく使われるパターンの習得を目的として作成しました。実用上のためKaggleとOutlookを使用しています。  
バッチファイル（`実行ボタンrun.bat`）からワンクリックで実行できます。

---

## 処理の流れ

1. Kaggle APIで指定データセットをZIPダウンロード
2. ZIPを解凍し、ファイル名に日付を付与（例：`superstore_20260610.csv`）
3. 元ZIPファイルを削除
4. Outlookで自分宛にファイルを添付してメール送信

---

## 環境変数の設定

以下の環境変数を事前に設定してください。

| 変数名 | 内容 |
|---|---|
| `KAGGLE_USERNAME` | Kaggle のユーザー名 |
| `KAGGLE_KEY` | Kaggle API キー |
| `DOWNLOAD_DIR` | ダウンロード先フォルダのパス（例：`C:\Users\ユーザー名\Documents`）。配下に`KaggleData`フォルダが自動作成されます |
| `MAIL_TO` | 送信先メールアドレス |

Kaggle APIキーは [https://www.kaggle.com/settings](https://www.kaggle.com/settings) にログイン後、右上のアイコンをクリックして「Your API tokens」＞「Generate New Token」から取得できます。

---

## 使用データセット

[Superstore Dataset](https://www.kaggle.com/datasets/vivek468/superstore-dataset-final)（vivek468/superstore-dataset-final）

---

## 実行方法

`実行ボタンrun.bat` をダブルクリック

---

## 必要なライブラリ

```
pip install kaggle pywin32
```

---

## 動作環境

- Windows（Outlookがインストールされていること）
- Python 3.x
- Outlookアカウントが設定済みであること

---

## 参考

- [バッチファイルでPythonを実行する方法](https://qiita.com/WenChunPan/items/ca46b02e4a8effa33a91)
- [GitHub - WenChunPan/Qiita-batchfile-run-python](https://github.com/WenChunPan/Qiita-batchfile-run-python)
