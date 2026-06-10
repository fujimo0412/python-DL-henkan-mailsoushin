# Kaggle データ自動取得・メール送信ツール

## 概要

Kaggleからデータセットをダウンロードし、日付付きファイル名に整理してOutlookで自分宛にメール送信するPython自動化ツールです。  
バッチファイル（`実行ボタンrun.bat`）からワンクリックで実行できます。

---

## 処理の流れ

1. Kaggle APIで指定データセットをZIPダウンロード
2. ZIPを解凍し、ファイル名に日付を付与（例：`superstore_20260610.csv`）
3. ZIPファイルを削除
4. Outlookで自分宛にファイルを添付してメール送信

---

## 環境変数の設定

実行前に以下の環境変数を設定してください。

| 変数名 | 内容 |
|---|---|
| `KAGGLE_USERNAME` | Kaggle のユーザー名 |
| `KAGGLE_KEY` | Kaggle API キー |
| `MY_EMAIL` | 送信先メールアドレス |

### 設定手順（Windows）

1. スタートメニューで「環境変数」と検索
2. 「システム環境変数の編集」を開く
3. 「環境変数」ボタンをクリック
4. 「新規」で上記3つを追加

Kaggle APIキーは [https://www.kaggle.com/settings](https://www.kaggle.com/settings) の「API」セクションから `kaggle.json` をダウンロードして確認できます。

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
