import os
import zipfile
import datetime
import subprocess
import glob
import win32com.client  # Outlook 送信用

download_dir = r"C:\Users\user\Documents"
dataset = "vivek468/superstore-dataset-final"

# 保存先フォルダ（固定）
output_dir = os.path.join(download_dir, "KaggleData")
os.makedirs(output_dir, exist_ok=True)

try:
    # Kaggle ZIP ダウンロード
    subprocess.run(
        ["kaggle", "datasets", "download", "-d", dataset, "-p", download_dir],
        check=True
    )

    # ZIP ファイルを特定
    zip_files = glob.glob(os.path.join(download_dir, "*.zip"))
    if not zip_files:
        raise FileNotFoundError("ZIP file not found.")
    zip_path = zip_files[0]

    # 日付
    today = datetime.datetime.now().strftime("%Y%m%d")

    with zipfile.ZipFile(zip_path, "r") as z:
        extracted_files = z.namelist()   # ZIP 内のファイル名だけ取得
        z.extractall(output_dir)         # 解凍

    # ZIP 削除
    os.remove(zip_path)

    # === 解凍されたファイルだけ日付を付ける ===
    attached_files = []  

    for file in extracted_files:
        old_path = os.path.join(output_dir, file)

        # 解凍されたファイルだけ処理（フォルダはスキップ）
        if os.path.isfile(old_path):
            name, ext = os.path.splitext(file)
            new_name = f"{name}_{today}{ext}"
            new_path = os.path.join(output_dir, new_name)

            os.rename(old_path, new_path)

            attached_files.append(new_path) 

    # ================================
    #  Outlook で自分自身にメール送信
    # ================================
    my_email = os.getenv("MY_EMAIL")
    if not my_email:
        raise EnvironmentError("環境変数 MY_EMAIL が設定されていません。")

    import subprocess

    # Outlook をウィンドウごと起動
    subprocess.Popen(["start", "outlook"], shell=True)

    outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = my_email
    mail.Subject = f"Kaggle データ取得 {today}"
    mail.Body = "本日のデータを送付します。"

    # 添付ファイルを追加（元コードの意図をそのまま実現）
    for f in attached_files:
        mail.Attachments.Add(f)

    mail.Send()
    print("メール送信完了")

except Exception as e:
    print(f"Error: {e}")
