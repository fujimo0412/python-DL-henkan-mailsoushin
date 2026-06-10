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

today = datetime.datetime.now().strftime("%Y%m%d")

# ===== ダウンロード =====
try:
    subprocess.run(
        ["kaggle", "datasets", "download", "-d", dataset, "-p", download_dir],
        check=True
    )
except Exception as e:
    print(f"ダウンロードに失敗しました: {e}")
    exit(1)

# ===== 解凍・リネーム =====
try:
    zip_files = glob.glob(os.path.join(download_dir, "*.zip"))
    if not zip_files:
        raise FileNotFoundError("ZIPファイルが見つかりません。")
    zip_path = zip_files[0]

    with zipfile.ZipFile(zip_path, "r") as z:
        extracted_files = z.namelist()
        z.extractall(output_dir)

    os.remove(zip_path)

    attached_files = []
    for file in extracted_files:
        old_path = os.path.join(output_dir, file)
        if os.path.isfile(old_path):
            name, ext = os.path.splitext(file)
            new_name = f"{name}_{today}{ext}"
            new_path = os.path.join(output_dir, new_name)
            os.rename(old_path, new_path)
            attached_files.append(new_path)
except Exception as e:
    print(f"ファイル処理に失敗しました: {e}")
    exit(1)

# ===== メール送信 =====
try:
    my_email = os.getenv("MY_EMAIL")
    if not my_email:
        raise EnvironmentError("環境変数 MY_EMAIL が設定されていません。")

    subprocess.Popen(["start", "outlook"], shell=True)

    outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = my_email
    mail.Subject = f"Kaggle データ取得 {today}"
    mail.Body = "本日のデータを送付します。"

    for f in attached_files:
        mail.Attachments.Add(f)

    mail.Send()
    print("完了しました")
except Exception as e:
    print(f"メール送信に失敗しました: {e}")
    exit(1)
