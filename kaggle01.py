import os
import zipfile
import datetime
import subprocess
import time
import win32com.client  # Outlook 送信用

download_dir = os.getenv("DOWNLOAD_DIR")
if not download_dir:
    raise EnvironmentError("環境変数 DOWNLOAD_DIR が設定されていません。")
dataset = "vivek468/superstore-dataset-final"

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
    dataset_name = dataset.split("/")[1]
    zip_path = os.path.join(download_dir, f"{dataset_name}.zip")
    if not os.path.exists(zip_path):
        raise FileNotFoundError(f"ZIPファイルが見つかりません: {zip_path}")

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
    my_email = os.getenv("MAIL_TO")
    if not my_email:
        raise EnvironmentError("環境変数 MAIL_TO が設定されていません。")

    subprocess.Popen(["start", "outlook"], shell=True)

    for _ in range(10):
        try:
            outlook = win32com.client.gencache.EnsureDispatch("Outlook.Application")
            break
        except Exception:
            time.sleep(2)
    else:
        raise RuntimeError("Outlookの起動がタイムアウトしました。")
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
