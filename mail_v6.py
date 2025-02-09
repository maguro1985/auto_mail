import tkinter as tk
import smtplib
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.header import Header
import datetime

# 送信元と送信先の情報を設定
from_email = "Maguro <maguro@xxx.jp>"  # 送信者名を含めて設定
to_email = "abc@xxx.jp"
cc_emails = ["", "", ""]

# 今日の日付を取得
today_date = datetime.date.today()

# 今週の月曜日と金曜日の日付を計算
monday_date = today_date - datetime.timedelta(days=today_date.weekday())
friday_date = monday_date + datetime.timedelta(days=4)

# ファイル名を生成
file_name = f"週報{friday_date.strftime('%Y%m%d')}.xlsx"

# ファイルのパスを指定
file_path = os.path.join(r'C:\Users\ユーザー名\Desktop\勤務報告書', file_name)  # フォルダのパスも適切に指定

# メールの件名と本文を生成
subject = f"週報提出:Maguro{monday_date.month}月{monday_date.day}日～{friday_date.month}月{friday_date.day}日"
body = """テスト
　

　
　
　
"""

# UTF-8でエンコード
encoded_file_name = Header(file_name, 'UTF-8').encode()
encoded_subject = Header(subject, 'UTF-8').encode()
encoded_body = body.encode('utf-8')

# SMTPサーバーの詳細情報を設定
smtp_host = 'smtp.xxx.jp'
smtp_port = xxx
smtp_username = 'maguro@xxx.jp'
smtp_password = 'xxxxx'  # ここにSMTPのパスワードを直接設定

# メールを送信する関数
def send_email():
    try:
        # メールサーバーと接続
        server = smtplib.SMTP_SSL(smtp_host, smtp_port)
        server.login(smtp_username, smtp_password)

        # メールを作成
        msg = MIMEMultipart()
        msg["From"] = from_email
        msg["To"] = to_email
        msg["Cc"] = ', '.join(cc_emails)
        msg["Subject"] = encoded_subject

        # メール本文を追加
        msg.attach(MIMEText(encoded_body, "plain", "utf-8"))

        # ファイルの添付処理
        if os.path.isfile(file_path):
            with open(file_path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename="{encoded_file_name}"')
                msg.attach(part)
        else:
            raise Exception("ファイルが見つかりませんでした.")

        # メールを送信
        server.send_message(msg)
        server.quit()

        # メール送信が成功したらTrueを表示
        result_label.config(text="メールを送信しました")

        # 現在の日時を取得
        current_time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")

        # 成功したログを記録
        with open("log.txt", "a") as log_file:
            log_file.write(f"{current_time},True\n")

    except Exception as e:
        # メール送信が失敗したらFalseを表示
        result_label.config(text=f"送信エラー\nエラーメッセージ: {e}")

        # エラーのログを記録
        current_time = datetime.datetime.now().strftime("%Y/%m/%d %H:%M:%S")
        with open("log.txt", "a") as log_file:
            log_file.write(f"{current_time},False\n")

# GUIウィンドウを作成
window = tk.Tk()
window.title("メール送信プログラム")
window.geometry("320x240")  # ウィンドウサイズを設定

# 送信ボタンを作成
send_button = tk.Button(window, text="送信", command=send_email)
send_button.pack()

# 結果表示用ラベル
result_label = tk.Label(window, text="", wraplength=300)
result_label.pack()

# ウィンドウを表示
window.mainloop()
