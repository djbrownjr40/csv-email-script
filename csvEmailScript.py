import chardet
import imaplib
import email
import csv
import re  # To handle removing excessive spaces
from email.header import decode_header
from email.utils import parsedate
from datetime import datetime


# サーバーの情報
imap_server = "mail.xyz.co.jp"
username = "username"
password = "password"


# IMAPサーバーに接続
mail = imaplib.IMAP4(imap_server, 143)
mail.login(username, password)


# INBOXを選択する
mail.select("inbox")


# メールごとを探す
status, messages = mail.search(None, "ALL")
email_ids = messages[0].split()


# CSVのファイルパス
csv_file_path = "C:\\"


emails_by_sender = {}


# メールをループする
for index, email_id in enumerate(email_ids, 1):
    # IDで取得する
    status, msg_data = mail.fetch(email_id, "(RFC822)")


    for response_part in msg_data:
        if isinstance(response_part, tuple):
            msg = email.message_from_bytes(response_part[1])


            # デコード：件名
            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding if encoding else "utf-8")


            # 送信者
            from_ = msg.get("From")


            # 日付
            date_str = msg.get("Date")
            if date_str:
                parsed_date = parsedate(date_str)
                if parsed_date:
                    date = datetime(*parsed_date[:6]).strftime('%Y-%m-%d %H:%M:%S')
                else:
                    date = date_str
            else:
                date = ""


            # 本文
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        try:
                            body_data = part.get_payload(decode=True)


                            # chardetでエンコーディングを取得する
                            detected_encoding = chardet.detect(body_data)


                            # utf-8でエンコーディングを取得する
                            encoding = detected_encoding['encoding'] if detected_encoding['encoding'] else 'utf-8'
                            body = body_data.decode(encoding, errors='ignore')
                        except UnicodeDecodeError:
                            body = part.get_payload(decode=True).decode("utf-8", errors="ignore")
                        break
            else:
                try:
                    body_data = msg.get_payload(decode=True)


                    # chardetでエンコーディングを取得する
                    detected_encoding = chardet.detect(body_data)


                    # utf-8でエンコーディングを取得する
                    encoding = detected_encoding['encoding'] if detected_encoding['encoding'] else 'utf-8'
                    body = body_data.decode(encoding, errors="ignore")
                except UnicodeDecodeError:
                    body = msg.get_payload(decode=True).decode("utf-8", errors="ignore")


            # スペース、改行を消す
            body = body.replace("\n", " ").replace("\r", " ")


            if from_ not in emails_by_sender:
                emails_by_sender[from_] = []


            emails_by_sender[from_].append([index, from_, subject, date, body])


# CSVを開く (after all emails are processed)
with open(csv_file_path, "w", newline='', encoding="utf-8-sig") as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(["Order", "From", "Subject", "Date", "Body"])


    # Write all emails to the CSV
    for sender, emails in emails_by_sender.items():
        for email_data in emails:
            writer.writerow(email_data)


# IMAPサーバーから切断する
mail.logout()


print(f"Emails have been exported to: {csv_file_path}")
