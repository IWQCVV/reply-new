import os
import pandas as pd
import requests
import smtplib
from email.mime.text import MIMEText

# 配置从环境变量读取
EXCEL_URL = os.getenv("EXCEL_URL")
SMTP_SERVER = "mail.fudan.edu.cn"
SMTP_PORT = 465
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASSWORD = os.getenv("SMTP_PASSWORD")

def load_last_row():
    try:
        with open("last_row.txt", "r") as f:
            return int(f.read().strip())
    except (FileNotFoundError, ValueError):
        return 0

def save_last_row(last_row):
    with open("last_row.txt", "w") as f:
        f.write(str(last_row))

def send_email(to_email, subject, body):
    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = SMTP_USER
    msg["To"] = to_email
    print(to_email)
    print(SMTP_SERVER,SMTP_PORT)
    with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASSWORD)
        server.sendmail(SMTP_USER, [to_email], msg.as_string())

def download_excel():
    response = requests.get(EXCEL_URL)
    response.raise_for_status()  # 检查请求是否成功
    with open("data.xlsx", "wb") as f:
        f.write(response.content)

def process_data():
    df = pd.read_excel("data.xlsx", engine="openpyxl")
    last_row = load_last_row()
    new_rows = df.iloc[last_row:]

    for _, row in new_rows.iterrows():
        send_email(
             row["Email Address"],
            "[QCVV2025] Registration Confirmation",
            f"尊敬的 {row['Full Name']}，您的报名已成功！"
        )

    if not new_rows.empty:
        save_last_row(len(df))

if __name__ == "__main__":
    download_excel()
    process_data()
    os.remove("data.xlsx")  # 删除下载的Excel
