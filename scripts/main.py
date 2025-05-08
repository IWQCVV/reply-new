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
    # 创建 HTML 邮件内容
    msg = MIMEText(body, 'html', 'utf-8')
    msg["Subject"] = subject
    msg["From"] = SMTP_USER
    msg["To"] = to_email
    msg.add_header('Content-Type', 'text/html')  # 确保内容类型声明

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
            server.login(SMTP_USER, SMTP_PASSWORD)
            server.sendmail(SMTP_USER, [to_email], msg.as_string())
            print(f"邮件发送成功至 {to_email}")
    except Exception as e:
        print(f"邮件发送失败: {str(e)}")
        raise

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
        body = f"""
        <html>
<head>
    <style>
        body {{ font-family: Arial, sans-serif; line-height: 1.6; }}
        h4 {{ color: #2c3e50; }}
        .bank-info {{ background: #f8f9fa; padding: 15px; border-radius: 5px; }}
        .important {{ color: #e74c3c; }}
        .footer {{ margin-top: 20px; border-top: 1px solid #eee; padding-top: 10px; }}
    </style>
</head>
<body>
        <p>Dear colleagues and friends {rows[Id]}:</p>
    
    <p>Thank you very much for registering for the International Workshop on Quantum Characterization, Verification, and Validation (IWQCVV 2025).</p>
    
    <h4>Payment Instructions:</h4>
    <p>To complete the registration, please transfer the registration fee:</p>
    <ul>
        <li><strong>Regular:</strong> ¥2000</li>
        <li><strong>Students:</strong> ¥1500</li>
    </ul>
    
    <p>to the following bank account <strong>before Aug 10, 2025</strong>.</p>
    
    <div class="bank-info">
        <p><strong>Account Name (开户名称):</strong> 复旦大学 (Fudan University)<br>
        <strong>Address (地址):</strong> 上海市杨浦区邯郸路220号<br>
        <strong>Account Number (银行账号):</strong> 03326708017003441<br>
        <strong>Bank Name (开户银行):</strong> 中国农业银行上海翔殷支行<br>
        &nbsp;&nbsp;(Agricultural Bank of China, Shanghai Xiangyin Branch)<br>
        <strong>联行号:</strong> 103290035039</p>
    </div>
    
    <p><strong>Please:</strong></p>
    <ol>
        <li>Indicate <strong>the name of the participant</strong> when transferring the money</li>
        <li>Send a screenshot of proof of payment (付款凭证截图) to Ms Xinli Yan (<a href="mailto:yanxinli@fudan.edu.cn">yanxinli@fudan.edu.cn</a>)</li>
        <li>Send the information required for the invoice/receipt</li>
    </ol>
    
    <h4 class="important">Important Notes:</h4>
    <ul>
        <li>We can help book the Fraser Place Hotel <strong>only after receiving the registration fee</strong></li>
        <li>The registration fee can be waived for invited speakers upon request before June 1</li>
        <li>The gym and swimming pool in the Fraser Place Hotel are free to hotel residents</li>
        <li>We are very sorry that we cannot help book other hotels</li>
        <li>If you need to change the hotel or check in/out date, please let us know as soon as possible</li>
        <li>Book your hotel early as it's very difficult to find accommodation near campus in summer, even at 50% higher prices</li>
    </ul>
    
    <div class="footer">
        <p>If you have any questions, please contact:</p>
        <p><strong>Ms Xinli Yan</strong><br>
        Email: <a href="mailto:yanxinli@fudan.edu.cn">yanxinli@fudan.edu.cn</a><br>
        Phone: 021-3124 3502</p>
        
        <p>Best regards,<br>
        <strong>Xinli Yan</strong><br>
        On behalf of the organizers (Huangjun Zhu, Jiangwei Shang, and You Zhou)</p>
    </div>
    </body>
</html>
"""
        send_email(
             row["Email Address"],
            "[QCVV2025] Registration Confirmation",
            body
        )

    if not new_rows.empty:
        save_last_row(len(df))

if __name__ == "__main__":
    download_excel()
    process_data()
    os.remove("data.xlsx")  # 删除下载的Excel
