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
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style type="text/css">
            /* 通用样式 */
            body {{
                font-family: Arial, sans-serif;
                line-height: 1.6;
                color: #333333;
                margin: 20px;
                -webkit-text-size-adjust: 100%;
                mso-line-height-rule: exactly;
            }}
            /* Outlook兼容样式 */
            .ExternalClass {{ width: 100%; }}
            .ExternalClass, .ExternalClass p, .ExternalClass span,
            .ExternalClass font, .ExternalClass td, .ExternalClass div {{ line-height: 100%; }}
            
            h4 {{
                color: #444444;
                border-bottom: 1px solid #dddddd;
                padding-bottom: 5px;
                mso-line-height-rule: exactly;
            }}
            
            .bank-info {{
                border-left: 4px solid #cccccc !important;
                mso-border-left-alt: solid #cccccc 4px; /* Outlook专用 */
                padding: 15px;
                margin: 15px 0;
                background-color: #f8f8f8;
            }}
            
            .important {{
                color: #555555;
                font-weight: bold;
                background-color: #f8f9fa;
                padding: 2px 4px;
                mso-color-alt: #555555; /* Outlook备用 */
            }}
            
            .footer {{
                margin-top: 20px;
                padding-top: 15px;
                color: #666666;
                font-size: 0.9em;
                border-top: 1px solid #eeeeee;
            }}
            
            a {{
                color: #1155cc;
                text-decoration: underline;
                mso-style-priority: 99;
            }}
            
            /* 响应式表格 */
            @media screen and (max-width: 600px) {{
                .bank-info-table td {{
                    display: block;
                    width: 100% !important;
                }}
            }}
        </style>
    </head>
    <body>
        <p>Dear {row["Title1"]} {row["Given Name"]} {row["Family Name"]}:</p>
    
        <p>Thank you very much for registering for the International Workshop on Quantum Characterization, Verification, and Validation (IWQCVV 2025).
        Please check the following website for the latest update: <a href="https://iwqcvv.org/">https://iwqcvv.org/</a></p>
    
        <h4>Payment Instructions:</h4>
        <p>To complete the registration, please transfer the registration fee:</p>
        <ul>
            <li><span class="important">Regular:</span> CNY 2000</li>
            <li><span class="important">Students:</span> CNY 1500</li>
        </ul>
        
        <p>to the following bank account <strong>before Aug 5, 2025</strong>.</p>
        
        <div class="bank-info">

            <table class="bank-info-table" cellpadding="8" style="width: 100%;">
                <tr><td><strong>Account Name:</strong></td><td>复旦大学</td></tr>
                <tr><td><strong>Address:</strong></td><td>上海市杨浦区邯郸路220号</td></tr>
                <tr><td><strong>Account No.:</strong></td><td>03326708017003441</td></tr>
                <tr><td><strong>Bank Name:</strong></td><td>中国农业银行上海翔殷支行</td></tr>
                <tr><td><strong>Interbank Routing Number:</strong></td><td>103290035039</td></tr>
            </table>
        </div>
        
        <p><strong>Please:</strong></p>
        <ol>
            <li>Indicate <strong class="important">the participant's full name</strong> in the transfer memo.</li>
            <li>Email payment confirmation to Ms. Xinli Yan:
                <a href="mailto:yanxinli@fudan.edu.cn" style="word-break: break-all;">yanxinli@fudan.edu.cn</a>.
            </li>
            <li>Provide invoice details:
                <ul>
                    <li>Affiliation (单位名称)</li>
                    <li>Tax ID (纳税人识别号)</li>
                    <li>Invoice Type (发票类型: 普通发票/增值税专用发票)</li>
                </ul>
            </li>
        </ol>
        
        <h4>Important Notes:</h4>
        <ul>
            <li>To book Hotel Fraser Place Wujiaochang Shanghai through the Organizer, it is necesary to transfer the registration fee before <strong>Aug 5, 2025</strong>.</strong></li>
            <li>Invited speakers may request fee waiver before June 30.</li>
            <li>Hotel amenities include complimentary gym and pool access.</li>
            <li>Early booking strongly recommended due to high demand.</li>
        </ul>

        <h4>For your convenience, here is a brief summary of your registration information:</h4>
        <ul>
            <li>Title: {row["Title1"]}</li>
            <li>Given Name: {row["Given Name"]}</li>
            <li>Family Name: {row["Family Name"]}</li>
            <li>Affiliation: {row["Affiliation"]}</li>
            <li>Email Address: {row["Email Address"]}</li>
            <li>Phone Number: {row["Phone number"]}</li>
            <li>ID: {row["ID1"]}</li>
            <li>Arrival Date: {row["Arrival Date"].strftime('%Y-%m-%d')}</li>
            <li>Departure Date: {row["Departure Date"].strftime('%Y-%m-%d')}</li>
            <li>Accommodation: {row["Accommodation"]}</li>
            <li>Poster Submission: {row["Poster Submission"]}</li>
        </ul>
        
        <div class="footer">
            <p>Contact Information:</p>
            <p><strong>Ms. Xinli Yan</strong><br>
            Conference Coordinator<br>
            Email: <a href="mailto:yanxinli@fudan.edu.cn">yanxinli@fudan.edu.cn</a><br>
            Tel: +86-21-3124 3502</p>
            
            <p style="margin-top: 15px;">Best regards,<br><br>
            <strong>IWQCVV 2025 Organizing Committee</strong><br>
            Huangjun Zhu, Jiangwei Shang, You Zhou</p>
        </div>
    </body>
</html>
"""
        send_email(
             row["Email Address"],
            "[QCVV2025] Registration Confirmation",
            body)
    if not new_rows.empty:
        save_last_row(len(df))

if __name__ == "__main__":
    download_excel()
    process_data()
    os.remove("data.xlsx")  # 删除下载的Excel
