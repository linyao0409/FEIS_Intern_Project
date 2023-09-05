import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import datetime
from email import encoders
from email.mime.base import MIMEBase

def parse_input(input_str):
    # 尋找分號的位置
    semicolon_index = input_str.find(';')

    # 分割NAME和EMAIL
    name = input_str[:semicolon_index].strip()
    email = input_str[semicolon_index + 1:].strip()

    return name, email

def sendmail(bankname):
    mail_df = pd.read_excel("test_mail.xlsx")
    mail_series = mail_df[bankname]
    recipients = []

    for item in mail_series.values:
        if pd.notna(item):
            recipients.append(parse_input(item)[1])

    smtp_server = 'Webmail.feis.com.tw'
    smtp_port = 25

    # 创建邮件消息
    msg = MIMEMultipart()
    msg['Subject'] = str(datetime.date.today()) + bankname+' 遠智證券報價單'
    msg['From'] = 'FEIS-FI <Neil.Lin@feis.com.tw>'
    msg['To'] = ', '.join(recipients)  # 收件人列表需要以逗号和空格分隔
    current_time = datetime.datetime.now()
    formatted_time = current_time.strftime("%Y-%m-%d %H:%M:%S")
    ctype = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"  # for .xlsx files
    maintype, subtype = ctype.split("/", 1)

    email_text = f"附件為 {bankname} {formatted_time} 報價單，謝謝您!!!"
    another_text = ("有任何問題需要協助請連絡遠智證券經濟部"
                    "110台北市信義區信義路五段7號51樓"
                    "02-8758-3399")

    combined_text = email_text + "\n\n" + another_text
    msg.attach(MIMEText(combined_text, 'plain'))

    # Attach the first text
    #msg.attach(MIMEText(email_text, 'plain'))

    # Attach the second text
    #msg.attach(MIMEText(another_text, 'plain'))


    ch_name_to_en_name = {"林曜銀行":"LINYAO","玉山銀行":"ESUN","星展銀行":"DBS","國泰世華":"CATHAY","陽信銀行":"SUNNY"}

    # 添加附件 "2023-08-24LINYAO.xlsx"
    attachment_path = f"{ch_name_to_en_name[bankname]}_table_{str(datetime.date.today())}.xlsx"
    #attachment_path = f"{str(datetime.date.today())}{ch_name_to_en_name[bankname]}.xlsx"
    with open(attachment_path, "rb") as attachment_file:
        attachment = MIMEBase(maintype, subtype)
        attachment.set_payload(attachment_file.read())
        encoders.encode_base64(attachment)
        attachment.add_header("Content-Disposition", f"attachment; filename={attachment_path.split('/')[-1]}")  # using split to get the file name only
        msg.attach(attachment)
    #with open(attachment_path, "rb") as attachment_file:
        #attachment = MIMEApplication(attachment_file.read())
        #attachment.add_header("Content-Disposition", f"attachment; filename={attachment_path}")
        #msg.attach(attachment)

    # 连接 SMTP 服务器并发送邮件
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.sendmail("it@feis.com.tw", recipients, msg.as_string())
        print(bankname,"郵件發送成功")
    except Exception as e:
        print(bankname,"郵件發送失敗:", str(e))
    finally:
        server.quit()