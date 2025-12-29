import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.utils import formataddr
import pandas as pd
from datetime import datetime

# 读取 Excel 文件
def read_data():
    file_path = "data.xlsx"
    try:
        df = pd.read_excel(file_path)  # 读取 Excel 文件
        print("Data successfully loaded")
        return df
    except FileNotFoundError:
        print(f"Error: File {file_path} not found.")
        return None

# 发送邮件函数
def send_email(subject, content, to_email):
    sender_email = "lzh210802066@163.com"  # 发件人邮箱
    sender_name = "赖志辉"  # 发件人名称
    password = "TDmi94Lv9dMP5WED"  # 发件人邮箱授权码（不是登录密码！）

    # 收件人邮箱和名称
    receiver_email = "lzh210802066@163.com"
    receiver_name = "收件人名称"  # 可以根据需求修改或从 Excel 中读取

    # SMTP服务器配置
    smtp_server = "smtp.163.com"
    smtp_port = 465  # SSL端口

    # 创建邮件对象
    message = MIMEMultipart()

    # 设置邮件头部信息（支持中文名称）
    message['From'] = formataddr([sender_name, sender_email])
    message['To'] = formataddr([receiver_name, receiver_email])
    message['Subject'] = Header(subject, 'utf-8')

    # 添加 HTML 邮件正文
    message.attach(MIMEText(content, 'html', 'utf-8'))

    try:
        # 建立SSL连接
        print(f"正在连接服务器 {smtp_server}:{smtp_port}...")
        server = smtplib.SMTP_SSL(smtp_server, smtp_port)
        server.set_debuglevel(1)  # 显示调试信息

        # 登录邮箱
        print(f"正在登录 {sender_email}...")
        server.login(sender_email, password)

        # 发送邮件
        print(f"正在发送邮件给 {receiver_email}...")
        server.sendmail(sender_email, [receiver_email], message.as_string())

        # 关闭连接
        server.quit()
        print("✅ 邮件发送成功！")

    except smtplib.SMTPAuthenticationError:
        print("❌ 认证失败：请检查邮箱和授权码是否正确")
    except smtplib.SMTPException as e:
        print(f"❌ SMTP错误：{e}")
    except Exception as e:
        print(f"❌ 发送失败：{e}")

# 生成 HTML 邮件内容
def generate_email_content(data, start_index, num_words=10):
    today = datetime.now().strftime("%Y/%m/%d")
    content = f"<h2>今日日语单词 - {today}</h2>"
    content += "<table border='1' cellpadding='8' cellspacing='0' style='border-collapse:collapse; width:100%;'>"
    content += "<tr style='background-color:#007bff; color:#ffffff;'><th>单词</th><th>读音</th><th>意思</th></tr>"
    
    for i in range(num_words):
        index = (start_index + i) % (len(data) - 1) + 1  # 跳过标题行
        word = data.iloc[index][0]  # 假设第一列为单词
        reading = data.iloc[index][1]  # 假设第二列为读音
        meaning = data.iloc[index][2]  # 假设第三列为意思
        bg_color = "#ffffff" if i % 2 == 0 else "#cce5ff"  # 偶数行和奇数行的颜色不同
        content += f"<tr style='background-color:{bg_color};'>"
        content += f"<td><b>{word}</b></td>"
        content += f"<td>{reading}</td>"
        content += f"<td>{meaning}</td>"
        content += "</tr>"

    content += "</table>"
    return content

# 更新索引
def update_index(sheet_name, new_index):
    # 这里可以加上 Google Sheets 更新索引的功能，若需要
    pass

# 主函数
def main():
    sheet_name = "data.xlsx"  # 更改为你的 Excel 文件名
    data = read_data()  # 获取 Excel 数据
    
    # 获取上次发送的索引，默认从某个位置开始
    last_index = 0  # 假设我们从第一个单词开始，实际可以从 Excel 文件或其他位置读取
    
    # 生成邮件内容
    email_content = generate_email_content(data, last_index)
    
    # 获取今天日期并作为邮件主题
    subject = f"每日日语单词 - {datetime.now().strftime('%Y/%m/%d')}"
    
    # 发送邮件
    send_email(subject, email_content, "lzh210802066@163.com")  # 替换为目标邮箱
    
    # 更新索引（更新为发送的下一个位置）
    new_index = (last_index + 10) % (len(data) - 1)
    update_index(sheet_name, new_index)

if __name__ == "__main__":
    main()

