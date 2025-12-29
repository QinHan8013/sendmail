import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.utils import formataddr
import pandas as pd
from datetime import datetime
import os

# 将索引放到 mail.py 内部（全局变量）
current_index = 0  # 初始化索引值

# 读取 Excel 文件
def read_data(file_path="data.xlsx"):
    try:
        df = pd.read_excel(file_path)
        print("Data successfully loaded")
        return df
    except (FileNotFoundError, ValueError) as e:
        print(f"Error: {e}")
        return None

# 获取上次发送的索引（直接读取全局变量）
def get_last_index():
    global current_index  # 使用全局变量
    return current_index

# 保存当前索引（更新全局变量）
def save_last_index(index_value):
    global current_index  # 更新全局变量
    current_index = index_value  # 更新索引
    print(f"✅ 更新索引：{index_value}")

# 发送邮件函数
def send_email(subject, content, to_email, from_email="lzh210802066@163.com", password="TDmi94Lv9dMP5WED"):
    message = MIMEMultipart()
    message['From'] = formataddr(["赖志辉", from_email])
    message['To'] = formataddr(["收件人名称", to_email])
    message['Subject'] = Header(subject, 'utf-8')
    message.attach(MIMEText(content, 'html', 'utf-8'))

    try:
        with smtplib.SMTP_SSL("smtp.163.com", 465) as server:
            server.login(from_email, password)
            server.sendmail(from_email, [to_email], message.as_string())
            print("✅ 邮件发送成功！")
    except smtplib.SMTPAuthenticationError:
        print("❌ 认证失败：请检查邮箱和授权码是否正确")
    except Exception as e:
        print(f"❌ 发送失败：{e}")

# 生成 HTML 邮件内容
def generate_email_content(data, start_index, num_words=10):
    today = datetime.now().strftime("%Y/%m/%d")
    content = f"<h2>今日日语单词 - {today}</h2><table border='1' cellpadding='8' cellspacing='0' style='width:100%;'><tr style='background-color:#007bff; color:#ffffff;'><th>单词</th><th>读音</th><th>意思</th></tr>"

    for i in range(num_words):
        index = (start_index + i) % len(data)  # 循环索引
        word, reading, meaning = data.iloc[index, :3]  # 假设前3列为单词、读音和意思
        bg_color = "#ffffff" if i % 2 == 0 else "#cce5ff"
        content += f"<tr style='background-color:{bg_color};'><td><b>{word}</b></td><td>{reading}</td><td>{meaning}</td></tr>"

    content += "</table>"
    return content

# 主函数
def main():
    data = read_data()  # 获取 Excel 数据
    if data is None or len(data.columns) < 3:
        print("❌ 数据格式不正确，至少需要3列。")
        return

    # 获取上次发送的索引（如果存在）
    last_index = get_last_index()
    print(f"从索引 {last_index} 开始发送...")

    # 生成邮件内容
    email_content = generate_email_content(data, last_index)
    subject = f"每日日语单词 - {datetime.now().strftime('%Y/%m/%d')}"
    send_email(subject, email_content, "lzh210802066@163.com")  # 目标邮箱

    # 更新索引（更新为发送的下一个位置）
    new_index = (last_index + 10) % len(data)  # 每次发送10行
    save_last_index(new_index)  # 保存更新后的索引

if __name__ == "__main__":
    main()
