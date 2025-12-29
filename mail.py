import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.utils import formataddr

def send_basic_email():
    """
    发送基础文本邮件
    """
    # 配置参数
    #两种发送方式
    #1.1、是联系人，在通讯录里选择收件人，若曾设置昵称，发送则显示设置的昵称；
    #1.2、是联系人，直接输入邮箱账户发送邮件，则取邮箱账户（去掉@***）为昵称；
    #2.1、非联系人，直接输入邮箱账户发送邮件，直接取邮箱账户（去掉@***）为昵称。

    #该程序发送方式为1.1
    sender_email = "lzh210802066@163.com"  # 发件人邮箱
    sender_name = "赖志辉"                 # 发件人名称
    password = "TDmi94Lv9dMP5WED"          # 发件人邮箱授权码（不是登录密码！）

    receiver_email = "adamlai6@163.com"   # 收件人邮箱
    receiver_name = "1345623254343"       # 收件人名称

    
    # SMTP服务器配置
    smtp_server = "smtp.163.com"
    smtp_port = 465  # SSL端口
    
    # 创建邮件对象
    message = MIMEMultipart()
    
    # 设置邮件头部信息（支持中文名称）
    message['From'] = formataddr([sender_name, sender_email])
    message['To'] = formataddr([receiver_name, receiver_email])
    message['Subject'] = Header('Python自动邮件测试', 'utf-8')
    
    # 添加纯文本正文
    text_content = """
    尊敬的{}：
    
    这是一封来自Python的测试邮件。
    
    发送时间：{}
    发件人：{}
    
    祝好！
    Python自动发送程序
    """.format(receiver_name, "2024年", sender_name)
    
    message.attach(MIMEText(text_content, 'plain', 'utf-8'))
    
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

# 调用
send_basic_email()