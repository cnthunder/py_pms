import time
import pandas as pd
import smtplib
import configparser
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# 版本：v240523，解决SSL协商问题，可以用公司邮箱了；
# 用途：将处理过的pms日志填写情况私发给各个工程师；

# 定义Excel文件路径
pms_file_path = "E:/WORK/Python_PMS/"  # 修改为实际文件目录
# 定义输入xls文件名
filename = input("输入待处理的文件名(邮件通知内容):")
file_path = f"{pms_file_path}{filename}.xlsx"
# 读取Excel文件
xls_df = pd.read_excel(file_path, sheet_name='Sheet1')  # sheet页

workdays = int(xls_df.iloc[1, 6])
if workdays <= 7:
    column_order = ['姓名', '工号', 'Base地', '岗位类别', '日志区间', '工作日', '工作日时长', '总日志时长', '请假时长', '日常日志时长', '项目日志时长', '项目日志占比', 'KPI参考', '排名', '日志时长考核', '邮箱']
elif workdays > 7:
    column_order = ['姓名', '工号', 'Base地', '岗位类别', '日志区间', '工作日', '工作日时长', '总日志时长', '请假时长', '日常日志时长', '项目日志时长', '项目日志占比', 'KPI参考', '排名', 'KPI有效值（0-150）', '邮箱']
df = pd.read_excel(file_path, sheet_name='Sheet1', usecols=column_order)

# 获取唯一的邮件地址列表
unique_emails = df['邮箱'].unique()

# 配置SMTP服务器和邮箱
config = configparser.ConfigParser()
config.read(f'{pms_file_path}mail_config.ini')
mail_settings = config['MailServer']
# mail_config.ini文件格式如下：
# [MailServer]
# smtp_server = smtp.xx.com
# smtp_port = 465
# smtp_username = xx@xx.com
# smtp_password = xxxx
smtp_server = mail_settings['smtp_server']
smtp_port = mail_settings['smtp_port']
smtp_username = mail_settings['smtp_username']
smtp_password = mail_settings['smtp_password']



if smtp_server == 'smtp.qq.com':
    print(f"使用QQ邮箱发送")
else:
    print(f"使用公司邮箱发送")
    import ssl
    ctx = ssl.create_default_context()
    ctx.set_ciphers('DEFAULT')
# 登录SMTP服务器
server = smtplib.SMTP_SSL(smtp_server, smtp_port, context=ctx)
server.login(smtp_username, smtp_password)

# 遍历每个唯一的邮件地址
for email in unique_emails:

    # 从原始数据中筛选具有相同邮件地址的行
    filtered_df = df[df['邮箱'] == email]
    # 将筛选后的数据转换为HTML表格
    html_table = filtered_df.to_html(index=False, classes='table table-bordered table-striped', escape=False)

    # 创建邮件
    msg = MIMEMultipart()
    msg['From'] = smtp_username
    msg['To'] = email
    msg['Subject'] = '工作数据报告'

    # 邮件正文，使用HTML格式
    body = f'<html><body><p>请查看以下工作数据报告：</p >{html_table}<p>发送人：王鹏（请勿回复，有问题企业微信联系）</p ></body></html>' #修改的邮件内容
    msg.attach(MIMEText(body, 'html'))

    # 发送邮件
    if email != '已离职':
        server.sendmail(smtp_username, email, msg.as_string())
        print(f"发送至{email}中...")
        time.sleep(1)
    else:
        print(f"已离职，跳过")

# 断开与邮件服务器的连接
server.quit()

print("邮件发送完毕")