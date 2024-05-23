from datetime import datetime
import time
import pandas as pd
import configparser
import yagmail

# 版本：v240523，解决SSL协商问题，可以用公司邮箱了；
# 用途：进行PMS日志处理，按分组写到各组的xlsx文件中，获取汇报对象的邮箱，并作为附件发送至汇报对象；

# 定义Excel文件路径
pms_file_path = "E:/WORK/Python_PMS/"  # 修改为实际文件目录

# 定义人员清单文件名，供后续处理合并参考
file_path_list = f"{pms_file_path}KPI-List.xlsx"

# 定义输入xls文件名
filename = input("输入PMS直接导出日志的文件名:")
file_path_xls = f"{pms_file_path}{filename}.xls"
# 读取人员分组信息
list_df = pd.read_excel(file_path_list, sheet_name='人员分组', dtype={'工号': str})
# 读取xls文件中日志
xls_df = pd.read_excel(file_path_xls)
# 将日志日期调整为%Y/%m/%d格式
xls_df['日志日期'] = xls_df['日志日期'].dt.strftime('%Y/%m/%d')
# 修改PMS导出姓名不符的情况
xls_df.loc[xls_df['支持工程师'] == '陈帅(武汉)', '支持工程师'] = '陈帅'
# 创建kpi评分表
kpi_df = pd.DataFrame(columns=['姓名', '日志及时性', '日志准确性', '项目日志占比', '周报及时性', '周报准确性', '工作闭环及时性', '工作闭环完整性','出差报备及时性', '出差报备完整性','出差反馈及时性','出差反馈完整性', '备注'])
# 创建一个group来存储分组后的数据
group = {}
# 遍历分组后的数据，并将每个组的数据存储到group中，格式为{('组名', '汇报对象'): ['成员']}
for group_name, group_data in list_df.iterrows():
    #
    group_name = (group_data['组别'], group_data['汇报对象'])
    # 将每个组的数据存储为列表
    if group_name not in group:
        group[group_name] = []
    group[group_name].append(group_data['姓名'])


# 配置SMTP服务器和邮箱
config = configparser.ConfigParser()
config.read(f'{pms_file_path}mail_config.ini')
mail_settings = config['MailServer']
# mail_config.ini文件格式如下：
# [MailServer]
# smtp_server = mail.xx.com
# smtp_port = 465
# smtp_username = xx@xx.com
# smtp_password = xxxx
smtp_server = mail_settings['smtp_server']
smtp_port = mail_settings['smtp_port']
smtp_username = mail_settings['smtp_username']
smtp_password = mail_settings['smtp_password']
# 判断是否使用公司邮箱，如果是，自定义SSL协商
if smtp_server == 'smtp.qq.com':
    print(f"使用QQ邮箱发送")
else:
    print(f"使用公司邮箱发送")
    import ssl
    ctx = ssl.create_default_context()
    ctx.set_ciphers('DEFAULT')
# 使用yagmail，正式使用请调整
#yag = yagmail.SMTP(user=smtp_username, password=smtp_password, host=smtp_server, port=smtp_port, context=ctx)
# 测试专用，避免发送邮件出去。
yag = yagmail.SMTP(user='username@mail.com', password='smtp_password', host='smtp.qq.com', port=smtp_port, context=ctx)
# 定义邮件内容
body = (f'<html><body>'
        f'<p>1、请按照PMS日志要求，对团队成员上周的PMS日志进行检查；</p>'
        f'<p>2、根据KPI考核要求完成上周KPI评分，邮件反馈评分表格；</p>'
        f'<p>3、敦促团队成员对存在问题的PMS日志进行修改；</p >'
        f'<p>4、协助团队成员对上周工作进行总结，对于可能存在的问题，协助其进行改进；</p >'
        f'</body></html>')
# 进行PMS日志处理，按分组写到各组的xlsx文件中，获取汇报对象的邮箱，并作为附件发送
for group_groupname, group_members in group.items():
    tomail_xlsx = f"{pms_file_path}pms日志_{group_groupname[0]}.xlsx"
    tomail_kpi = f"{pms_file_path}kpi_{group_groupname[0]}.xlsx"
    with pd.ExcelWriter(tomail_xlsx, engine='openpyxl') as writer:
        for member in group_members:
            if member in group_members:
                filtered_df = xls_df[xls_df['支持工程师'] == member]
                member_mail = list_df[list_df['姓名'] == member]['邮箱'].iloc[0]
                filtered_df.to_excel(writer, sheet_name=member, index=False)
                # 通过邮箱对比，不将组长自己加入到kpi评分表中
                if group_groupname[1] != member_mail:
                    kpi_row = {'姓名': member, '日志及时性': '本项不评分', '项目日志占比': '本项不评分',}
                    kpi_df = kpi_df._append(kpi_row, ignore_index=True)
    kpi_df.to_excel(tomail_kpi, index=False)
    kpi_df.drop(index=kpi_df.index[0:], inplace=True)
    # 邮件发送各组附件
    if group_groupname[1] is not None:
        subject = f"{group_groupname[0]}本周日志，发送时间{datetime.now().strftime("%Y-%m-%d")}"
        try:
            yag.send(to=group_groupname[1], subject=subject, contents=body, attachments=[tomail_xlsx, tomail_kpi])
            print(f"{group_groupname[0]}邮件发送至{group_groupname[1]}")
            time.sleep(1)
        except Exception as e:
            print(f"{group_groupname[0]}邮件发送失败: {e}")

print("邮件发送完毕")