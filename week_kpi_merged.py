import os
import pandas as pd
from datetime import datetime

# 合并各小组的周KPI为区域周KPI

# 设置包含Excel文件的目录路径
week_kpi_path = "E:/WORK/Python_PMS/区域周KPI合并/"
file_path_output = f"{week_kpi_path}周kpi合并_{datetime.now().strftime("%Y-%m-%d")}.xlsx"
# pms_file_path = "E:/WORK/Python_PMS/"
# file_path_output = f"{pms_file_path}周kpi合并_{datetime.now().strftime("%Y-%m-%d")}.xlsx"

# 初始化一个空的DataFrame，用于存储合并后的数据
kpi_merged_df = pd.DataFrame()

# 遍历目录中的所有文件
for filename in os.listdir(week_kpi_path):
    # 检查文件扩展名是否为xlsx
    if filename.endswith('.xlsx') and not filename.startswith('周kpi合并'):
        # 构建完整的文件路径
        file_path = os.path.join(week_kpi_path, filename)
        # 读取Excel文件
        kpi_df = pd.read_excel(file_path)
        # 填充空白单元格
        # kpi_df.fillna(0, inplace=True)
        # 将读取的数据追加到kpi_merged_df中
        kpi_merged_df = pd.concat([kpi_merged_df, kpi_df], ignore_index=True)

# 保存合并后的DataFrame到一个新的Excel文件

kpi_merged_df.to_excel(file_path_output, index=False)

print(f'All Excel files have been merged into {file_path_output}')