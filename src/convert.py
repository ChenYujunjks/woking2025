import pandas as pd

# 读取 .xls 文件
df = pd.read_excel('your_file.xls', engine='xlrd')

# 写入为 .xlsx 文件
df.to_excel('your_file_converted.xlsx', index=False, engine='openpyxl')
