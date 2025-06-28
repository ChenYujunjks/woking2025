import pandas as pd

# ==== 文件路径 ====
master_file = "total.xlsx"
student_file = "student.xlsx"
output_file = "提取后的结果.xlsx"

# ==== 读取总表格 ====
master_df = pd.read_excel(master_file, header=1)

# 构造 master_df 的匹配键（第6列 + 第8列）
院校代码列 = master_df.iloc[:, 5].astype(str).str.strip()
专业代码列 = master_df.iloc[:, 7].astype(str).str.strip()
master_df["匹配键"] = 院校代码列 + "-" + 专业代码列

print("📌 master 匹配键示例：", master_df["匹配键"].head(5).to_list())

# ==== 读取学生志愿表 ====
student_df = pd.read_excel(student_file, header=1)

# 构造匹配键（第2列 + 第4列），需处理小数 & 缺失值
院校_raw = student_df.iloc[:, 1]
专业_raw = student_df.iloc[:, 3]

# 处理 nan 和小数（转 int -> str），如 1.0 → 1，nan → 跳过
院校_clean = pd.to_numeric(院校_raw, errors='coerce').dropna().astype(int).astype(str)
专业_clean = pd.to_numeric(专业_raw, errors='coerce').dropna().astype(int).astype(str)

# 构造匹配键列表（注意保持行数对齐）
student_df = student_df.loc[院校_clean.index.intersection(专业_clean.index)].copy()
student_df["匹配键"] = 院校_clean + "-" + 专业_clean

print("📌 student 匹配键示例：", student_df["匹配键"].head(5).to_list())

# ==== 筛选 master 中匹配成功的记录 ====
student_keys_set = set(student_df["匹配键"])
matched_df = master_df[master_df["匹配键"].isin(student_keys_set)]

print(f"\n✅ 匹配成功记录数: {len(matched_df)}")

# ==== 输出内容：只保留 total 的前两列 + 匹配行的所有字段 ====
columns_to_export = list(master_df.columns[:2]) + list(matched_df.columns)
columns_to_export = list(dict.fromkeys(columns_to_export))  # 去重列名
final_df = matched_df.loc[:, columns_to_export]

# ==== 导出结果 ====
final_df.to_excel(output_file, index=False)
print(f"\n📁 匹配结果已保存到：{output_file}")
