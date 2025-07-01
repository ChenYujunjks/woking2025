import pandas as pd
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))  # 即 src/ 目录

# ==== 文件路径 ====
master_file = os.path.join(BASE_DIR, "total.xlsx")  # src/total.xlsx

student_filename = "金以诺.xlsx"
student_file = os.path.join(BASE_DIR, "..", student_filename)

# 输出文件放在项目根目录的 output 文件夹
output_filename = student_filename.replace(".xlsx", "output.xlsx")
output_file = os.path.join(BASE_DIR, "..", "output", output_filename)

# 确保输出目录存在
os.makedirs(os.path.dirname(output_file), exist_ok=True)
 
# ==== 读取 master 总表格 ====
master_df = pd.read_excel(master_file, header=1)

# ✅ 构造 master 匹配键（补全前导零：院校 4 位，专业 3 位）
院校代码列 = pd.to_numeric(master_df.iloc[:, 5], errors="coerce").astype("Int64").astype(str).str.zfill(4)
专业代码列 = pd.to_numeric(master_df.iloc[:, 7], errors="coerce").astype("Int64").astype(str).str.zfill(3)
master_df["匹配键"] = 院校代码列 + "-" + 专业代码列

# ✅ 安全覆盖原始列（防止类型冲突）
master_df.iloc[:, 5] = 院校代码列.astype(object)
master_df.iloc[:, 7] = 专业代码列.astype(object)

print("📌 master 匹配键示例：", master_df["匹配键"].head(5).to_list())

# ==== 读取学生志愿表 ====
student_df_raw = pd.read_excel(student_file, header=1)

# 构造 student 的匹配键（第2列 + 第4列），需处理 nan 和小数
院校_raw = student_df_raw.iloc[:, 1]
专业_raw = student_df_raw.iloc[:, 3]

院校_clean = pd.to_numeric(院校_raw, errors='coerce')
专业_clean = pd.to_numeric(专业_raw, errors='coerce')
valid_index = 院校_clean.notna() & 专业_clean.notna()

院校_clean_str = 院校_clean[valid_index].astype(int).astype(str).str.zfill(4)
专业_clean_str = 专业_clean[valid_index].astype(int).astype(str).str.zfill(3)

student_df = student_df_raw.loc[valid_index].copy()
student_df["匹配键"] = 院校_clean_str + "-" + 专业_clean_str

# ✅ 安全覆盖原 DataFrame 中原始列（防止 dtype 警告）
student_df.iloc[:, 1] = 院校_clean_str.astype(object)  # 院校代码列（第2列）
student_df.iloc[:, 3] = 专业_clean_str.astype(object)  # 专业代码列（第4列）

print("📌 student 匹配键示例：", student_df["匹配键"].head(5).to_list())

# ==== 合并：以 student_df 为主表，补充 master 信息 ====
merged_df = pd.merge(
    student_df,
    master_df,
    on="匹配键",
    how="left",                # 保留 student_df 的顺序
    suffixes=('', '_master')  # 避免列名冲突
)

# 删除 student 原有的学制和学费列，避免重复
for col in ["学费", "学制"]:
    if col in merged_df.columns and col + "_master" in merged_df.columns:
        merged_df.drop(columns=[col], inplace=True)

# ✅ 重命名最低分、最低位次字段，及 master 学制/学费字段
rename_map = {
    "最低分": "最低分_24",
    "最低分.1": "最低分_23",
    "最低分.2": "最低分_22",
    "最低\n位次": "最低位次_24",
    "最低\n位次.1": "最低位次_23",
    "最低\n位次.2": "最低位次_22",
    "学制_master": "学制",
    "学费_master": "学费",
}
merged_df.rename(columns=rename_map, inplace=True)

# ✅ 删除 master 表中冗余的“专业代码_master”列
if "专业\n代码_master" in merged_df.columns:
    merged_df.drop(columns=["专业\n代码_master"], inplace=True)
# ==== 构造导出字段（保留志愿序号前两列 + 所有合并字段） ====
columns_to_export = list(student_df.columns[:2]) + list(merged_df.columns)
columns_to_export = list(dict.fromkeys(columns_to_export))  # 去重列名
final_df = merged_df.loc[:, columns_to_export]

# ✅ 可选：删除 F~V 列（如有需要再打开）
final_df.drop(final_df.columns[4:18], axis=1, inplace=True)

# ✅ 匹配统计提示
student_total = len(student_df)
matched_total = len(final_df)
print(f"\n📊 学生志愿总数：{student_total} 个")
print(f"✅ 成功匹配导出：{matched_total} 条记录")

# ==== 保存结果 ====
final_df.to_excel(output_file, index=False)
print(f"\n📁 匹配结果已保存至：{output_file}")
