# -*- coding: utf-8 -*-
import pandas as pd
import os

# === 配置区 ===
excel_file = r"grade.xlsx"  # 您的Excel文件名
save_dir = r"grade_text"  # 文本文件保存目录
dispatch_list_file = "dispatch_list.csv" # 输出的发送清单文件名

# === 读Excel ===
#【关键修改】使用 header=0 来正确地将Excel的第一行作为列标题
# 这将自动识别所有列，包括可选的"家长昵称"列
df = pd.read_excel(excel_file, header=0)
print("Excel文件已读取，第一行已用作标题。")
print("识别到的列名:", df.columns.tolist())


# === 创建文本保存目录 ===
os.makedirs(save_dir, exist_ok=True)

# === 准备发送清单数据 ===
dispatch_data = []

# === 按学号分组，生成文本和清单条目 ===
for sid, group in df.groupby('学号'):
    # 确保学号不是空的，防止处理空行
    if pd.isna(sid):
        continue

    student_name = group['姓名'].iloc[0]

    # --- 核心逻辑：确定要搜索的微信联系人名称 ---
    parent_name_primary = ""
    parent_name_secondary = ""
    parent_name_tertiary = ""

    # 检查 '家长昵称' 列是否存在且当前行的值不为空
    if '家长昵称' in df.columns and pd.notna(group['家长昵称'].iloc[0]):
        parent_name_primary = group['家长昵称'].iloc[0]
        parent_name_secondary = ""
        parent_name_tertiary = f"{student_name}家长"
        print(f"学生 [{student_name}] 使用指定的家长昵称: {parent_name_primary}, 备用: {parent_name_tertiary}")
    else:
        parent_name_primary = f"{student_name}爸爸"
        parent_name_secondary = f"{student_name}妈妈"
        parent_name_tertiary = f"{student_name}家长"

    # 生成格式化的成绩文本
    text_content = f"【{student_name}同学成绩单】\n"
    text_content += "=" * 0 + ""
    text_content += f"学号：{sid}\n"
    text_content += f"姓名：{student_name}\n"
    text_content += "-" * 0 + ""

    # 添加各科成绩
    for idx, row in group.iterrows():
        subject = row['科目']
        score = row['成绩']
        text_content += f"{subject}：{score}\n"

    text_content += "=" * 0 + ""

    # 保存文本到文件（可选）
    text_filename = f"{student_name}成绩.txt"
    text_path = os.path.join(save_dir, text_filename)
    with open(text_path, 'w', encoding='utf-8') as f:
        f.write(text_content)
    print(f"-> 已生成文本文件: {text_path}")

    # 将该学生的所有信息添加到发送清单中，包含文本内容
    dispatch_data.append({
        'student_name': student_name,
        'wechat_contact_primary': parent_name_primary,
        'wechat_contact_secondary': parent_name_secondary,
        'wechat_contact_tertiary': parent_name_tertiary,
        'text_content': text_content,  # 直接包含文本内容
        'text_file_path': os.path.abspath(text_path),  # 文本文件路径（备用）
    })

# === 将清单数据保存到CSV文件 ===
dispatch_df = pd.DataFrame(dispatch_data)
dispatch_df.to_csv(dispatch_list_file, index=False, encoding='utf-8-sig')

print("\n=======================================================")
print(f"所有学生成绩文本已生成于: {os.path.abspath(save_dir)}")
print(f"发送清单 '{dispatch_list_file}' 已生成完毕。")
print("清单中包含了可直接发送的文本内容。")
print("现在可以去 Power Automate Desktop 中运行您的机器人了。")
print("=======================================================\n")