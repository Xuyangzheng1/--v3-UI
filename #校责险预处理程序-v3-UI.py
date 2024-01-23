#校责险预处理程序-v3-UI
#姓名，班级，身份证
#序号，学校*，姓名*，性别*，证件类型*，身份证号或学号*，年龄*，年纪，备注
#终极版==》-v2 2023-10-16
#增加了身份证验真函数validate_id_card
#源码
#UI
#修改性别，男1女2
#2023/10/17 发现问题，1.应增加身份证号重复验证
#                    2.姓名列存入预处理文件时去掉空格
#                   2023/10/18修改完毕
#学生版

import os
import tkinter as tk
from tkinter import filedialog
import pandas as pd
from datetime import datetime
import re

def validate_id_card(id_card):
    if len(id_card) != 18:
        return False

    factors = [7, 9, 10, 5, 8, 4, 2, 1, 6, 3, 7, 9, 10, 5, 8, 4, 2]
    check_code_mapping = {0: 1, 1: 0, 2: 'X', 3: 9, 4: 8, 5: 7, 6: 6, 7: 5, 8: 4, 9: 3, 10: 2}

    total = sum(int(id_card[i]) * factors[i] for i in range(17))
    remainder = total % 11
    check_code = check_code_mapping[remainder]

    return str(check_code).lower() == id_card[17].lower()

def clean_id(id):
    return re.sub(r'\s+', '', str(id))


def calculate_age(id_number):
    id_number = clean_id(id_number)
    if len(id_number) != 18:
        return "无效的身份证号码"
    
    birth_date_str = id_number[6:14]
    
    try:
        birth_date = datetime.strptime(birth_date_str, '%Y%m%d')
        today = datetime.now()
        age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
        return age
    except ValueError:
        return "无效的身份证号码"

def determine_gender(id_number):
    id_number = id_number.strip()
    if len(id_number) != 18:
        return "无法确定性别"
    seventeenth_digit = int(id_number[16])
    return "1" if seventeenth_digit % 2 == 1 else "2"

def remove_keywords(input_string):
    keywords = ['.xls', '.xlsx', '学生']
    for keyword in keywords:
        input_string = re.sub(re.escape(keyword), '', input_string)
    return input_string

# Create a function to run the script with the selected paths
def run_script():
    directory_path = directory_path_entry.get()
    output_path = output_path_entry.get()
    
    if os.path.exists(directory_path):
        id_list = []
        age_list = []
        gender_list = []
        filename_list = []
        id_len_list = []
        name_list = []
        class_list = []
        problematic_ids = []

        #=======================================================================================================================================================

        # 添加一个集合来跟踪已处理的身份证号
        processed_id_set = set()

        for filename in os.listdir(directory_path):
            if filename.endswith('.xls') or filename.endswith('.xlsx'):
                try:
                    file_path = os.path.join(directory_path, filename)
                    df = pd.read_excel(file_path, dtype=str)
                    if '身份证' in df.columns:
                        print(file_path)
                        for id, name, banji in zip(df['身份证'], df['姓名'], df['班级']):
                            id = clean_id(id)
                            name=clean_id(name)
                            # 检查身份证号是否已经被处理
                            if id in processed_id_set:
                                problematic_ids.append({
                                    '姓名': name,
                                    '身份证号': id,
                                    '班级': banji,
                                    '来源': filename,
                                    '年龄': '重复的身份证号'
                                })
                                print(id,name,filename)
                                id_list.append(id)
                                age_list.append(age)
                                gender_list.append(gender)
                                filename = remove_keywords(filename)
                                filename_list.append('镇平县' + filename)
                                id_len_list.append(len(id))
                                name_list.append(name)
                                class_list.append(banji)
                            else:
                                age = calculate_age(id)
                                gender = determine_gender(id)

                                if age == "无效的身份证号码" or gender == "无法确定性别" or int(age) < 0 or int(age) > 18 or not validate_id_card(id):#《<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<学生版，用于判断年龄
                                    problematic_ids.append({
                                        '姓名': name,
                                        '身份证号': id,
                                        '班级': banji,
                                        '来源': filename,
                                        '年龄': age
                                    })

                                id_list.append(id)
                                age_list.append(age)
                                gender_list.append(gender)
                                filename = remove_keywords(filename)
                                filename_list.append('镇平县' + filename)#=========================<<<<<<
                                id_len_list.append(len(id))
                                name_list.append(name)
                                class_list.append(banji)

                            # 将已处理的身份证号添加到集合中
                            processed_id_set.add(id)
                    else:
                        print('在文件', file_path, '中没有找到 "身份证" 列')
                except Exception as e:
                    print(f"读取 {filename} 时出现错误: str{e}")


 #=======================================================================================================================================================

        problematic_ids_df = pd.DataFrame(problematic_ids)
        problematic_ids_file_path = os.path.join(output_path, 'problematic_ids.xlsx')
        problematic_ids_df.to_excel(problematic_ids_file_path, index=False, header=True)

        data = {
            '学校*': filename_list,
            '姓名*': name_list,
            '性别*': gender_list,
            '身份证号或学号*': id_list,
            '年龄*': age_list,
            '年级': class_list,
        }
        data['备注'] = [''] * len(id_list)
        results_df = pd.DataFrame(data)
        results_df['序号'] = range(len(results_df))
        results_df = results_df.set_index('序号')
        print(results_df)
        print(results_df.columns)
        print(processed_id_set)

        #==========================
        results_df_path=os.path.join(output_path, 'all_data.xlsx')
        
        results_df.to_excel(results_df_path,index=False, header=True)
        #==========================

        grouped = results_df.groupby('学校*')

        with pd.ExcelWriter(os.path.join(output_path, '分校学生数据.xlsx'), engine='xlsxwriter') as writer:
            for school, group in grouped:
                group.insert(0, '序号', range(1, len(group) + 1))
                group['证件类型*'] = "01"
                column_order = ['序号', '学校*', '姓名*', '性别*', '证件类型*', '身份证号或学号*', '年龄*', '年级', '备注']
                group = group[column_order]
                file_name = os.path.join(output_path, f'{school}.xlsx')
                with pd.ExcelWriter(file_name, engine='xlsxwriter') as writer:
                    group.to_excel(writer, index=False, header=True)

        print(f"分校学生数据已保存到路径: {output_path}")

# Create the main window
root = tk.Tk()
root.title("Data Processing")
def select_directory():
    directory = filedialog.askdirectory()
    directory_path_entry.delete(0, tk.END)  # Clear the entry field
    directory_path_entry.insert(0, directory)  # Set the selected directory

# Function to select the output path
def select_output_path():
    output = filedialog.askdirectory()
    output_path_entry.delete(0, tk.END)  # Clear the entry field
    output_path_entry.insert(0, output)  # Set the selected output directory
# Create and place UI elements
directory_label = tk.Label(root, text="Select Directory Path:")
directory_label.pack()

directory_path_entry = tk.Entry(root)
directory_path_entry.pack()

directory_button = tk.Button(root, text="Browse", command=select_directory)
directory_button.pack()

output_label = tk.Label(root, text="Select Output Path:")
output_label.pack()

output_path_entry = tk.Entry(root)
output_path_entry.pack()

output_button = tk.Button(root, text="Browse", command=select_output_path)
output_button.pack()

run_button = tk.Button(root, text="Run Script", command=run_script)
run_button.pack()

# Function to select the directory path


root.mainloop()
