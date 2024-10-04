import csv
import os
import tkinter as tk
from tkinter import filedialog, messagebox


# 读取CSV并处理
def process_csv(input_csv_path):
    output_data = []
    # 打开输入CSV文件并读取内容
    with open(input_csv_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.reader(csvfile)
        for row in reader:
            if len(row) > 6:  # 确保行的列数足够
                email_content = row[6]  # 假设第7列是邮件内容
                if email_content:
                    output_data.append(email_content)
    return output_data


# 将处理后的数据写入.txt文件
def write_output(output_path, processed_data):
    with open(output_path, 'w', encoding='utf-8') as outfile:
        for data in processed_data:
            outfile.write(data + "\n\n###\n\n")  # 每条记录后添加分隔符 ###


if __name__ == "__main__":
    # 初始化Tkinter
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # 选择输入CSV文件
    input_csv_path = filedialog.askopenfilename(title="选择CSV文件", filetypes=[("CSV files", "*.csv")])

    if not input_csv_path:
        print("没有选择文件，程序将退出。")
        exit()

    # 输出文件名为原文件名加后缀 _processed.txt
    output_txt_path = os.path.splitext(input_csv_path)[0] + "_processed.txt"

    # 处理CSV并输出
    processed_data = process_csv(input_csv_path)
    if processed_data:
        write_output(output_txt_path, processed_data)

        # 弹窗提示已完成
        messagebox.showinfo("完成", f"处理后的内容已保存至：{output_txt_path}")

        # 打开文件夹
        folder_path = os.path.dirname(output_txt_path)
        os.startfile(folder_path)  # 打开包含输出文件的文件夹
    else:
        messagebox.showwarning("警告", "没有处理到有效的数据。")
