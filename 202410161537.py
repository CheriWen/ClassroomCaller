import os
import socket
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import shutil
import json
import sys

# 初始化主窗口
root = tk.Tk()
root.title("教师端点名系统")
root.geometry("800x500")

# 全局变量
devices = {}
config_file = "config.json"
classroom_menu = None  # 全局变量，用于下拉框
subject_var = tk.StringVar()  # 储存任教科目

# 获取程序路径
if getattr(sys, 'frozen', False):
    base_path = os.path.dirname(os.path.realpath(sys.executable))
else:
    base_path = os.path.dirname(os.path.realpath(__file__))

# 加载配置文件
def load_config():
    global devices, subject_var
    if os.path.exists(config_file):
        with open(config_file, 'r', encoding='utf-8') as file:
            config = json.load(file)
            devices = config.get("devices", {})
            subject_var.set(config.get("subject", ""))  # 加载保存的任教科目
    else:
        devices = {}  # 如果文件不存在，初始化为空

# 保存配置文件
def save_config():
    config = {
        "devices": devices,
        "subject": subject_var.get()  # 保存任教科目
    }
    with open(config_file, 'w', encoding='utf-8') as file:
        json.dump(config, file, ensure_ascii=False, indent=4)

# 检查设备是否在线
def is_device_online(ip, port=12345):
    try:
        with socket.create_connection((ip, port), timeout=2):
            return True
    except OSError:
        return False

# 添加教室端
def add_classroom():
    classroom_name = simpledialog.askstring("添加教室端", "请输入教室端名称:")
    if not classroom_name:
        return

    file_path = filedialog.askopenfilename(
        title="选择学生名单文件",
        filetypes=[("Excel文件", "*.xlsx *.xls")]
    )
    
    if not file_path:
        messagebox.showwarning("警告", "未选择学生名单文件")
        return

    try:
        ext = os.path.splitext(file_path)[-1]
        new_file_path = os.path.join(base_path, classroom_name + ext)
        shutil.copy(file_path, new_file_path)

        classroom_ip = simpledialog.askstring("添加教室端", f"请输入 {classroom_name} 教室的IP地址:")
        if not classroom_ip:
            return

        devices[classroom_name] = classroom_ip
        save_config()
        update_classroom_list()
        messagebox.showinfo("成功", f"成功添加 {classroom_name} 教室端")
    except Exception as e:
        messagebox.showerror("错误", f"添加教室端时出现错误: {str(e)}")

# 删除教室端
def delete_classroom():
    selected_classroom = classroom_listbox.get(tk.ACTIVE)
    if not selected_classroom:
        messagebox.showwarning("警告", "请选择要删除的教室端")
        return

    classroom_name = selected_classroom.split(":")[0]
    if classroom_name in devices:
        del devices[classroom_name]
        save_config()
        update_classroom_list()
        messagebox.showinfo("成功", f"{classroom_name} 已成功删除")
    else:
        messagebox.showerror("错误", "教室端不存在")

# 编辑教室端
def edit_classroom():
    selected_classroom = classroom_listbox.get(tk.ACTIVE)
    if not selected_classroom:
        messagebox.showwarning("警告", "请选择要编辑的教室端")
        return

    classroom_name = selected_classroom.split(":")[0]
    if classroom_name in devices:
        new_name = simpledialog.askstring("编辑教室端", "请输入新的教室端名称:", initialvalue=classroom_name)
        if not new_name:
            return

        new_ip = simpledialog.askstring("编辑教室端", "请输入新的教室端IP地址:", initialvalue=devices[classroom_name])
        if not new_ip:
            return

        devices.pop(classroom_name)  # 删除旧名称的记录
        devices[new_name] = new_ip
        save_config()
        update_classroom_list()
        messagebox.showinfo("成功", f"{classroom_name} 已成功修改为 {new_name}")
    else:
        messagebox.showerror("错误", "教室端不存在")

# 更新教室列表下拉框和列表框
def update_classroom_list():
    global classroom_menu
    classroom_listbox.delete(0, tk.END)
    for classroom_name, ip_address in devices.items():
        classroom_listbox.insert(tk.END, f"{classroom_name}: {ip_address}")
    if classroom_menu:
        classroom_menu['values'] = [name for name in devices.keys() if name not in ['devices', 'subject']]

# 刷新在线班级教室端列表
def refresh_classroom_list():
    classroom_listbox.delete(0, tk.END)
    for classroom_name, ip_address in devices.items():
        if is_device_online(ip_address):
            classroom_listbox.insert(tk.END, f"{classroom_name}: {ip_address} (在线)")
        else:
            classroom_listbox.insert(tk.END, f"{classroom_name}: {ip_address} (离线)")

# 选择教室端并开始叫号
def select_classroom():
    selected_classroom = classroom_menu.get()
    
    if selected_classroom not in devices or selected_classroom == 'devices' or selected_classroom == 'subject':
        messagebox.showerror("错误", "请选择一个有效的教室端")
        return

    ip_address = devices[selected_classroom]
    
    if not is_device_online(ip_address):
        messagebox.showwarning("连接错误", f"{selected_classroom} 教室端未连接")
        return

    student_input = simpledialog.askstring("输入", "请输入座号或姓名:")
    if not student_input:
        return
    
    class_file_path = os.path.join(base_path, f"{selected_classroom}.xlsx")
    
    try:
        df = pd.read_excel(class_file_path)
        df = df.dropna()

        if student_input.isdigit():
            row = df[df['A'] == int(student_input)]
        else:
            row = df[df['B'] == student_input]

        if row.empty:
            messagebox.showwarning("未找到", "未找到匹配的座号或姓名")
            return

        seat_number = row.iloc[0]['A']
        student_name = row.iloc[0]['B']

        send_data_to_classroom(ip_address, seat_number, student_name)

        messagebox.showinfo("成功", f"叫号成功！座号: {seat_number}, 姓名: {student_name}")

    except Exception as e:
        messagebox.showerror("错误", f"读取Excel文件时出现错误: {str(e)}")

# 向教室端发送数据
def send_data_to_classroom(ip, seat_number, student_name):
    try:
        with socket.create_connection((ip, 12345), timeout=2) as s:
            data = f"{seat_number},{student_name},{subject_var.get()}"
            s.sendall(data.encode())
    except OSError as e:
        messagebox.showerror("连接错误", f"无法连接到教室端: {str(e)}")

# 主界面布局
main_frame = tk.Frame(root)
main_frame.pack(fill=tk.BOTH, expand=True)

left_frame = tk.Frame(main_frame)
left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=20, pady=20)

right_frame = tk.Frame(main_frame)
right_frame.pack(side=tk.LEFT, fill=tk.BOTH, padx=20, pady=20)

# 左侧教室端管理部分
classroom_list_label = tk.Label(left_frame, text="已添加的教室端列表:")
classroom_list_label.pack(padx=10, pady=10)

classroom_listbox = tk.Listbox(left_frame, width=30)
classroom_listbox.pack(padx=10, pady=10)

refresh_button = tk.Button(left_frame, text="刷新列表", command=refresh_classroom_list)
refresh_button.pack(pady=10)

edit_button = tk.Button(left_frame, text="编辑教室端", command=edit_classroom)
edit_button.pack(pady=10)

delete_button = tk.Button(left_frame, text="删除教室端", command=delete_classroom)
delete_button.pack(pady=10)

# 右侧任教科目和叫号部分
subject_label = tk.Label(right_frame, text="请输入您的任教科目:")
subject_label.pack(padx=10, pady=10)

subject_entry = tk.Entry(right_frame, textvariable=subject_var)
subject_entry.pack(padx=10, pady=10)

classroom_menu_label = tk.Label(right_frame, text="选择班级:")
classroom_menu_label.pack(padx=10, pady=10)

classroom_menu = ttk.Combobox(right_frame)
classroom_menu.pack(padx=10, pady=10)

call_button = tk.Button(right_frame, text="开始叫号", command=select_classroom)
call_button.pack(padx=20, pady=20)

add_classroom_button = tk.Button(right_frame, text="添加教室端", command=add_classroom)
add_classroom_button.pack(pady=10)

# 加载配置并初始化教室端列表
load_config()
update_classroom_list()

# 启动主循环
root.mainloop()
