import psutil
import pyautogui
import time
import threading
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import pygetwindow as gw

# 检查是否有Microsoft Office程序在前台运行
def is_office_running():
    # 获取当前前台窗口的标题
    try:
        active_window = gw.getActiveWindow()
        if active_window is None:
            return False

        window_title = active_window.title.lower()

        # 检查当前窗口标题是否包含Microsoft Office程序的名称
        office_keywords = ['word', 'excel', 'powerpoint']
        return any(keyword in window_title for keyword in office_keywords)

    except Exception as e:
        print(f"错误: {e}")
        return False

# 模拟按下Ctrl + S
def save_document():
    pyautogui.hotkey('ctrl', 's')

# 更新GUI中的状态信息
def update_status(status_text, label):
    label.config(text=status_text)

# 获取当前日期和时间，格式化为字符串
def get_current_datetime():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# 自动保存逻辑
def auto_save(save_interval):
    last_save_time = time.time()
    while running:
        # 每秒检测一次是否有Microsoft Office程序在前台运行
        if is_office_running():
            update_status("检测到Microsoft Office程序正在台前运行...", status_label_1)
        else:
            update_status("检测到Microsoft Office程序没有在台前运行", status_label_1)

        # 每save_interval秒执行一次保存
        if time.time() - last_save_time >= save_interval:
            if is_office_running():
                save_document()
                current_time = get_current_datetime()
                update_status(f"文档已成功于{current_time}保存！", status_label_2)
            last_save_time = time.time()

        time.sleep(1)  # 每秒检查一次

# 启动和停止自动保存的控制逻辑
def toggle_auto_save():
    global running, save_interval
    if running:
        running = False
        start_button.config(text="开始自动保存")
        update_status("自动保存已停止。", status_label_1)
        update_status("", status_label_2)  # 清空成功保存信息
    else:
        # 获取用户输入的保存间隔（秒）
        try:
            save_interval = int(save_interval_entry.get())
            if save_interval <= 0:
                raise ValueError("保存间隔必须是正整数。")
        except ValueError as e:
            messagebox.showerror("无效输入", f"请输入一个有效的正整数：{str(e)}")
            return

        running = True
        start_button.config(text="停止自动保存")
        # 在后台启动自动保存线程
        threading.Thread(target=auto_save, args=(save_interval,), daemon=True).start()
        update_status(f"自动保存已启动，间隔 {save_interval} 秒", status_label_1)
        update_status("", status_label_2)  # 清空成功保存信息

# 创建GUI
root = tk.Tk()
root.title("自动保存Microsoft Office文档")
root.geometry("600x700")  # 设置窗口大小为800x600

# 设置初始状态
running = False
save_interval = 30  # 默认保存间隔为30秒

# 使用更现代的字体和样式
font_large = ("Segoe UI", 14, "bold")  # 标题和状态标签的大号加粗字体
font_small = ("Segoe UI", 12)  # 按钮的中等字体
font_footer = ("Segoe UI", 10)  # 用于版本号和作者的较小字体

# 创建一个Frame，用于上下显示两个状态标签
frame_status = tk.Frame(root, bg="#f0f0f0")
frame_status.pack(pady=15, expand=True)

# 状态显示标签1（检测是否运行）
status_label_1 = tk.Label(frame_status, text="点击按钮启动自动保存", font=font_large, height=3, anchor="center", width=40, bg="#f0f0f0")
status_label_1.pack(pady=5)  # 减少间距

# 状态显示标签2（保存成功信息）
status_label_2 = tk.Label(frame_status, text="", font=font_large, height=3, anchor="center", width=40, bg="#f0f0f0")
status_label_2.pack(pady=5)  # 减少间距

# 输入保存间隔的标签和输入框
interval_label = tk.Label(root, text="请输入保存间隔（秒）:", font=font_small, bg="#f0f0f0")
interval_label.pack(pady=5)

save_interval_entry = tk.Entry(root, font=font_small)
save_interval_entry.insert(0, str(save_interval))  # 设置默认值
save_interval_entry.pack(pady=5)

# 启动/停止按钮
start_button = tk.Button(root, text="开始自动保存", font=font_small, command=toggle_auto_save, width=20, height=2, bg="#4CAF50", fg="white", relief="flat")
start_button.pack(pady=10)

# 退出按钮
exit_button = tk.Button(root, text="退出", font=font_small, command=root.quit, width=20, height=2, bg="#F44336", fg="white", relief="flat")
exit_button.pack(pady=10)

# 创建底部显示版本号和作者的Frame
frame_footer = tk.Frame(root, bg="#f0f0f0")
frame_footer.pack(side="bottom", fill="x", padx=10, pady=5)  # 将整个footer区域放到窗口底部

# 右下角的版本号和作者标签
version_label = tk.Label(frame_footer, text="版本号: 1.0.0", font=font_footer, bg="#f0f0f0", anchor="se")
version_label.pack(side="right", padx=10)

author_label = tk.Label(frame_footer, text="作者: Def", font=font_footer, bg="#f0f0f0", anchor="se")
author_label.pack(side="right", padx=10)

# 设置背景色和主界面的边距
root.config(bg="#f0f0f0")

# 启动GUI事件循环
root.mainloop()
