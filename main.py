import ctypes
import os
import shutil
import subprocess
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import random
import ttkbootstrap as ttk


def create_symlink(source, link_name):
    """
    创建一个符号链接。

    参数:
    source (str): 源文件或目录的路径。
    link_name (str): 符号链接的路径。
    """
    # 检查Windows版本是否支持mklink
    if os.name != 'nt':
        raise OSError("这个程序只能在Windows上运行")

    # 构建命令
    command = f'mklink /J "{link_name}" "{source}"'

    # 执行命令
    try:
        subprocess.run(command, shell=True, check=True)
        print(f"符号链接已创建: {link_name} -> {source}")
    except subprocess.CalledProcessError as e:
        print(f"创建符号链接时出错: {e}")


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


if is_admin():
    def get_size(start_path):
        """ 计算给定路径的总大小（递归计算所有文件的大小）。 """
        total_size = 0
        for dirpath, dirnames, filenames in os.walk(start_path):
            for f in filenames:
                fp = os.path.join(dirpath, f)
                if os.path.exists(fp):
                    total_size += os.path.getsize(fp)
        return total_size


    def update_progress_bar(start_size, dest_path):
        """ 更新进度条的函数，它计算目标文件夹的当前大小并更新进度。 """
        while True:
            current_size = get_size(dest_path)
            progress = (current_size / start_size) * 100
            progress_bar['value'] = progress
            app.update_idletasks()
            if progress >= 100:
                break
            time.sleep(0.1)


    def choose_source():
        folder_selected = filedialog.askdirectory()
        source_entry.delete(0, tk.END)
        source_entry.insert(0, folder_selected)


    def choose_destination():
        folder_selected = filedialog.askdirectory()
        destination_entry.delete(0, tk.END)
        destination_entry.insert(0, folder_selected)


    def start_move():
        source = source_entry.get()
        folder_name = os.path.basename(source)
        destination = destination_entry.get()

        if not source or not destination:
            messagebox.showerror("错误", "请指定源路径和目标路径")
            return

        start_size = get_size(source)
        threading.Thread(target=move_folder, args=(source, destination, f"{destination}/{folder_name}"), daemon=True).start()
        threading.Thread(target=update_progress_bar, args=(start_size, f"{destination}/{folder_name}"),
                         daemon=True).start()


    def move_folder(source, destination, destined):
        try:
            shutil.move(source, destination)
            create_symlink(destined, source)
            messagebox.showinfo("完成", "文件夹移动完成")
        except Exception as e:
            messagebox.showerror("错误", str(e))


    def closeWindow():
        result = tk.messagebox.askyesno("请确定", "为了您的数据安全，如果您正在移动文件，请务必不要退出程序！！！")
        if result:
            app.destroy()
        else:
            pass

    tem = ["superhero", "vapor", "cyborg", "solar", "cosmo", "flatly", "journal", "litera", "minty", "pulse", "morph"]
    app = ttk.Window(themename=tem[random.randint(0, 10)])
    app.protocol('WM_DELETE_WINDOW', closeWindow)
    source_entry = ttk.Entry(app, width=32)
    source_entry.pack()
    ttk.Button(app, text="选择源文件夹", command=choose_source, bootstyle="outline", width=30).pack()

    destination_entry = ttk.Entry(app, width=32)
    destination_entry.pack()
    ttk.Button(app, text="选择目标文件夹", command=choose_destination, bootstyle="outline", width=30).pack()

    ttk.Button(app, text="开始迁移并设置符号链接", command=start_move, bootstyle="outline", width=30).pack()

    progress_bar = ttk.Progressbar(app, mode='determinate', maximum=100, length=300, bootstyle="striped")
    progress_bar.pack()

    app.mainloop()

else:
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)