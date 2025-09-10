import ctypes
import tkinter as tk
from tkinter import messagebox
import threading
import time
from PIL import Image, ImageDraw
import pystray
from pystray import MenuItem as item
import sys
import logging
import os
from datetime import datetime, timedelta
import shutil
import subprocess
import platform

# 快捷方式依赖
try:
    import pythoncom
    from win32com.client import Dispatch
except ImportError:
    print("请先安装必要依赖库:")
    print("pip install pywin32")
    sys.exit(1)

# 单实例限制依赖
try:
    import win32event
    import win32api
    import winerror
except ImportError:
    print("请先安装必要的依赖库:")
    print("pip install pywin32")
    sys.exit(1)

# ===== 单实例运行限制 =====
mutex_name = "HydrationReminderSingletonMutex"
hMutex = win32event.CreateMutex(None, False, mutex_name)
last_error = win32api.GetLastError()
if last_error == winerror.ERROR_ALREADY_EXISTS:
    ctypes.windll.user32.MessageBoxW(
        0, "喝水提醒助手已在运行，将切换到正在运行的程序。", "提示", 0
    )
    sys.exit(0)
# ===== 单实例限制结束 =====


class HydrationReminder:
    def __init__(self):
        self.root = None
        self.tray_icon = None
        self.reminder_timer = None
        self.is_running = False
        self.reminder_interval = 30
        self.log_dir = "HydrationReminderLog"

        self.setup_logging()
        self.auto_clean_old_logs()
        self.logger.info("程序启动")

        self.create_shortcut()
        self.setup_tray()

    # ===== 新增：创建桌面快捷方式方法 =====
    def create_shortcut(self):
        try:
            desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
            shortcut_path = os.path.join(desktop_path, "HydrationReminder.lnk")

            if os.path.exists(shortcut_path):
                self.logger.info("桌面已存在 HydrationReminder 快捷方式，跳过创建")
                return

            target = sys.executable  # 当前 exe 路径
            if target.lower().endswith("python.exe") or target.lower().endswith("pythonw.exe"):
                target = os.path.abspath(sys.argv[0])  # 脚本路径

            pythoncom.CoInitialize()
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortcut(shortcut_path)
            shortcut.TargetPath = target
            shortcut.WorkingDirectory = os.path.dirname(target)
            shortcut.IconLocation = target
            shortcut.Save()
            self.logger.info(f"已在桌面创建 HydrationReminder 快捷方式: {shortcut_path}")
        except Exception as e:
            self.logger.error(f"创建快捷方式失败: {str(e)}")

    def setup_logging(self):
        if not os.path.exists(self.log_dir):
            os.makedirs(self.log_dir)
        now = datetime.now()
        log_file_name = now.strftime("%Y-%m-%d_%H-%M-%S") + ".log"
        self.current_log_path = os.path.join(self.log_dir, log_file_name)

        self.logger = logging.getLogger("HydrationReminder")
        self.logger.setLevel(logging.INFO)
        if not self.logger.handlers:
            formatter = logging.Formatter('%(asctime)s [%(levelname)s] %(message)s')
            file_handler = logging.FileHandler(self.current_log_path, mode='a', encoding='utf-8')
            file_handler.setFormatter(formatter)
            self.logger.addHandler(file_handler)

    def auto_clean_old_logs(self):
        if not os.path.exists(self.log_dir):
            self.logger.info("日志目录不存在，无需清理旧日志")
            return

        one_month_ago = datetime.now() - timedelta(days=30)
        deleted_count = 0
        for filename in os.listdir(self.log_dir):
            file_path = os.path.join(self.log_dir, filename)
            if filename.endswith(".log") and os.path.isfile(file_path):
                try:
                    file_datetime_str = filename.split(".")[0]
                    file_datetime = datetime.strptime(file_datetime_str, "%Y-%m-%d_%H-%M-%S")
                    if file_datetime < one_month_ago and file_path != self.current_log_path:
                        os.remove(file_path)
                        deleted_count += 1
                        self.logger.info(f"自动清理旧日志：{filename}")
                except Exception as e:
                    self.logger.warning(f"跳过异常日志文件 {filename}，原因：{str(e)}")

        self.logger.info(f"自动清理完成，共删除{deleted_count}个1个月前的日志文件")

    def create_icon_image(self):
        width = height = 64
        image = Image.new('RGBA', (width, height), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        draw.ellipse([16, 20, 48, 52], fill=(70, 130, 220, 255))
        draw.polygon([(32, 12), (24, 28), (40, 28)], fill=(70, 130, 220, 255))
        return image

    def setup_tray(self):
        icon_image = self.create_icon_image()
        menu = pystray.Menu(
            item('打开主界面', self.show_main_window),
            item('开始提醒', self.start_reminder, checked=lambda _: self.is_running),
            item('停止提醒', self.stop_reminder),
            pystray.Menu.SEPARATOR,
            item('打开日志文件夹', self.open_log_folder),
            item('删除历史日志', self.delete_history_logs),
            item('退出程序', self.quit_application)
        )
        self.tray_icon = pystray.Icon("HydrationReminder", icon_image, "喝水提醒助手", menu)
        self.logger.info("系统托盘初始化完成")

    def show_main_window(self, icon=None, item=None):
        if self.root is None or not self.root.winfo_exists():
            self.create_main_window()
            self.logger.info("主界面创建完成")
        else:
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
            self.logger.info("主界面显示")

    def create_main_window(self):
        self.root = tk.Tk()
        self.root.title("喝水提醒助手")
        self.root.geometry("400x500")
        self.root.configure(bg='white')
        self.root.resizable(False, False)
        self.center_window()

        main_frame = tk.Frame(self.root, bg='white', padx=40, pady=40)
        main_frame.pack(fill='both', expand=True)

        tk.Label(main_frame, text="💧 喝水提醒助手", font=("Arial", 24, "bold"),
                 fg='#1d1d1f', bg='white').pack(pady=(0, 8))
        tk.Label(main_frame, text="保持健康的饮水习惯", font=("Arial", 13),
                 fg='#86868b', bg='white').pack(pady=(0, 20))

        self.status_label = tk.Label(main_frame, text="● 提醒服务已停止",
                                     font=("Arial", 16), fg='#ff3b30', bg='#f5f5f7',
                                     relief='solid', bd=1, padx=10, pady=10)
        self.status_label.pack(fill='x', pady=(0, 30))

        tk.Label(main_frame, text="提醒间隔设置", font=("Arial", 18, "bold"),
                 fg='#1d1d1f', bg='white').pack(pady=(0, 4))
        tk.Label(main_frame, text="设置多长时间提醒您喝水一次", font=("Arial", 13),
                 fg='#86868b', bg='white').pack(pady=(0, 16))

        input_frame = tk.Frame(main_frame, bg='white')
        input_frame.pack(pady=(0, 8))
        self.time_var = tk.StringVar(value=str(self.reminder_interval))
        self.time_entry = tk.Entry(input_frame, textvariable=self.time_var,
                                   font=("Arial", 14), width=6, justify='center',
                                   relief='solid', bd=2)
        self.time_entry.configure(insertbackground='#007aff')
        label_minute = tk.Label(input_frame, text="分钟", font=("Arial", 14),
                                fg='#1d1d1f', bg='white')
        self.time_entry.pack(side='left')
        label_minute.pack(side='left', padx=(12, 0))

        self.countdown_var = tk.StringVar(value="")
        tk.Label(main_frame, textvariable=self.countdown_var, font=("Arial", 14),
                 fg='#1d1d1f', bg='white').pack(pady=(0, 20))

        self.start_btn = tk.Button(main_frame, text="开始提醒",
                                   font=("Arial", 14, "bold"), fg='white',
                                   bg='#007aff', activebackground='#0056cc',
                                   activeforeground='white', relief='flat', bd=0,
                                   padx=20, pady=6, cursor='hand2',
                                   command=self.toggle_reminder)
        self.start_btn.pack(fill='x', pady=(0, 12))

        close_btn = tk.Button(main_frame, text="关闭程序",
                              font=("Arial", 12), fg='#1d1d1f',
                              bg='#e6e6e6', activebackground='#d9d9d9',
                              activeforeground='#1d1d1f', relief='flat', bd=0,
                              padx=20, pady=4, cursor='hand2',
                              command=self.quit_application)
        close_btn.pack(fill='x')

        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)
        self.update_status_display()
        self.logger.info("主界面显示完成")
        self.root.mainloop()

    def center_window(self):
        self.root.update_idletasks()
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - 400) // 2
        y = (screen_height - 500) // 2
        self.root.geometry(f"400x500+{x}+{y}")

    def hide_window(self):
        if self.root:
            self.root.withdraw()
            self.logger.info("主界面隐藏")

    def update_status_display(self):
        if self.root and self.root.winfo_exists():
            if self.is_running:
                self.status_label.config(
                    text=f"● 提醒服务运行中 (每{self.reminder_interval}分钟)", fg='#30d158')
                self.start_btn.config(text="停止提醒", bg='#ff3b30', activebackground='#cc2e24')
                self.time_entry.config(state='disabled')
            else:
                self.status_label.config(text="● 提醒服务已停止", fg='#ff3b30')
                self.start_btn.config(text="开始提醒", bg='#007aff', activebackground='#0056cc')
                self.time_entry.config(state='normal')

    def toggle_reminder(self):
        if self.is_running:
            self.stop_reminder()
        else:
            try:
                interval = int(self.time_var.get())
                if interval <= 0:
                    self.logger.error("用户输入错误时间间隔 <=0")
                    messagebox.showerror("错误", "请输入大于0的时间间隔")
                    self.logger.info("用户点击了错误弹窗的 OK 按钮")
                    return
                self.reminder_interval = interval
                self.logger.info(f"用户设置提醒间隔: {self.reminder_interval}分钟")
                self.start_reminder()
            except ValueError:
                self.logger.error("用户输入非法时间")
                messagebox.showerror("错误", "请输入有效的数字")
                self.logger.info("用户点击了错误弹窗的 OK 按钮")

    def start_reminder(self, icon=None, item=None):
        if not self.is_running:
            self.is_running = True
            self.schedule_next_reminder()
            self.update_status_display()
            self.logger.info(f"提醒服务启动: {self.reminder_interval}分钟一次")

    def stop_reminder(self, icon=None, item=None):
        if self.is_running:
            self.is_running = False
            if self.reminder_timer:
                self.reminder_timer.cancel()
            self.countdown_var.set("")
            self.update_status_display()
            self.logger.info("提醒服务停止")

    def schedule_next_reminder(self):
        if self.is_running:
            next_time = time.time() + self.reminder_interval * 60
            self.reminder_timer = threading.Timer(self.reminder_interval * 60, self.show_reminder)
            self.reminder_timer.finished_at = next_time
            self.reminder_timer.start()
            if self.root and self.root.winfo_exists():
                self.update_countdown()

    def update_countdown(self):
        if self.is_running and self.reminder_timer:
            remaining = int(self.reminder_timer.finished_at - time.time())
            if remaining < 0:
                remaining = 0
            minutes, seconds = divmod(remaining, 60)
            self.countdown_var.set(f"距离下一次提醒: {minutes:02d}:{seconds:02d}")
            self.root.after(1000, self.update_countdown)
        else:
            self.countdown_var.set("")

    def show_reminder(self):
        if self.is_running:
            top = tk.Toplevel()
            top.attributes('-topmost', True)
            top.withdraw()
            messagebox.showinfo(
                "喝水提醒助手",
                f"该喝水啦！💧\n\n保持健康的饮水习惯。\n下次提醒将在{self.reminder_interval}分钟后。",
                parent=top
            )
            self.logger.info("用户点击了提醒弹窗的 OK 按钮")
            top.destroy()
            self.schedule_next_reminder()

    def open_log_folder(self, icon=None, item=None):
        try:
            if not os.path.exists(self.log_dir):
                os.makedirs(self.log_dir)
                self.logger.info("日志目录不存在，已自动创建")

            system = platform.system()
            if system == "Windows":
                os.startfile(self.log_dir)
            elif system == "Darwin":
                subprocess.run(['open', self.log_dir])
            elif system == "Linux":
                subprocess.run(['xdg-open', self.log_dir])
            else:
                raise Exception(f"不支持的操作系统: {system}")

            self.logger.info(f"用户打开日志文件夹: {self.log_dir}")
        except Exception as e:
            error_msg = f"无法打开日志文件夹：{str(e)}"
            messagebox.showerror("错误", error_msg)
            self.logger.error(f"打开日志文件夹失败: {str(e)}")

    def delete_history_logs(self, icon=None, item=None):
        if not os.path.exists(self.log_dir):
            messagebox.showinfo("提示", "日志目录不存在，无历史日志可删除")
            self.logger.info("用户触发删除历史日志，但日志目录不存在")
            return

        history_logs = []
        for filename in os.listdir(self.log_dir):
            file_path = os.path.join(self.log_dir, filename)
            if filename.endswith(".log") and os.path.isfile(file_path) and file_path != self.current_log_path:
                history_logs.append(file_path)

        if not history_logs:
            messagebox.showinfo("提示", "当前无历史日志可删除（仅保留当前运行日志）")
            self.logger.info("用户触发删除历史日志，但无符合条件的历史日志")
            return

        confirm = messagebox.askyesno(
            "确认删除",
            f"即将删除{len(history_logs)}个历史日志文件（当前运行日志不会被删除）\n\n是否继续？"
        )
        if not confirm:
            self.logger.info("用户取消删除历史日志操作")
            return

        deleted_count = 0
        failed_count = 0
        for log_path in history_logs:
            try:
                os.remove(log_path)
                deleted_count += 1
                self.logger.info(f"手动删除历史日志：{os.path.basename(log_path)}")
            except Exception as e:
                failed_count += 1
                self.logger.error(f"删除历史日志失败：{os.path.basename(log_path)}，原因：{str(e)}")

        result_msg = f"删除完成！\n\n成功删除：{deleted_count}个\n删除失败：{failed_count}个"
        messagebox.showinfo("删除结果", result_msg)
        self.logger.info(f"手动删除历史日志操作完成，成功{deleted_count}个，失败{failed_count}个")

    def quit_application(self, icon=None, item=None):
        self.stop_reminder()
        if self.root and self.root.winfo_exists():
            self.root.quit()
        if self.tray_icon:
            self.tray_icon.stop()
        self.logger.info("程序退出")
        sys.exit(0)

    def run(self):
        tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
        tray_thread.start()
        self.show_main_window()


if __name__ == "__main__":
    app = HydrationReminder()
    app.run()
