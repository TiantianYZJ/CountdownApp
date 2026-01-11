import datetime
import json
import os
import math
import random
import re
import requests
import sqlite3
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from win10toast import ToastNotifier
import win32api
import win32con
import winerror
import win32event
import win32gui

class CountdownApp:
    notepad_count = 0  # 便签计数类变量

    def __init__(self):
        # 创建主窗口
        self.root = tk.Tk()
        self.root.title("桌面时钟倒计时组件")
        
        # 设置窗口背景为黑色且透明度
        self.root.configure(bg='black')
        self.root.attributes('-alpha', 0.65)
        self.root.geometry("950x610")

        # 设置窗口最小尺寸
        self.root.minsize(550, 200)
        
        # 初始化数据库并加载设置
        self.init_database()
        self.load_settings()
        
        # 创建ttk样式
        self.setup_styles(base_font_size=0)
        
        # 创建主框架
        self.main_frame = ttk.Frame(self.root, style="Black.TFrame", padding=(20, 20, 20, 20))
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建左右分栏框架
        self.left_frame = ttk.Frame(self.main_frame, style="Black.TFrame")
        self.left_frame.pack(side=tk.LEFT, padx=(0, 20), fill=tk.BOTH, expand=True)
        
        self.right_frame = ttk.Frame(self.main_frame, style="Black.TFrame")
        self.right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # 左侧框架
        # 创建标签显示当前时间
        # self.time_label = ttk.Label(self.left_frame, font=('Microsoft YaHei UI', 85, 'bold'), style="Time.TLabel")
        self.time_label = ttk.Label(self.left_frame, style="Time.TLabel")
        self.time_label.pack(pady=(0, 5))
        
        # 创建标签显示当前日期
        # self.date_label = ttk.Label(self.left_frame, font=('Microsoft YaHei UI', 28), style="Date.TLabel")
        self.date_label = ttk.Label(self.left_frame, style="Date.TLabel")
        self.date_label.pack(pady=(5, 10))
        
        # 创建分隔线
        separator = ttk.Separator(self.left_frame, orient='horizontal', style="Separator.TSeparator")
        separator.pack(fill='x', pady=(10, 10))
        
        # 创建标签显示倒计时
        # self.countdown_label = ttk.Label(self.left_frame, font=('Microsoft YaHei UI', 32, 'bold'), style="Countdown.TLabel")
        self.countdown_label = ttk.Label(self.left_frame, style="Countdown.TLabel")
        self.countdown_label.pack(pady=(10, 10))
        
        # 创建分隔线
        separator2 = ttk.Separator(self.left_frame, orient='horizontal', style="Separator.TSeparator")
        separator2.pack(fill='x', pady=(10, 10))
        
        # 创建标签显示名言
        # self.quote_label = ttk.Label(self.left_frame, font=('Microsoft YaHei UI', 18, 'italic'), style="Quote.TLabel",
        #                             wraplength=500, justify='center')
        self.quote_label = ttk.Label(self.left_frame, style="Quote.TLabel",
                            wraplength=500, justify='center')
        self.quote_label.pack(pady=(5, 5))
        
        # 创建标签显示名言来源
        # self.from_label = ttk.Label(self.left_frame, font=('Microsoft YaHei UI', 14), style="Source.TLabel", anchor='e')
        self.from_label = ttk.Label(self.left_frame, style="Source.TLabel", anchor='e')
        self.from_label.pack(fill='x', pady=(5, 10))

        # 在右侧框架中创建课程表标题
        # self.schedule_title = ttk.Label(self.right_frame, font=('Microsoft YaHei UI', 20, 'bold'), style="Title.TLabel", text='今日课程')
        self.schedule_title = ttk.Label(self.right_frame, style="Title.TLabel", text='今日课程')
        self.schedule_title.pack(pady=(5, 5), anchor='w')
        
        # 创建分隔线
        schedule_separator = ttk.Separator(self.right_frame, orient='horizontal', style="Separator.TSeparator")
        schedule_separator.pack(fill='x', pady=(0, 5))
        
        # 创建课程表内容框架（Canvas）
        self.schedule_canvas = tk.Canvas(self.right_frame, bg='black', highlightthickness=0)
        self.schedule_canvas.pack(fill=tk.BOTH, expand=True)
        
        # 创建课程表内容容器
        self.schedule_content = ttk.Frame(self.schedule_canvas, style="Black.TFrame")
        self.schedule_content.bind(
            "<Configure>",
            lambda e: self.schedule_canvas.configure(
                scrollregion=self.schedule_canvas.bbox("all")
            )
        )
        
        # 保存内容窗口的ID并绑定事件
        self.schedule_content_id = self.schedule_canvas.create_window((0, 0), window=self.schedule_content, anchor="nw")
        
        # 绑定Canvas的Configure事件，使内容容器宽度与Canvas一致
        def on_canvas_configure(event):
            self.schedule_canvas.itemconfig(self.schedule_content_id, width=event.width)
        
        self.schedule_canvas.bind("<Configure>", on_canvas_configure)

        # 绑定窗口大小变化事件
        self.root.bind("<Configure>", self.update_font_sizes)
        # 初始调用一次更新字体大小
        self.update_font_sizes()
        
        # 创建状态标签容器（用于在display_todays_schedule中添加
        self.status_label = None
        
        # 鼠标事件变量
        self.drag_start_time = None
        self.is_dragging = False
        self.drag_threshold = 10  # 拖动阈值（px）
        self.click_threshold = 300  # 点击阈值（ms）

        # 添加显示模式标志
        self.current_display_mode = "word"  # "quote" 或 "word"
        
        # 绑定鼠标事件
        self.bind_mouse_events()
        
        # 绑定右键菜单
        self.create_context_menu()
        
        # 加载课程表
        self.load_schedule()

        # 初始化通知器
        self.update_status_text()

        # 添加小窗口位置跟踪变量
        self.mini_window_position_set = False
        
        # 初始化PPT全屏检测小窗口
        self.create_mini_window()

        # 更新倒计时
        self.update_countdown()
        
        # 获取并显示第一条名言
        self.refresh_content(None)
        
        # 设置窗口位置
        self.set_window_position()

        # 属性初始化
        self.last_morning_notification = None
        self.last_evening_notification = None

        # 便签窗口管理
        self.notepad_count = 0
        self.notepads = []  # 存储所有便签窗口实例
    
    def update_font_sizes(self, event=None):
        """根据窗口尺寸更新所有标签的字体大小"""
        # 获取窗口尺寸
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        
        # 计算基础字体大小（基于窗口宽度和高度的平均值）
        base_font_size = min(max(int((height) / 47), 12), 20)
        # print(f"Updating font sizes with height:{height}, base_font_size: {base_font_size}")
        
        # 更新所有样式的字体大小
        self.setup_styles(base_font_size)
        
        # 更新名言标签的wraplength以适应新的窗口宽度
        if self.quote_label:
            self.quote_label.config(wraplength=width // 2)

    def setup_styles(self, base_font_size):
        """设置ttk样式"""
        style = ttk.Style()

        # 如果没有提供base_font_size，计算默认值
        if base_font_size == 0:
            width = self.root.winfo_width()
            height = self.root.winfo_height()
            base_font_size = min(max(int((height) / 47), 12), 20)
            # print(f"Calculated height:{height}, base_font_size: {base_font_size}")
    
        # 黑色背景框架
        style.configure("Black.TFrame", background="black")
        style.configure("Black.TLabelframe", background="black", foreground="white")
        style.configure("Black.TLabelframe.Label", background="black", foreground="white")
        
        # 时间标签
        # style.configure("Time.TLabel", background="black", foreground="white")
        style.configure("Time.TLabel", background="black", foreground="white", 
                   font=('Tahoma', int(base_font_size * 7), 'bold'))
        
        # 日期标签
        # style.configure("Date.TLabel", background="black", foreground="white")
        style.configure("Date.TLabel", background="black", foreground="white", 
                   font=('Microsoft YaHei UI', int(base_font_size * 2.3)))
        
        # 倒计时标签（默认蓝绿色）
        # style.configure("Countdown.TLabel", background="black", foreground="#00CED1")
        # style.configure("Countdown.Blue.TLabel", background="black", foreground="#00CED1")  # 蓝绿色
        # style.configure("Countdown.Orange.TLabel", background="black", foreground="#FFA500")  # 橙色
        # style.configure("Countdown.Red.TLabel", background="black", foreground="#FF6347")  # 红色
        # style.configure("Countdown.Gold.TLabel", background="black", foreground="#FFD700")  # 金色
        # style.configure("Countdown.Purple.TLabel", background="black", foreground="#9370DB")  # 紫色
        style.configure("Countdown.TLabel", background="black", foreground="#00CED1",
                   font=('Microsoft YaHei UI', int(base_font_size * 2.7), 'bold'))
        style.configure("Countdown.Blue.TLabel", background="black", foreground="#00CED1",
                    font=('Microsoft YaHei UI', int(base_font_size * 2.7), 'bold'))
        style.configure("Countdown.Orange.TLabel", background="black", foreground="#FFA500",
                    font=('Microsoft YaHei UI', int(base_font_size * 2.7), 'bold'))
        style.configure("Countdown.Red.TLabel", background="black", foreground="#FF6347",
                    font=('Microsoft YaHei UI', int(base_font_size * 2.7), 'bold'))
        style.configure("Countdown.Gold.TLabel", background="black", foreground="#FFD700",
                    font=('Microsoft YaHei UI', int(base_font_size * 2.7), 'bold'))
        style.configure("Countdown.Purple.TLabel", background="black", foreground="#9370DB",
                    font=('Microsoft YaHei UI', int(base_font_size * 2.7), 'bold'))
        

        # 名言标签
        # style.configure("Quote.TLabel", background="black", foreground="#FFFFFF")
        style.configure("Quote.TLabel", background="black", foreground="#FFFFFF",
                   font=('Segoe UI', int(base_font_size * 1.8), 'italic'))
    
        # 名言来源标签
        # style.configure("Source.TLabel", background="black", foreground="#CCCCCC")
        style.configure("Source.TLabel", background="black", foreground="#CCCCCC",
                   font=('Segoe UI', int(base_font_size * 1.2)))
    
        # 标题标签
        # style.configure("Title.TLabel", background="black", foreground="white")
        style.configure("Title.TLabel", background="black", foreground="white",
                   font=('YouYuan', int(base_font_size * 1.7), 'bold'))
        
        # 分隔线
        style.configure("Separator.TSeparator", background="#333333")
        
        # 课程时间标签
        style.configure("ClassTime.TLabel", background="black", foreground="#98FB98", 
                        font=('YouYuan', int(base_font_size + 2)), padding=0)
        style.configure("ClassName.TLabel", background="black", foreground="white", 
                        font=('YouYuan', int(base_font_size + 2), 'bold'), padding=(100, 0))
        
        # 当前课程时间标签
        style.configure("CurrentClassTime.TLabel", background="#DDE3D2", foreground="#262626", 
                        font=('YouYuan', int(base_font_size + 2)), padding=0)
        style.configure("CurrentClassName.TLabel", background="#DDE3D2", foreground="#262626", 
                        font=('YouYuan', int(base_font_size + 2), 'bold'), padding=(100, 0))
        
        # 已结束课程标签
        style.configure("ClassTime.Gray.TLabel", background="black", foreground="#666666", 
                        font=('YouYuan', int(base_font_size + 2)), padding=0)
        style.configure("ClassName.Gray.TLabel", background="black", foreground="#888888", 
                        font=('YouYuan', int(base_font_size + 2), 'bold'), padding=(100, 0))
        
        # 状态标签
        style.configure("Status.TLabel", background="#1a1a1a", foreground="#FFFF00", 
                        font=('Microsoft YaHei UI', int(base_font_size + 2)), padding=(1, 1, 1, 1))
        # style.configure("Status.Orange.TLabel", background="#1a1a1a", foreground="#FFA500", 
        #                 font=('YouYuan', int(base_font_size + 3)), padding=0)
        # style.configure("Status.Green.TLabel", background="#1a1a1a", foreground="#32CD32", 
        #                 font=('YouYuan', int(base_font_size + 3)), padding=0)
        
        # 设置窗口专用样式
        style.configure("TButton", foreground="black", 
                        font=('Microsoft YaHei UI', 10), padding=0)
        style.configure("Main.TButton", background="black", foreground="black", 
                        font=('Microsoft YaHei UI', int(base_font_size * 0.8)), padding=0)
        style.configure("TLabel", foreground="black", 
                        font=('Microsoft YaHei UI', 10))
        style.configure("TCheckbutton", foreground="black", 
                        font=('Microsoft YaHei UI', 10))
        style.configure("TCombobox", foreground="black", 
                        font=('Microsoft YaHei UI', 10))
        style.configure("TScale", foreground="black")
        

    # 小窗口专用 - 鼠标事件处理方法
    def create_mini_window(self):
        """创建用于在PPT全屏时显示的小窗口"""
        self.mini_window = tk.Toplevel(self.root)
        self.mini_window.overrideredirect(True)  # 无边框
        self.mini_window.configure(bg='black')
        self.mini_window.attributes('-alpha', 0.5)  # 半透明
        self.mini_window.attributes('-topmost', True)  # 超级置顶
        
        # 创建显示时间的标签
        self.mini_time_label = tk.Label(self.mini_window, font=('Microsoft YaHei UI', 20), bg='black', fg='white')
        self.mini_time_label.pack(padx=5, pady=5)

        # 专用鼠标事件 - 移动功能
        self.mini_window.bind("<Button-1>", self.on_mini_window_down)
        self.mini_window.bind("<ButtonRelease-1>", self.on_mini_window_up)
        self.mini_window.bind("<B1-Motion>", self.on_mini_window_drag)
        self.mini_time_label.bind("<Button-1>", self.on_mini_window_down)
        self.mini_time_label.bind("<ButtonRelease-1>", self.on_mini_window_up)
        self.mini_time_label.bind("<B1-Motion>", self.on_mini_window_drag)
        
        # 默认隐藏小窗口
        self.mini_window.withdraw()
        
        # 初始化小窗口拖动相关属性
        self.mini_drag_start_time = None
        self.mini_drag_start_x = 0
        self.mini_drag_start_y = 0
        self.mini_is_dragging = False

    def on_mini_window_down(self, event):
        """小窗口鼠标按下事件"""
        self.mini_drag_start_time = datetime.datetime.now()
        self.mini_drag_start_x = event.x
        self.mini_drag_start_y = event.y
        self.mini_is_dragging = False
        
    def on_mini_window_up(self, event):
        """小窗口鼠标释放事件"""
        self.mini_drag_start_time = None
        self.mini_is_dragging = False

    def on_mini_window_drag(self, event):
        """处理小窗口的鼠标拖动事件"""
        if self.mini_drag_start_time is not None:
            # 计算鼠标移动距离
            dx = event.x - self.mini_drag_start_x
            dy = event.y - self.mini_drag_start_y
            distance = (dx**2 + dy**2)**0.5

            # 如果移动距离超过阈值，则认为拖动
            if distance > self.drag_threshold:
                self.mini_is_dragging = True
                self.move_mini_window(event)

    def move_mini_window(self, event):
        """移动小窗口"""
        # 计算小窗口新位置
        deltax = event.x - self.mini_drag_start_x
        deltay = event.y - self.mini_drag_start_y
        x = self.mini_window.winfo_x() + deltax
        y = self.mini_window.winfo_y() + deltay
        self.mini_window.geometry(f"+{x}+{y}")

    # 获取AppData目录
    def get_appdata_path(self):
        """获取Windows的AppData目录路径，并创建应用程序数据文件夹"""
        # 创建应用程序数据文件夹
        app_data_dir = os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'CountdownApp')
        # 确保目录存在
        if not os.path.exists(app_data_dir):
            os.makedirs(app_data_dir)
        return app_data_dir

    def init_database(self):
        """初始化本地数据库"""
        # 获取用户AppData目录
        app_data_dir = self.get_appdata_path()
        
        # 确保目录存在
        os.makedirs(app_data_dir, exist_ok=True)
        
        # 设置数据库路径
        db_path = os.path.join(app_data_dir, 'data.db')
        
        # 连接数据库
        self.conn = sqlite3.connect(db_path)
        self.cursor = self.conn.cursor()
        
        # 创建设置表
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
''')
        
        # 创建通知表（新增）
        self.cursor.execute('''
        CREATE TABLE IF NOT EXISTS notifications (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            message TEXT NOT NULL,
            hour INTEGER NOT NULL,
            minute INTEGER NOT NULL,
            enabled INTEGER DEFAULT 1,
            sound_enabled INTEGER DEFAULT 0
        )
        ''')
        
        self.conn.commit()
        
    def load_settings(self):
        """从数据库加载设置"""
        # 获取中考年份设置
        self.cursor.execute("SELECT value FROM settings WHERE key='exam_year'")
        result = self.cursor.fetchone()
        if result:
            try:
                exam_year = int(result[0])
                self.exam_date = datetime.date(exam_year, 6, 26)
            except ValueError:
                # 使用默认值
                self.exam_date = datetime.date(2026, 6, 26)
        else:
            # 默认中考年份为2026年
            self.exam_date = datetime.date(2026, 6, 26)
            # 保存默认设置
            self.cursor.execute("INSERT OR REPLACE INTO settings VALUES (?, ?)", ('exam_year', '2026'))
            self.conn.commit()
        
        # 获取通知开关设置
        self.cursor.execute("SELECT value FROM settings WHERE key='notifications_enabled'")
        result = self.cursor.fetchone()
        if result:
            self.notifications_enabled = result[0].lower() == 'true'
        else:
            # 默认启用通知
            self.notifications_enabled = True
            # 保存默认设置
            self.cursor.execute("INSERT OR REPLACE INTO settings VALUES (?, ?)", ('notifications_enabled', 'True'))
            self.conn.commit()
        
        # 获取全屏显示迷你时钟设置
        self.cursor.execute("SELECT value FROM settings WHERE key='show_mini_on_fullscreen'")
        result = self.cursor.fetchone()
        if result:
            self.show_mini_on_fullscreen = result[0].lower() == 'true'
        else:
            # 默认启用全屏显示迷你时钟
            self.show_mini_on_fullscreen = True
            # 保存默认设置
            self.cursor.execute("INSERT OR REPLACE INTO settings VALUES (?, ?)", ('show_mini_on_fullscreen', 'True'))
            self.conn.commit()
        
        # 加载通知列表（新增）
        self.load_notifications()
        
        # 初始化上次通知日期跟踪
        self.last_notification_dates = {}

    def load_notifications(self):
        """从数据库加载所有通知设置"""
        self.cursor.execute("SELECT id, title, message, hour, minute, enabled, sound_enabled FROM notifications")
        results = self.cursor.fetchall()
        
        # 如果没有通知，添加默认通知
        if not results:
            self.add_default_notifications()
            self.cursor.execute("SELECT id, title, message, hour, minute, enabled, sound_enabled FROM notifications")
            results = self.cursor.fetchall()
        
        # 存储通知列表
        self.notifications = []
        for row in results:
            self.notifications.append({
                'id': row[0],
                'title': row[1],
                'message': row[2],
                'hour': row[3],
                'minute': row[4],
                'enabled': bool(row[5]),
                'sound_enabled': bool(row[6])
            })
    
    def add_default_notifications(self):
        """添加默认通知"""
        default_notifications = [
            ('早读时间', '美好的一天开始了', 7, 30, True, False),
            ('放学了', '记得拿跳绳', 21, 30, True, False)
        ]
        
        for title, message, hour, minute, enabled, sound_enabled in default_notifications:
            self.cursor.execute(
                "INSERT INTO notifications (title, message, hour, minute, enabled, sound_enabled) VALUES (?, ?, ?, ?, ?, ?)",
                (title, message, hour, minute, 1 if enabled else 0, 1 if sound_enabled else 0)
            )
        self.conn.commit()

    def save_settings(self, exam_year, notifications_enabled, show_mini_on_fullscreen):
        """保存设置到数据库"""
        # 保存中考年份
        self.cursor.execute("INSERT OR REPLACE INTO settings VALUES (?, ?)", ('exam_year', str(exam_year)))
        # 保存通知开关
        self.cursor.execute("INSERT OR REPLACE INTO settings VALUES (?, ?)", ('notifications_enabled', str(notifications_enabled)))
        # 保存全屏显示迷你时钟设置
        self.cursor.execute("INSERT OR REPLACE INTO settings VALUES (?, ?)", ('show_mini_on_fullscreen', str(show_mini_on_fullscreen)))
        self.conn.commit()
        
        # 更新当前运行时的设置
        self.exam_date = datetime.date(exam_year, 6, 26)
        self.notifications_enabled = notifications_enabled
        self.show_mini_on_fullscreen = show_mini_on_fullscreen

    # 修改CountdownApp类中的show_settings方法
    def show_settings(self):
        """显示设置窗口"""
        settings_window = tk.Toplevel(self.root)
        settings_window.title("设置")
        settings_window.geometry("500x500")
        settings_window.resizable(False, False)
        settings_window.configure(bg="#f0f0f0")
        
        # 设置窗口居中显示
        settings_window.update_idletasks()
        width = settings_window.winfo_width()
        height = settings_window.winfo_height()
        x = (settings_window.winfo_screenwidth() // 2) - (width // 2)
        y = (settings_window.winfo_screenheight() // 2) - (height // 2)
        settings_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        # 创建主框架
        main_frame = ttk.Frame(settings_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 添加标题
        title_frame = ttk.Frame(main_frame)
        title_frame.pack(fill=tk.X, pady=(0, 20))
        title_label = ttk.Label(title_frame, text="设置", font=('Microsoft YaHei UI', 16, 'bold'))
        title_label.pack()
        # 设置ttk样式
        style = ttk.Style()
        style.configure("TLabel", font=('Microsoft YaHei UI', 10))
        style.configure("TButton", font=('Microsoft YaHei UI', 10))
        style.configure("TRadiobutton", font=('Microsoft YaHei UI', 10))
        style.configure("TCheckbutton", font=('Microsoft YaHei UI', 10))
        style.configure("TCombobox", font=('Microsoft YaHei UI', 10))
        
        # ===== 基本设置分组 =====
        basic_frame = ttk.LabelFrame(main_frame, text="基本设置", padding="10")
        basic_frame.pack(fill=tk.X, padx=5, pady=(0, 15))
        
        # 中考年份设置
        year_frame = ttk.Frame(basic_frame)
        year_frame.pack(fill=tk.X, pady=8)
        
        year_label = ttk.Label(year_frame, text="中考年份:", width=12)
        year_label.pack(side=tk.LEFT)
        
        # 创建下拉列表
        current_year = self.exam_date.year
        year_var = tk.StringVar(value=str(current_year))
        
        year_options = [str(year) for year in range(2026, 2048)]
        year_dropdown = ttk.Combobox(year_frame, textvariable=year_var, values=year_options, width=7)
        year_dropdown.pack(side=tk.LEFT, padx=5)
        year_dropdown.current(year_options.index(str(current_year)) if str(current_year) in year_options else 0)
        
        # ===== 显示设置分组 =====
        display_frame = ttk.LabelFrame(main_frame, text="显示设置", padding="10")
        display_frame.pack(fill=tk.X, padx=5, pady=(0, 15))
        
        # 全屏显示迷你时钟设置
        mini_clock_frame = ttk.Frame(display_frame)
        mini_clock_frame.pack(fill=tk.X, pady=8, anchor="w")
        
        mini_clock_var = tk.BooleanVar(value=self.show_mini_on_fullscreen)
        mini_clock_checkbox = ttk.Checkbutton(mini_clock_frame, text="显示迷你时钟", variable=mini_clock_var)
        mini_clock_checkbox.pack(anchor="w", pady=4)
        
        # ===== 通知设置分组 =====
        notify_frame = ttk.LabelFrame(main_frame, text="通知设置", padding="10")
        notify_frame.pack(fill=tk.X, padx=5, pady=(0, 15))
        
        # 通知开关设置
        notification_var = tk.BooleanVar(value=self.notifications_enabled)
        notification_checkbox = ttk.Checkbutton(notify_frame, text="启用通知", variable=notification_var)
        notification_checkbox.pack(anchor="w", pady=4)
        
        # 添加通知管理按钮
        manage_notifications_frame = ttk.Frame(notify_frame)
        manage_notifications_frame.pack(fill=tk.X, pady=4, anchor="w")
        
        manage_notifications_button = ttk.Button(notify_frame, text="管理通知", 
                                               command=self.manage_notifications, width=10)
        manage_notifications_button.pack(side=tk.LEFT, padx=5)
        
        # 保存按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=20)
        
        def save_settings_action():
            try:
                exam_year = int(year_var.get())
                # 检查年份范围是否在2026-2048之间
                if exam_year < 2026 or exam_year > 2048:
                    self.show_windows_notification("再玩就玩坏了", "中考年份必须在2026-2048之间")
                else:
                    self.save_settings(exam_year, notification_var.get(), mini_clock_var.get())
                    self.show_windows_notification("搞定！", "设置项更改已保存")
                    settings_window.destroy()
            except ValueError:
                self.show_windows_notification("嗯？", "请输入有效的年份")
        save_button = ttk.Button(button_frame, text="保存", command=save_settings_action, width=10)
        save_button.pack(side=tk.RIGHT)
    
    def load_schedule(self):
        """加载课程表数据"""
        try:
            app_data_dir = self.get_appdata_path()
            json_path = os.path.join(app_data_dir, 'schedule.json')
            with open(json_path, 'r', encoding='utf-8') as f:
                self.schedule_data = json.load(f)
        except Exception as e:
            self.show_windows_notification("加载课程表失败", str(e))
            self.schedule_data = {'time_slots': [], 'school_days': {}}
    
    def display_todays_schedule(self):
        """显示当天的课程表（现代化美化版本）"""
        # 清空原有课程表内容
        for widget in self.schedule_content.winfo_children():
            widget.destroy()
        
        # 获取星期
        today_weekday = datetime.datetime.now().strftime('%A')
        
        # 保存所有课程框架引用，以便后续高亮当前课程
        self.class_frames = {}
        
        # 获取今天的课程
        if today_weekday in self.schedule_data['school_days']:
            today_classes = self.schedule_data['school_days'][today_weekday]
            
            # 创建时间槽映射
            time_slots_map = {slot['slot_id']: slot for slot in self.schedule_data['time_slots']}
            
            # 将课程分为上午和下午两部分
            morning_classes = []
            afternoon_classes = []
            
            for class_info in today_classes:
                slot_id = class_info['slot_id']
                if slot_id in time_slots_map:
                    time_slot = time_slots_map[slot_id]
                    start_time = time_slot['start_time']
                    start_hour = int(start_time.split(':')[0])
                    
                    if start_hour < 13:
                        morning_classes.append(class_info)
                    else:
                        afternoon_classes.append(class_info)
            
            # 检查上午是否有课程，如果没有则添加占位符
            if not morning_classes:
                # 添加上午无课占位符
                placeholder_frame = ttk.Frame(self.schedule_content, style="Black.TFrame", padding=0)
                placeholder_frame.pack(fill=tk.X, pady=0, padx=0)
                
                time_label = ttk.Label(placeholder_frame, text="00:00-12:00", style="ClassTime.Gray.TLabel")
                time_label.pack(side=tk.LEFT, padx=0, pady=0)
                
                name_label = ttk.Label(placeholder_frame, text="暂无", style="ClassName.Gray.TLabel")
                name_label.pack(side=tk.LEFT, pady=0, fill=tk.X, expand=True)
            else:
                # 显示上午课程
                for i, class_info in enumerate(morning_classes):
                    slot_id = class_info['slot_id']
                    class_name = class_info['name']
                    
                    # 获取对应的时间段
                    if slot_id in time_slots_map:
                        time_slot = time_slots_map[slot_id]
                        start_time = time_slot['start_time']
                        end_time = time_slot['end_time']
                        
                        style_name = "ClassTime.TLabel"
                        frame_style = "Black.TFrame"
                        
                        # 创建课程条目框架，使用ttk样式
                        class_frame = ttk.Frame(self.schedule_content, style=frame_style, padding=0)
                        class_frame.pack(fill=tk.X, pady=0, padx=0)
                        
                        # 保存课程框架引用，以便后续高亮
                        self.class_frames[slot_id] = class_frame
                        
                        # 创建课程时间标签，使用ttk样式
                        time_label = ttk.Label(class_frame, text=f"{start_time}-{end_time}", 
                                            style=style_name)
                        time_label.pack(side=tk.LEFT, padx=0, pady=0)
                        
                        # 创建课程名称标签，使用ttk样式
                        name_style = "ClassName.TLabel"
                        name_label = ttk.Label(class_frame, text=class_name, 
                                            style=name_style)
                        name_label.pack(side=tk.LEFT, pady=0, fill=tk.X, expand=True)
            
            # 如果上午有课程或下午有课程，添加午休分隔线
            if morning_classes or afternoon_classes:
                lunch_separator = ttk.Separator(self.schedule_content, orient='horizontal', 
                                            style="Separator.TSeparator")
                lunch_separator.pack(fill='x', pady=5)
            
            # 检查下午是否有课程，如果没有则添加占位符
            if not afternoon_classes:
                # 添加下午无课占位符
                placeholder_frame = ttk.Frame(self.schedule_content, style="Black.TFrame", padding=0)
                placeholder_frame.pack(fill=tk.X, pady=0, padx=0)
                
                time_label = ttk.Label(placeholder_frame, text="12:01-23:59", style="ClassTime.Gray.TLabel")
                time_label.pack(side=tk.LEFT, padx=0, pady=0)
                
                name_label = ttk.Label(placeholder_frame, text="暂无", style="ClassName.Gray.TLabel")
                name_label.pack(side=tk.LEFT, pady=0, fill=tk.X, expand=True)
            else:
                # 显示下午课程
                for i, class_info in enumerate(afternoon_classes):
                    slot_id = class_info['slot_id']
                    class_name = class_info['name']
                    
                    # 获取对应的时间段
                    if slot_id in time_slots_map:
                        time_slot = time_slots_map[slot_id]
                        start_time = time_slot['start_time']
                        end_time = time_slot['end_time']
                        
                        style_name = "ClassTime.TLabel"
                        frame_style = "Black.TFrame"
                        
                        # 创建课程条目框架，使用ttk样式
                        class_frame = ttk.Frame(self.schedule_content, style=frame_style, padding=0)
                        class_frame.pack(fill=tk.X, pady=0, padx=0)
                        
                        # 保存课程框架引用，以便后续高亮
                        self.class_frames[slot_id] = class_frame
                        
                        # 创建课程时间标签，使用ttk样式
                        time_label = ttk.Label(class_frame, text=f"{start_time}-{end_time}", 
                                            style=style_name)
                        time_label.pack(side=tk.LEFT, padx=0, pady=0)
                        
                        # 创建课程名称标签，使用ttk样式
                        name_style = "ClassName.TLabel"
                        name_label = ttk.Label(class_frame, text=class_name, 
                                            style=name_style)
                        name_label.pack(side=tk.LEFT, pady=0, fill=tk.X, expand=True)

        # 在课程表内容容器下方添加分隔线
        schedule_bottom_separator = ttk.Separator(self.schedule_content, orient='horizontal', 
                                                style="Separator.TSeparator")
        schedule_bottom_separator.pack(fill='x', pady=5)
        
        # 添加课程状态显示区域，使用ttk样式
        self.status_frame = ttk.Frame(self.schedule_content, style="Status.TLabel", padding=(10, 8))
        self.status_frame.pack(fill=tk.X, pady=5, padx=0)
        
        # 创建状态标签，使用ttk样式
        self.status_label = ttk.Label(self.status_frame, style="Status.TLabel", 
                                    justify='left', wraplength=400, text="⭕️ 内容加载中...")
        self.status_label.pack(fill=tk.X)

        # 创建底部按钮容器
        buttons_container = ttk.Frame(self.schedule_content, style="Black.TFrame")
        buttons_container.pack(fill=tk.X, pady=5, padx=0)

        # 补课设置按钮
        self.makeup_class_button = ttk.Button(buttons_container, text="补课设置", 
                                            command=self.show_makeup_class_window,
                                            style="Main.TButton")
        self.makeup_class_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=0)

        # 查看课表按钮
        self.view_schedule_button = ttk.Button(buttons_container, text="完整课表", 
                                            command=self.show_schedule_window,
                                            style="Main.TButton")
        self.view_schedule_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=0)

        # 全局设置按钮
        self.global_settings_button = ttk.Button(buttons_container, text="全局设置", 
                                            command=self.show_settings,
                                            style="Main.TButton")
        self.global_settings_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=0)

        # 右键菜单按钮
        self.right_click_menu_button = ttk.Button(buttons_container, text="更多功能", 
                                                command=lambda: self.show_context_menu(None),
                                                style="Main.TButton")
        self.right_click_menu_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=0)

        # 绑定鼠标事件到状态显示区域
        self.status_frame.bind("<Button-1>", self.on_mouse_down)
        self.status_frame.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.status_frame.bind("<B1-Motion>", self.on_mouse_drag)
        # self.status_label.bind("<Button-1>", self.on_mouse_down)
        # self.status_label.bind("<ButtonRelease-1>", self.on_mouse_up)
        # self.status_label.bind("<B1-Motion>", self.on_mouse_drag)

    def show_makeup_class_window(self):
        """显示补课设置窗口"""
        # 保存原始课程表数据，用于取消时恢复
        self.original_schedule_data = self.schedule_data.copy()
        
        # 创建补课设置窗口
        self.makeup_window = tk.Toplevel(self.root)
        self.makeup_window.title("补课设置")
        self.makeup_window.geometry("350x250")
        self.makeup_window.resizable(False, False)
        
        # 设置窗口居中显示
        self.makeup_window.update_idletasks()
        width = self.makeup_window.winfo_width()
        height = self.makeup_window.winfo_height()
        x = (self.makeup_window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.makeup_window.winfo_screenheight() // 2) - (height // 2)
        self.makeup_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        # 创建主框架
        main_frame = ttk.Frame(self.makeup_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建标题（使用ttk.Label）
        title_label = ttk.Label(main_frame, text="补课设置", font=('Microsoft YaHei UI', 14, 'bold'))
        title_label.pack(pady=(10, 15))
        
        # 创建选择框架（使用ttk.Frame）
        select_frame = ttk.Frame(main_frame)
        select_frame.pack(pady=10)
        
        # 添加文本和下拉列表
        ttk.Label(select_frame, text="今天改为上", font=('Microsoft YaHei UI', 12)).pack(side=tk.LEFT, padx=5)
        
        # 星期选项（周一至周五，周日）
        weekday_options = ['周一', '周二', '周三', '周四', '周五', '周日']
        self.weekday_var = tk.StringVar(value=weekday_options[0])
        
        # 创建ttk风格的下拉菜单
        weekday_menu = ttk.Combobox(select_frame, textvariable=self.weekday_var, values=weekday_options, 
                                  font=('Microsoft YaHei UI', 12), width=5, state="readonly")
        weekday_menu.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(select_frame, text="的课程", font=('Microsoft YaHei UI', 12)).pack(side=tk.LEFT, padx=5)
        
        # 添加提示文字（使用ttk.Label，设置样式）
        hint_label = ttk.Label(main_frame, text="更改单次有效，下次启动将自动还原", 
                            font=('Microsoft YaHei UI', 9), foreground="#666666")
        hint_label.pack(pady=10)
        
        # 创建按钮框架（使用ttk.Frame）
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=10)
        
        # 添加ttk风格的取消按钮
        cancel_button = ttk.Button(button_frame, text="还原", 
                                 command=self.cancel_makeup_class)
        cancel_button.pack(side=tk.LEFT, padx=10)

        # 添加ttk风格的确定按钮
        confirm_button = ttk.Button(button_frame, text="确定", 
                                  style="Accent.TButton",
                                  command=self.apply_makeup_class)
        confirm_button.pack(side=tk.LEFT, padx=10)
        
    def apply_makeup_class(self):
        """应用补课设置，刷新当天课表"""
        # 获取用户选择的星期
        selected_weekday_zh = self.weekday_var.get()
        
        # 星期转换映射
        weekday_map = {
            '周一': 'Monday',
            '周二': 'Tuesday', 
            '周三': 'Wednesday',
            '周四': 'Thursday',
            '周五': 'Friday',
            '周日': 'Sunday'
        }
        
        # 获取对应英文星期
        selected_weekday = weekday_map.get(selected_weekday_zh, None)
        
        # 获取今天的星期
        today_weekday = datetime.datetime.now().strftime('%A')
        
        # 如果选择了有效的星期，则更新今天的课程为所选星期的课程
        if selected_weekday and selected_weekday in self.schedule_data['school_days']:
            # 保存原始课程
            if hasattr(self, 'original_today_classes'):
                # 如果已经有过修改，则不再覆盖原始数据
                pass
            else:
                # 第一次修改时保存原始数据
                if today_weekday in self.schedule_data['school_days']:
                    self.original_today_classes = self.schedule_data['school_days'][today_weekday].copy()
                else:
                    self.original_today_classes = []
            
            # 更新今天的课程表
            self.schedule_data['school_days'][today_weekday] = self.schedule_data['school_days'][selected_weekday].copy()
            
            # 刷新显示
            self.display_todays_schedule()
            self.update_class_status()
            
            
            # 显示通知
            self.show_windows_notification("设置成功", f"今天已改为上{selected_weekday_zh}的课程")
        
        # 关闭窗口
        self.makeup_window.destroy()

    def cancel_makeup_class(self):
        """取消补课设置"""
        # 恢复原始课程
        if hasattr(self, 'original_today_classes'):
            today_weekday = datetime.datetime.now().strftime('%A')
            self.schedule_data['school_days'][today_weekday] = self.original_today_classes.copy()
            
            # 刷新显示
            self.display_todays_schedule()
            self.update_class_status()
        
        # 关闭窗口
        self.makeup_window.destroy()

    def update_class_status(self):
        """更新课程状态显示"""
        # 获取当前时间
        current_time = datetime.datetime.now()
        # # 调试用：固定时间点
        # current_time = current_time.replace(hour=21, minute=15)
        today_weekday = current_time.strftime('%A')
        
        # 初始化状态变量(全局变量)
        global current_status
        current_status = "课间"

        global next_class
        next_class = "无"

        current_slot_id = None
        
        # 重置所有课程条目的样式 - 使用ttk样式机制
        if hasattr(self, 'class_frames'):
            for i, (slot_id, frame) in enumerate(self.class_frames.items()):
                # 恢复默认的深灰色背景
                frame.configure(style="Black.TFrame")
                
                # 更新框架内所有标签的样式
                for j, child in enumerate(frame.winfo_children()):
                    if isinstance(child, ttk.Label):
                        if j == 0:  # 时间标签
                            child.configure(style="ClassTime.TLabel")
                        else:  # 课程名称标签
                            child.configure(style="ClassName.TLabel")
    
        # 检查是否有今天的课程
        if today_weekday in self.schedule_data['school_days']:
            today_classes = self.schedule_data['school_days'][today_weekday]
            time_slots_map = {slot['slot_id']: slot for slot in self.schedule_data['time_slots']}
            
            # 将当前时间转换为分钟
            current_minutes = current_time.hour * 60 + current_time.minute

            # 统计已结束课程数量
            completed_classes = 0
            total_classes = len(today_classes)
            
            # 遍历所有课程，查找当前 & 下一节课
            found_current_class = False
            for i, class_info in enumerate(today_classes):
                slot_id = class_info['slot_id']
                if slot_id in time_slots_map:
                    time_slot = time_slots_map[slot_id]
                    start_time = time_slot['start_time']
                    end_time = time_slot['end_time']
                    
                    # 解析开始 & 结束时间
                    start_hour, start_minute = map(int, start_time.split(':'))
                    end_hour, end_minute = map(int, end_time.split(':'))
                    
                    start_minutes = start_hour * 60 + start_minute
                    # 提前2分钟进行判断（预备铃
                    start_minutes_with_preparation = start_minutes - 2
                    end_minutes = end_hour * 60 + end_minute

                    # 已结束课程标为灰色
                    if end_minutes < current_minutes and hasattr(self, 'class_frames') and slot_id in self.class_frames:
                        completed_classes += 1  # 增加已结束课程计数
                        class_frame = self.class_frames[slot_id]
                        for j, child in enumerate(class_frame.winfo_children()):
                            if isinstance(child, ttk.Label):
                                if j == 0:  # 时间标签
                                    child.configure(style="ClassTime.Gray.TLabel")
                                else:  # 课程名称标签
                                    child.configure(style="ClassName.Gray.TLabel")
                
                    # 检查当前是否在这节课的时间范围内
                    if start_minutes_with_preparation <= current_minutes <= end_minutes:
                        current_status = class_info['name']
                        found_current_class = True
                        current_slot_id = slot_id
                        
                        # 查找下一节课
                        if i < len(today_classes) - 1:
                            next_class_info = today_classes[i + 1]
                            next_class = next_class_info['name']

                        # 高亮当前课程
                        if hasattr(self, 'class_frames') and slot_id in self.class_frames:
                            class_frame = self.class_frames[slot_id]
                            for j, child in enumerate(class_frame.winfo_children()):
                                if isinstance(child, ttk.Label):
                                    if j == 0:  # 时间标签
                                        child.configure(style="CurrentClassTime.TLabel")
                                    else:  # 课程名称标签
                                        child.configure(style="CurrentClassName.TLabel")
                        break
                    
                    # 检查当前是否在下一节课之前
                    elif current_minutes < start_minutes_with_preparation and not found_current_class:
                        current_status = "课间"
                        next_class = class_info['name']
                        break
            
            # 如果当前时间在所有课程之后
            if not found_current_class and today_classes:
                last_class_slot = today_classes[-1]['slot_id']
                if last_class_slot in time_slots_map:
                    last_class_end = time_slots_map[last_class_slot]['end_time']
                    last_end_hour, last_end_minute = map(int, last_class_end.split(':'))
                    last_end_minutes = last_end_hour * 60 + last_end_minute
                    
                    # 判断是否在午休时间
                    if 12 * 60 <= current_minutes < 14 * 60:
                        current_status = "午休"
                        # 查找下午的第一节课
                        for class_info in today_classes:
                            slot_id = class_info['slot_id']
                            if slot_id in time_slots_map:
                                start_hour, _ = map(int, time_slots_map[slot_id]['start_time'].split(':'))
                                if start_hour >= 14:
                                    next_class = class_info['name']
                                    break
                    # 判断是否已经放学
                    elif current_minutes >= last_end_minutes:
                        current_status = "放学"
                        next_class = "明日课程"

            # 更新课程标题，添加进度信息
            self.schedule_title.configure(text=f"今日课程（{completed_classes}/{total_classes}）")
            self.current_next_class = next_class
    
    def update_status_text(self):
        """更新状态标签文本，自动刷新"""
        if self.status_label is not None:
            # 根据时间添加问候语
            now = datetime.datetime.now()
            hour = now.hour
            if 6 <= hour < 12:
                greeting = "🌅 早上好"
            elif 12 <= hour < 14:
                greeting = "🍱 中午好"
            elif 14 <= hour < 18:
                greeting = "💪 下午好"
            else:
                greeting = "🌙 晚上好"
            
            # 1. 获取天气信息
            weather_info = ""
            try:
                weather_response = requests.get("https://api.seniverse.com/v3/weather/now.json?key=SIHZWG1tgvaojxn_N&location=ip&language=zh-Hans&unit=c", timeout=3)
                if weather_response.status_code == 200:
                    weather_data = weather_response.json()
                    if "results" in weather_data and weather_data["results"]:
                        now = weather_data["results"][0]["now"]
                        weather_text = now["text"]
                        temperature = now["temperature"]
                        weather_info = f"当前{weather_text}，{temperature}°C"
            except Exception as e:
                # 获取失败
                weather_info = "今天天气不错哦"
            
            # 2. 检查特殊日期
            special_dates = [
                {"name": "期末考试", "month": 1, "start_day": 22, "end_day": 23}
                # 这里添加更多特殊日期
            ]
            special_date_info = ""
            today = datetime.datetime.now().date()
            next_special_date = None
            next_special_name = ""
            
            # 检查今天是否是特殊日
            for date_info in special_dates:
                start_date = datetime.date(today.year, date_info["month"], date_info["start_day"])
                end_date = datetime.date(today.year, date_info["month"], date_info["end_day"])
                if start_date <= today <= end_date:
                    special_date_info = f"📅 今天{date_info['name']}"
                    break
                # 查找下一个特殊日
                if not next_special_date or start_date > today and start_date < next_special_date:
                    next_special_date = start_date
                    next_special_name = date_info['name']
            
            # 如果当天没有特殊日，显示下一个特殊日的倒计时
            if not special_date_info and next_special_date:
                days_left = (next_special_date - today).days
                special_date_info = f"📅 距离「{next_special_name}」还有 {days_left} 天"
            
            # 3. 计算学习统计
            study_stats = "⏱ 暂无学习任务"
            try:
                today_weekday = datetime.datetime.now().strftime('%A')
                if today_weekday in self.schedule_data['school_days']:
                    today_classes = self.schedule_data['school_days'][today_weekday]
                    time_slots_map = {slot['slot_id']: slot for slot in self.schedule_data['time_slots']}
                    current_minutes = datetime.datetime.now().hour * 60 + datetime.datetime.now().minute
                    
                    # 计算总学习时间（分钟）
                    total_study_minutes = 0
                    current_class_progress = 0
                    class_end_minutes = 0
                    is_in_class = False
                    
                    for class_info in today_classes:
                        slot_id = class_info['slot_id']
                        if slot_id in time_slots_map:
                            time_slot = time_slots_map[slot_id]
                            start_hour, start_minute = map(int, time_slot['start_time'].split(':'))
                            end_hour, end_minute = map(int, time_slot['end_time'].split(':'))
                            
                            start_minutes = start_hour * 60 + start_minute
                            end_minutes = end_hour * 60 + end_minute
                            class_duration = end_minutes - start_minutes
                            
                            # 已结束的课程
                            if current_minutes >= end_minutes:
                                total_study_minutes += class_duration
                            # 正在上的课
                            elif current_minutes >= start_minutes:
                                total_study_minutes += (current_minutes - start_minutes)
                                current_class_progress = current_minutes - start_minutes
                                class_end_minutes = end_minutes
                                is_in_class = True
                                break
                    
                    # 转换为小时和分钟
                    study_hours = total_study_minutes // 60
                    study_mins = total_study_minutes % 60
                    study_stats = f"⏱ 今日已学习 {study_hours} 小时 {study_mins} 分钟"
                    
                    # 即将下课提醒
                    if is_in_class:
                        remaining_mins = class_end_minutes - current_minutes
                        if 1 <= remaining_mins <= 15:
                            study_stats += f"\n🔔 「{current_status}」还有 {remaining_mins} 分钟就要下课啦"
            except Exception as e:
                study_stats = "⏱ 学习统计中..."
            
            # 4. 预设鸡汤
            motivational_quotes = [
                "恭喜你抽到彩蛋了（概率1%）",
                "今天也要元气满满哦！",
                "学习使我快乐😊",
                "加油，你是最棒的！",
                "坚持就是胜利✊",
                "每一份努力都不会白费",
                "越努力，越幸运✨",
                "相信自己，你可以的！",
                "今天的努力，明天的收获",
                "保持专注，成就辉煌",
                "学习是最美的遇见",
                "心有多大，舞台就有多大",
                "成功属于坚持不懈的人",
                "每天进步一点点，就是最大的成功",
                "只要努力，就没有过不去的坎",
                "梦想需要行动，而不是空想",
                "时间是最公平的，你付出多少，就会得到多少",
                "困难是暂时的，胜利是必然的",
                "学习没有捷径，只有脚踏实地",
                "态度决定一切，细节决定成败",
                "不要害怕失败，失败是成功之母",
                "现在的努力，是为了未来的自由",
                "坚持就是胜利，努力总会有回报",
                "每一次失败，都是成功的垫脚石",
                "相信自己，你比想象中更强大",
                "学习是一场马拉松，不是短跑",
                "只要不放弃，就永远有希望",
                "今天的汗水，明天的欢笑",
                "知识改变命运，学习成就未来",
                "机会总是留给有准备的人",
                "没有做不到的事，只有不想做的人",
                "努力吧，未来的你会感谢现在的自己",
                "学习是投资，不是消费",
                "成功需要耐心，需要坚持",
                "每一次努力，都是在接近梦想",
                "不要等待机会，要创造机会",
                "学习使你成长，成长让你快乐",
                "困难像弹簧，你强它就弱",
                "相信努力，相信未来",
                "今天的付出，明天的收获",
                "坚持到底，就是胜利",
                "学习是为了成为更好的自己",
                "只要有梦想，就有动力",
                "努力的人，运气都不会太差",
                "成功没有秘诀，只有坚持和努力",
                "每一步都算数，每一份努力都值得",
                "不要怕慢，就怕站",
                "学习是终身的事业",
                "现在的辛苦，是为了将来的幸福",
                "相信自己，你一定能行",
                "坚持就是胜利，加油！",
                "努力吧，少年！未来属于你",
                "学习是一件快乐的事",
                "只要努力，就会有奇迹",
                "成功的路上，没有捷径",
                "每一份努力，都会有回报",
                "坚持到底，永不放弃",
                "学习改变人生，知识成就梦想",
                "相信自己，你是最棒的",
                "今天的努力，明天的成功",
                "努力吧，未来可期",
                "学习是进步的阶梯",
                "只要有信心，就能成功",
                "每一次坚持，都是成长",
                "努力的人，最美",
                "学习是为了更好的生活",
                "坚持就是胜利，成功就在前方",
                "相信自己，你一定可以",
                "今天的努力，明天的辉煌",
                "努力吧，少年！",
                "学习是一种享受",
                "只要不放弃，就会成功",
                "每一份努力，都不会被辜负",
                "坚持到底，就是胜利",
                "学习是为了实现梦想",
                "相信自己，你能行",
                "今天的汗水，明天的成功",
                "努力吧，未来属于努力的人",
                "学习是一种快乐",
                "只要努力，就会有收获",
                "每一次努力，都是在成长",
                "坚持就是胜利，加油吧！",
                "相信自己，你一定能成功",
                "今天的努力，明天的果实",
                "努力吧，少年！未来是你的",
                "学习是一种成长",
                "只要有梦想，就有希望",
                "每一份努力，都值得尊重",
                "坚持到底，永不言弃",
                "学习是为了更好的自己",
                "相信自己，你是最棒的！",
                "今天的努力，明天的成就",
                "努力吧，未来在等你",
                "学习是一种幸福",
                "只要努力，就会有回报",
                "每一次坚持，都是胜利",
                "坚持就是胜利，成功属于你",
                "相信自己，你可以的！",
                "今天的努力，明天的快乐",
                "努力吧，少年！加油！",
                "学习是一种财富",
                "只要不放弃，就有希望",
                "每一份努力，都在接近成功",
                "坚持到底，就是成功",
                "相信自己，你一定能做到",
                "今天的努力，明天的美好",
                "努力吧，未来属于你！"
            ]
            motivational_quote = random.choice(motivational_quotes)
            
            # 组合最终状态文本
            status_text = f"""{greeting}！{weather_info}
{special_date_info}
{study_stats}
>> {motivational_quote}"""
            self.status_label.configure(style="Status.TLabel", text=status_text)
        
        # 自动刷新
        self.root.after(10 * 1000, self.update_status_text)  # 10秒刷新一次

    def bind_mouse_events(self):
        """绑定鼠标事件"""
        # 倒计时标签
        self.countdown_label.bind("<Button-1>", self.on_mouse_down)
        self.countdown_label.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.countdown_label.bind("<B1-Motion>", self.on_mouse_drag)
        
        # 名言标签
        self.quote_label.bind("<Button-1>", self.on_mouse_down)
        self.quote_label.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.quote_label.bind("<B1-Motion>", self.on_mouse_drag)
        
        # 来源标签
        self.from_label.bind("<Button-1>", self.on_mouse_down)
        self.from_label.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.from_label.bind("<B1-Motion>", self.on_mouse_drag)
        
        # 主框架
        self.main_frame.bind("<Button-1>", self.on_mouse_down)
        self.main_frame.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.main_frame.bind("<B1-Motion>", self.on_mouse_drag)
        
        # 左侧框架
        self.left_frame.bind("<Button-1>", self.on_mouse_down)
        self.left_frame.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.left_frame.bind("<B1-Motion>", self.on_mouse_drag)
        
        # 右侧框架
        self.right_frame.bind("<Button-1>", self.on_mouse_down)
        self.right_frame.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.right_frame.bind("<B1-Motion>", self.on_mouse_drag)
        
        # 右键菜单
        self.countdown_label.bind("<Button-3>", self.show_context_menu)
        self.quote_label.bind("<Button-3>", self.show_context_menu)
        self.from_label.bind("<Button-3>", self.show_context_menu)
        self.main_frame.bind("<Button-3>", self.show_context_menu)
        self.left_frame.bind("<Button-3>", self.show_context_menu)
        self.right_frame.bind("<Button-3>", self.show_context_menu)
        
    def on_mouse_down(self, event):
        """鼠标按下事件"""
        self.drag_start_time = datetime.datetime.now()
        self.is_dragging = False
        self.drag_start_x = event.x
        self.drag_start_y = event.y
        
    def on_mouse_up(self, event):
        """鼠标释放事件"""
        if self.drag_start_time is not None:
            # 计算按下和释放之间的时间差
            press_duration = (datetime.datetime.now() - self.drag_start_time).total_seconds() * 1000
            
            # 如果不是拖动状态且按下的时间小于阈值，则认为是点击
            if not self.is_dragging and press_duration < self.click_threshold:
                self.refresh_content(event)
                
            self.drag_start_time = None
            self.is_dragging = False
        
    def on_mouse_drag(self, event):
        """鼠标拖动事件"""
        if self.drag_start_time is not None:
            # 计算鼠标移动距离
            dx = event.x - self.drag_start_x
            dy = event.y - self.drag_start_y
            distance = (dx**2 + dy**2)**0.5
            
            # 如果移动距离超过阈值，则认为是拖动
            if distance > self.drag_threshold:
                self.is_dragging = True
                self.move_window(event)
                
    def move_window(self, event):
        """移动窗口"""
        # 计算窗口新位置
        deltax = event.x - self.drag_start_x
        deltay = event.y - self.drag_start_y
        x = self.root.winfo_x() + deltax
        y = self.root.winfo_y() + deltay
        self.root.geometry(f"+{x}+{y}")
        
    def set_window_position(self):
        """设置窗口到屏幕正中央"""
        # 等待窗口内容渲染完成
        self.root.update_idletasks()
        
        # 获取屏幕尺寸
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # 获取窗口尺寸
        window_width = self.root.winfo_width()
        window_height = self.root.winfo_height()
        
        # 计算屏幕正中央位置
        x = (screen_width - window_width) // 2  # 整除确保是整数像素
        y = (screen_height - window_height) // 2
        
        # 设置窗口位置
        self.root.geometry(f"+{x}+{y}")
        
    def show_windows_notification(self, title, message):
        """显示Windows原生通知"""
        # 检查是否启用了通知
        if not self.notifications_enabled:
            return
        
        def show_notification_thread():
            try:
                toaster = ToastNotifier()

                # 确定图标文件的路径，区分开发环境和生产环境
                if hasattr(sys, '_MEIPASS'):
                    # 打包后的环境（生产环境）
                    icon_path = os.path.join(sys._MEIPASS, 'logo.ico')
                else:
                    # 开发环境
                    icon_path = './logo.ico'
                        
                toaster.show_toast(
                    title=title,
                    msg=message,
                    icon_path=icon_path,  # 图标路径
                    duration=10,  # 显示10秒
                    threaded=True  # 非阻塞模式
                )
            except Exception as e:
                return
    
        # 创建并启动新线程来显示通知
        notification_thread = threading.Thread(target=show_notification_thread)
        notification_thread.daemon = True  # 设置为守护线程，主线程结束时自动结束
        notification_thread.start()
    
    def update_countdown(self):
        """更新倒计时显示"""
        # 获取当前时间并格式化显示
        current_time = datetime.datetime.now().strftime("%H:%M:%S")
        self.time_label.config(text=current_time)

        if self.show_mini_on_fullscreen:
            # 显示小窗口并更新时间
            if not self.mini_window.winfo_ismapped():
                self.mini_window.deiconify()
                # 只有在位置未设置或未被用户移动过的情况下才设置位置
                if not self.mini_window_position_set:
                    # 放置在上方居中位置
                    screen_width = self.root.winfo_screenwidth()
                    mini_width = 360  # 小窗口宽度
                    mini_height = 40  # 小窗口高度
                    x = (screen_width - mini_width) // 5
                    y = 1  # 距离顶部像素
                    self.mini_window.geometry(f"{mini_width}x{mini_height}+{x}+{y}")
                    self.mini_window_position_set = True

            # 更新小窗口显示，添加下节课信息
            if hasattr(self, 'current_next_class'):
                if self.current_next_class == "明日课程":
                    mini_text = f"{current_time} · 放学啦！"
                elif self.current_next_class == "无":
                    # 计算距离放学时间
                    today_weekday = datetime.datetime.now().strftime('%A')
                    current_time_standard = datetime.datetime.now()
                    # 调试用：指定时间点
                    current_time_standard = current_time_standard.replace(hour=21, minute=29)
                    if today_weekday in self.schedule_data['school_days']:
                        today_classes = self.schedule_data['school_days'][today_weekday]
                        time_slots_map = {slot['slot_id']: slot for slot in self.schedule_data['time_slots']}
                        if today_classes:
                            last_class_slot = today_classes[-1]['slot_id']
                            if last_class_slot in time_slots_map:
                                last_class_end = time_slots_map[last_class_slot]['end_time']
                                last_end_hour, last_end_minute = map(int, last_class_end.split(':'))
                                last_end_time = current_time_standard.replace(hour=last_end_hour, minute=last_end_minute, second=0, microsecond=0)
                                time_until_end = math.ceil((last_end_time - current_time_standard).total_seconds() / 60)
                    mini_text = f"{current_time} · {time_until_end}分钟后放学"
                else:
                    mini_text = f"{current_time} · 下节「{self.current_next_class}」"
            else:
                mini_text = current_time

            self.mini_time_label.config(text=mini_text)
        else:
            # 如果主窗口可见或设置了不显示迷你时钟，则隐藏小窗口
            if self.mini_window.winfo_ismapped():
                self.mini_window.withdraw()

        # 获取当前日期并格式化显示为"XXXX年X月X日"格式
        weekdays = {'Monday': '星期一', 'Tuesday': '星期二', 'Wednesday': '星期三', 'Thursday': '星期四', 'Friday': '星期五', 'Saturday': '星期六', 'Sunday': '星期日'}
        current_date = datetime.datetime.now().strftime(f"%Y年%m月%d日")
        weekday = weekdays[datetime.datetime.now().strftime('%A')]
        self.date_label.config(text=f"{current_date} {weekday}")
        
        # 检查是否需要发送通知
        self.check_and_send_notifications()

        today = datetime.date.today()
        days_left = (self.exam_date - today).days
        
        # 根据剩余天数改变倒计时颜色
        if days_left > 0:
            text = f"距离中考还有「{days_left}」天"
            # 根据剩余天数改变颜色：充足(>180天)、中等(90-180天)、紧急(<90天)
            if days_left > 180:
                self.countdown_label.configure(text=text, style="Countdown.Blue.TLabel")
            elif days_left > 90:
                self.countdown_label.configure(text=text, style="Countdown.Orange.TLabel")
            else:
                self.countdown_label.configure(text=text, style="Countdown.Red.TLabel")
        elif days_left == 0:
            text = "今天中考！加油！"
            self.countdown_label.configure(text=text, style="Countdown.Gold.TLabel")
        else:
            text = "中考已结束！有缘再会！"
            self.countdown_label.configure(text=text, style="Countdown.Purple.TLabel")
            
        # 只在星期改变时才刷新整个课程表
        current_weekday = datetime.datetime.now().strftime('%A')
        if not hasattr(self, 'last_weekday') or self.last_weekday != current_weekday:
            self.display_todays_schedule()
            self.last_weekday = current_weekday
        else:
            # 只更新课程状态，不重新创建整个课程表
            self.update_class_status()
        
        # 每秒更新
        self.root.after(1000, self.update_countdown)

    def check_and_send_notifications(self):
        """检查并发送所有需要发送的通知"""
        if not self.notifications_enabled:
            return
        
        now = datetime.datetime.now()
        current_date_str = now.strftime("%Y-%m-%d")
        current_hour = now.hour
        current_minute = now.minute
        
        # 遍历所有通知
        for notification in self.notifications:
            # 检查通知是否启用
            if not notification['enabled']:
                continue
                
            # 检查是否到了通知时间（添加1分钟的容差，确保不会错过）
            time_match = (current_hour == notification['hour'] and current_minute == notification['minute'])
            
            if time_match:
                # 检查今天是否已经发送过该通知
                notification_id = notification['id']
                if notification_id not in self.last_notification_dates or self.last_notification_dates[notification_id] != current_date_str:
                    # 发送通知
                    print(f"发送通知: {notification['title']} - {notification['message']}")
                    self.show_windows_notification(notification['title'], notification['message'])
                    # 更新最后发送日期
                    self.last_notification_dates[notification_id] = current_date_str
                    
    def manage_notifications(self):
        """显示通知管理窗口"""
        # 重新加载通知列表
        self.load_notifications()
        
        # 创建通知管理窗口
        notification_window = tk.Toplevel(self.root)
        notification_window.title("通知管理")
        notification_window.geometry("500x400")
        notification_window.configure(bg="#f0f0f0")
        
        # 设置窗口居中显示
        notification_window.update_idletasks()
        width = notification_window.winfo_width()
        height = notification_window.winfo_height()
        x = (notification_window.winfo_screenwidth() // 2) - (width // 2)
        y = (notification_window.winfo_screenheight() // 2) - (height // 2)
        notification_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
        
        # 创建主框架
        main_frame = ttk.Frame(notification_window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建列表框架
        list_frame = ttk.LabelFrame(main_frame, text="通知列表", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 创建滚动区域
        canvas = tk.Canvas(list_frame)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 存储通知项的引用
        notification_items = {}
        
        # 创建表头
        headers = ["状态", "标题", "时间", "操作"]
        # 使用统一的列权重配置
        column_weights = [1, 3, 1, 2]  # 列权重，使标题列占据更多空间
        
        # 创建ttk样式，使界面更统一
        style = ttk.Style()
        style.configure("Notification.TLabel", font=('Microsoft YaHei UI', 10))
        style.configure("Header.TLabel", font=('Microsoft YaHei UI', 12, 'bold'), padding=5)

        # 创建统一的表格容器框架，用于包含表头和所有表体项
        table_container = ttk.Frame(scrollable_frame)
        table_container.pack(fill=tk.X)
        
        # 设置表格容器的列权重
        for i, weight in enumerate(column_weights):
            table_container.grid_columnconfigure(i, weight=weight, uniform="table")
        
        # 创建表头行
        header_row = ttk.Frame(table_container)
        header_row.grid(row=0, column=0, sticky="nsew", columnspan=len(column_weights))
        
        # 设置表头行的列权重
        for i, weight in enumerate(column_weights):
            header_row.grid_columnconfigure(i, weight=weight, uniform="table")
        
        # 创建表头标签
        for i, header in enumerate(headers):
            header_label = ttk.Label(header_row, text=header, font=('Microsoft YaHei UI', 12, 'bold'), anchor="center")
            header_label.grid(row=0, column=i, sticky="nsew", padx=1, pady=1)
        
        # 添加分割线
        separator = ttk.Separator(table_container, orient="horizontal")
        separator.grid(row=1, column=0, sticky="ew", columnspan=len(column_weights))

        # 填充通知列表
        def populate_notification_list():
            # 清除现有列表（除了表头和分割线）
            for widget in table_container.winfo_children():
                if widget not in [header_row, separator]:
                    widget.destroy()
            notification_items.clear()
            
            # 重新加载通知
            self.load_notifications()
            
            # 添加到列表
            for idx, notification in enumerate(self.notifications):
                # 创建通知项框架，直接放在表格容器中
                item_frame = ttk.Frame(table_container)
                item_frame.grid(row=idx+2, column=0, sticky="nsew", columnspan=len(column_weights), pady=2)
                
                # 设置通知项框架的列权重
                for i, weight in enumerate(column_weights):
                    item_frame.grid_columnconfigure(i, weight=weight, uniform="table")
                
                # 状态
                enabled_var = tk.BooleanVar(value=notification['enabled'])
                enabled_checkbox = ttk.Checkbutton(item_frame, variable=enabled_var, command=lambda n=notification: toggle_notification(n))
                enabled_checkbox.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
                
                # 标题
                title_label = ttk.Label(item_frame, text=notification['title'], font=('Microsoft YaHei UI', 10), anchor="w")
                title_label.grid(row=0, column=1, sticky="nsew", padx=1, pady=1)
                
                # 时间
                time_str = f"{notification['hour']:02d}:{notification['minute']:02d}"
                time_label = ttk.Label(item_frame, text=time_str, font=('Microsoft YaHei UI', 10), anchor="center")
                time_label.grid(row=0, column=2, sticky="nsew", padx=1, pady=1)
                
                # 操作按钮
                actions_frame = ttk.Frame(item_frame)
                actions_frame.grid(row=0, column=3, sticky="nsew", padx=1, pady=1)
                
                # 设置操作框架内部布局
                actions_frame.grid_columnconfigure(0, weight=1)
                actions_frame.grid_columnconfigure(1, weight=1)

                edit_button = ttk.Button(actions_frame, text="编辑", width=6, command=lambda n=notification: edit_notification(n['id']))
                edit_button.pack(side=tk.LEFT, padx=2)
                
                delete_button = ttk.Button(actions_frame, text="删除", width=6, command=lambda n=notification: delete_notification(n))
                delete_button.pack(side=tk.LEFT, padx=2)
                
                # 存储引用
                notification_items[notification['id']] = {
                    'frame': item_frame,
                    'enabled_var': enabled_var,
                    'title_label': title_label,
                    'time_label': time_label
                }
        
        # 切换通知启用状态
        def toggle_notification(notification):
            # 更新数据库
            self.cursor.execute("UPDATE notifications SET enabled = ? WHERE id = ?", 
                              (1 if not notification['enabled'] else 0, notification['id']))
            self.conn.commit()
            
            # 更新内存中的数据
            for n in self.notifications:
                if n['id'] == notification['id']:
                    n['enabled'] = not n['enabled']
                    break
        
        # 删除通知
        def delete_notification(notification):
            # 显示确认对话框
            if messagebox.askyesno("确认删除", f"确定要删除通知'{notification['title']}'吗？"):
                # 从数据库删除
                self.cursor.execute("DELETE FROM notifications WHERE id = ?", (notification['id'],))
                self.conn.commit()
                
                # 刷新列表
                populate_notification_list()
        
        # 加载通知列表
        populate_notification_list()
        
        # 添加新通知
        def add_notification():
            # 创建添加窗口
            add_window = tk.Toplevel(notification_window)
            add_window.title("添加通知")
            add_window.geometry("400x320")  # 增大窗口尺寸，提供更舒适的操作空间
            add_window.resizable(False, False)
            
            # 创建框架
            frame = ttk.Frame(add_window, padding="20")
            frame.pack(fill=tk.BOTH, expand=True)
            
            # 配置网格布局的列权重，让第二列可以适当扩展
            frame.columnconfigure(0, weight=0)
            frame.columnconfigure(1, weight=1)
            
            # 标题输入
            ttk.Label(frame, text="标题:", font=('Microsoft YaHei UI', 10)).grid(
                row=0, column=0, sticky=tk.W + tk.N, pady=(0, 10), padx=(0, 10))
            title_var = tk.StringVar()
            title_entry = ttk.Entry(frame, textvariable=title_var, width=35)
            title_entry.grid(row=0, column=1, sticky=tk.E + tk.W, pady=(0, 10))
            title_entry.focus_set()  # 自动聚焦到标题输入框
            
            # 消息输入
            ttk.Label(frame, text="消息:", font=('Microsoft YaHei UI', 10)).grid(
                row=1, column=0, sticky=tk.W + tk.N, pady=(0, 10), padx=(0, 10))
            text_widget = tk.Text(frame, height=6, width=30, font=('Microsoft YaHei UI', 10))
            text_widget.grid(row=1, column=1, sticky=tk.E + tk.W + tk.N + tk.S, pady=(0, 10))
            
            # 添加消息框的垂直滚动条
            msg_scrollbar = ttk.Scrollbar(frame, orient="vertical", command=text_widget.yview)
            msg_scrollbar.grid(row=1, column=2, sticky=tk.N + tk.S, pady=(0, 10))
            text_widget.configure(yscrollcommand=msg_scrollbar.set)
            
            # 时间选择
            ttk.Label(frame, text="时间:", font=('Microsoft YaHei UI', 10)).grid(
                row=2, column=0, sticky=tk.W + tk.N, pady=(0, 15), padx=(0, 10))
            time_frame = ttk.Frame(frame)
            time_frame.grid(row=2, column=1, sticky=tk.W, pady=(0, 15))
            
            # 获取当前时间
            now = datetime.datetime.now()
            
            # 小时步进器
            hour_var = tk.IntVar(value=now.hour)
            hour_spinbox = ttk.Spinbox(time_frame, from_=0, to=23, textvariable=hour_var, 
                                    width=4, font=('Microsoft YaHei UI', 10), justify=tk.CENTER)
            hour_spinbox.pack(side=tk.LEFT, padx=2)
            
            # 分隔符
            ttk.Label(time_frame, text=":", font=('Microsoft YaHei UI', 12, 'bold')).pack(side=tk.LEFT)
            
            # 分钟步进器
            minute_var = tk.IntVar(value=now.minute)  # 按5分钟步进
            minute_spinbox = ttk.Spinbox(time_frame, from_=0, to=59, textvariable=minute_var,
                                        width=4, font=('Microsoft YaHei UI', 10), justify=tk.CENTER)
            minute_spinbox.pack(side=tk.LEFT, padx=2)
            
            # 启用复选框
            enabled_var = tk.BooleanVar(value=True)
            enabled_checkbox = ttk.Checkbutton(frame, text="启用该通知", variable=enabled_var,
                                            style='TCheckbutton')
            enabled_checkbox.grid(row=3, column=1, sticky=tk.W, pady=(0, 20))
            
            # 保存按钮
            def save_new_notification():
                try:
                    title = title_var.get().strip()
                    message = text_widget.get("1.0", tk.END).strip()
                    hour = hour_var.get()
                    minute = minute_var.get()
                    
                    if not title:
                        self.show_windows_notification("嗯？", "标题不能为空")
                        return
                    
                    self.cursor.execute(
                        "INSERT INTO notifications (title, message, hour, minute, enabled) VALUES (?, ?, ?, ?, ?)",
                        (title, message, hour, minute, 1 if enabled_var.get() else 0)
                    )
                    self.conn.commit()
                    
                    # 刷新列表
                    populate_notification_list()
                    add_window.destroy()
                    self.show_windows_notification("搞定！", "新通知已添加")
                except Exception as e:
                    self.show_windows_notification("Oops！", str(e))
            
            # 按钮框架，居中显示保存按钮
            button_frame = ttk.Frame(frame)
            button_frame.grid(row=4, column=0, columnspan=3, pady=10)
            
            # 美化的保存按钮
            save_button = ttk.Button(button_frame, text="保存通知", command=save_new_notification,
                                    style='TButton', width=15)
            save_button.pack(pady=10)
        
        # 添加新通知按钮
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=(10, 0))
        
        add_button = ttk.Button(buttons_frame, text="添加新通知", command=lambda: add_notification())
        add_button.pack(side=tk.LEFT, padx=5)
        
        # 关闭按钮
        close_button = ttk.Button(buttons_frame, text="关闭", command=notification_window.destroy)
        close_button.pack(side=tk.RIGHT, padx=5)
        
        # 刷新按钮回调
        notification_window.bind("<<RefreshEvent>>", lambda e: populate_notification_list())
        
        # 编辑通知
        def edit_notification(notification_id):
            # 查找通知
            notification = None
            for n in self.notifications:
                if n['id'] == notification_id:
                    notification = n
                    break
            
            if not notification:
                return
            
            # 创建编辑窗口
            edit_window = tk.Toplevel(notification_window)
            edit_window.title("编辑通知")
            edit_window.geometry("400x320")  # 增大窗口尺寸，与添加窗口一致
            edit_window.resizable(False, False)
            
            # 创建框架
            frame = ttk.Frame(edit_window, padding="20")
            frame.pack(fill=tk.BOTH, expand=True)
            
            # 配置网格布局的列权重
            frame.columnconfigure(0, weight=0)
            frame.columnconfigure(1, weight=1)
            
            # 标题输入
            ttk.Label(frame, text="标题:", font=('Microsoft YaHei UI', 10)).grid(
                row=0, column=0, sticky=tk.W + tk.N, pady=(0, 10), padx=(0, 10))
            title_var = tk.StringVar(value=notification['title'])
            title_entry = ttk.Entry(frame, textvariable=title_var, width=35)
            title_entry.grid(row=0, column=1, sticky=tk.E + tk.W, pady=(0, 10))
            title_entry.focus_set()  # 自动聚焦到标题输入框
            
            # 消息输入
            ttk.Label(frame, text="消息:", font=('Microsoft YaHei UI', 10)).grid(
                row=1, column=0, sticky=tk.W + tk.N, pady=(0, 10), padx=(0, 10))
            text_widget = tk.Text(frame, height=6, width=30, font=('Microsoft YaHei UI', 10))
            text_widget.insert(tk.END, notification['message'])
            text_widget.grid(row=1, column=1, sticky=tk.E + tk.W + tk.N + tk.S, pady=(0, 10))
            
            # 添加消息框的垂直滚动条
            msg_scrollbar = ttk.Scrollbar(frame, orient="vertical", command=text_widget.yview)
            msg_scrollbar.grid(row=1, column=2, sticky=tk.N + tk.S, pady=(0, 10))
            text_widget.configure(yscrollcommand=msg_scrollbar.set)
            
            # 时间选择
            ttk.Label(frame, text="时间:", font=('Microsoft YaHei UI', 10)).grid(
                row=2, column=0, sticky=tk.W + tk.N, pady=(0, 15), padx=(0, 10))
            time_frame = ttk.Frame(frame)
            time_frame.grid(row=2, column=1, sticky=tk.W, pady=(0, 15))
            
            # 小时步进器
            hour_var = tk.IntVar(value=notification['hour'])
            hour_spinbox = ttk.Spinbox(time_frame, from_=0, to=23, textvariable=hour_var, 
                                    width=4, font=('Microsoft YaHei UI', 10), justify=tk.CENTER)
            hour_spinbox.pack(side=tk.LEFT, padx=2)
            
            # 分隔符
            ttk.Label(time_frame, text=":", font=('Microsoft YaHei UI', 12, 'bold')).pack(side=tk.LEFT)
            
            # 分钟步进器
            minute_var = tk.IntVar(value=notification['minute'])
            minute_spinbox = ttk.Spinbox(time_frame, from_=0, to=59, textvariable=minute_var,
                                        width=4, font=('Microsoft YaHei UI', 10), justify=tk.CENTER)
            minute_spinbox.pack(side=tk.LEFT, padx=2)
            
            # 启用复选框
            enabled_var = tk.BooleanVar(value=notification['enabled'])
            enabled_checkbox = ttk.Checkbutton(frame, text="启用该通知", variable=enabled_var,
                                            style='TCheckbutton')
            enabled_checkbox.grid(row=3, column=1, sticky=tk.W, pady=(0, 20))
            
            # 保存按钮
            def save_changes():
                try:
                    title = title_var.get().strip()
                    message = text_widget.get("1.0", tk.END).strip()
                    hour = hour_var.get()
                    minute = minute_var.get()
                    
                    if not title:
                        self.show_windows_notification("嗯？", "标题不能为空")
                        return
                    
                    self.cursor.execute(
                        "UPDATE notifications SET title = ?, message = ?, hour = ?, minute = ?, enabled = ? WHERE id = ?",
                        (title, message, hour, minute, 1 if enabled_var.get() else 0, notification_id)
                    )
                    self.conn.commit()
                    
                    # 刷新列表
                    populate_notification_list()
                    edit_window.destroy()
                    self.show_windows_notification("搞定！", "通知已更改")
                except Exception as e:
                    self.show_windows_notification("Oops！", str(e))
            
            # 按钮框架，居中显示保存按钮
            button_frame = ttk.Frame(frame)
            button_frame.grid(row=4, column=0, columnspan=3, pady=10)
            
            # 美化的保存按钮
            save_button = ttk.Button(button_frame, text="保存修改", command=save_changes,
                                    style='TButton', width=15)
            save_button.pack(pady=10)
        
    def toggle_getting_mode(self):
        """切换获取模式（名言/单词）"""
        if self.current_display_mode == "quote":
            self.current_display_mode = "word"
        else:
            self.current_display_mode = "quote"
        self.refresh_content()
    
    def get_quote(self):
        """从一言API获取名言并解析JSON数据"""
        try:
            response = requests.get("https://v1.hitokoto.cn/?c=k&encode=json", timeout=5)
            if response.status_code == 200:
                # 解析JSON数据
                data = response.json()
                hitokoto = data.get('hitokoto', '').strip()
                from_who = data.get('from_who', '')
                
                # 如果没有作者信息，尝试获取来源
                if not from_who:
                    from_who = data.get('from', '')
                    
                return hitokoto, from_who
            else:
                return "书山有路勤为径，学海无涯苦作舟。", "—— 韩愈"
        except Exception as e:
            self.show_windows_notification("获取名言失败", str(e))
            return "书山有路勤为径，学海无涯苦作舟。", "—— 韩愈"
            
    def update_quote(self):
        """在新线程中更新名言显示"""
        hitokoto, from_who = self.get_quote()
        
        # 处理显示格式
        quote_text = hitokoto
        
        # 来源文本格式化
        if from_who:
            from_text = f"—— {from_who}"
        else:
            from_text = "—— 侠名"
        
        # 在主线程中更新UI
        self.root.after(0, lambda: self.quote_label.config(text=quote_text))
        self.root.after(0, lambda: self.from_label.config(text=from_text))
        
    def get_english_word(self):
        """从英语词典API获取随机单词信息"""
        try:
            response = requests.get("https://oiapi.net/api/RandEnglishDict", timeout=5)
            if response.status_code == 200:
                # 解析JSON数据
                data = response.json()
                if data.get('code') == 1 and data.get('data'):
                    word_data = data['data']
                    word = word_data.get('content', '')
                    translation = word_data.get('trans', '')
                    
                    # 获取第一个例句
                    sentences = word_data.get('sentences', [])
                    example = ""
                    if sentences:
                        example = sentences[0].get('sContent', '')
                        example_trans = sentences[0].get('sCn', '')
                        if example and example_trans:
                            example = f"{example}\n{example_trans}"
                            # 如果example长度超过150字符，则重新获取单词
                            # print(f"获取到的例句长度: {len(example+word+translation)}")
                            if len(example+word+translation) > 120:
                                # print("例句过长，重新获取单词...")
                                return self.get_english_word()
                    
                    return word, translation, example
                else:
                    return "error", "获取数据失败", ""
            else:
                return "error", f"API返回错误: {response.status_code}", ""
        except Exception as e:
            # self.show_windows_notification("获取单词失败", str(e))
            return "error", str(e), ""
    
    def update_english_word(self):
        """更新英语单词显示"""
        word, translation, example = self.get_english_word()
        
        if word != "error":
            processed_example = example
            pattern = r'\b' + re.escape(word) + r'\b'
            # 使用sub替换单词，保留原始大小写
            processed_example = re.sub(pattern, f'「{word}」', example, flags=re.IGNORECASE)
            word_text = processed_example
            
            # 在主线程中更新UI
            self.root.after(0, lambda: self.quote_label.config(text=word_text))
            self.root.after(0, lambda: self.from_label.config(text=f"* {word}：{translation}"))
        else:
            # 失败时使用备用文本
            self.root.after(0, lambda: self.quote_label.config(text="No pain, no gain."))
            self.root.after(0, lambda: self.from_label.config(text="* 没有付出，就没有收获。"))

    def refresh_content(self, event=None):
        """刷新当前显示内容（名言或单词）"""
        # 先显示加载提示
        self.root.after(0, lambda: self.quote_label.config(text="获取中..."))
        self.root.after(0, lambda: self.from_label.config(text=""))

        # 在新线程中获取内容
        if self.current_display_mode == "quote":
            thread = threading.Thread(target=self.update_quote)
        else:
            thread = threading.Thread(target=self.update_english_word)
        thread.daemon = True
        thread.start()
        
    def show_about(self):
        """显示关于信息"""
        about_message = "桌面时钟倒计时组件-CountdownApp\n版本 1.6.0-260111\n开发：TiantianYZJ（yzjtiantian@126.com）\n\n感谢：\n- @zhy_0928_fc (←此人鬼点子多)\n- 一言API (https://hitokoto.cn/)\n- OIAPI (https://oiapi.net/)\n- 心知天气API (https://www.seniverse.com/)\n\n本软件遵循MIT开源协议，\n可在 GitHub 获取源代码\n（https://github.com/TiantianYZJ/CountdownApp/）。"
        messagebox.showinfo("关于", about_message)

    def quit(self):
        """退出程序"""
        self.show_windows_notification("这就叫丝滑~", "桌面时钟倒计时组件 已安全退出")

        # 关闭数据库连接
        if hasattr(self, 'conn'):
            self.conn.close()
            
        # 添加短暂延迟让通知显示，同时确保UI事件正常处理
        def delayed_quit():
            # 先更新一次UI以确保所有事件处理完毕
            self.root.update()
            # 短延迟
            self.root.after(500, lambda: self.root.destroy())  # 使用destroy，更彻底地关闭
        
        # 立即启动延迟退出流程
        self.root.after(0, delayed_quit)
        
    def create_context_menu(self):
        """创建右键菜单"""
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="每日一笑", command=self.show_joke_window)
        self.context_menu.add_command(label="AI绘画", command=self.show_ai_painting_window)
        self.context_menu.add_command(label="系统计算器", command=self.show_calculator_window)
        self.context_menu.add_command(label="添加桌面便签（NEW）", command=self.show_notepad_window)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="切换名言/单词", command=self.toggle_getting_mode)
        self.context_menu.add_command(label="重置迷你时钟位置", command=self.reset_mini_window_position)
        self.context_menu.add_command(label="重置主窗口位置大小", command=self.reset_main_window_position)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="关于", command=self.show_about)
        self.context_menu.add_command(label="退出", command=self.quit)
    
    def show_calculator_window(self):
        os.system("start calc")

    def reset_mini_window_position(self):
        """重置迷你窗口位置"""
        screen_width = self.root.winfo_screenwidth()
        mini_width = 360  # 小窗口宽度
        mini_height = 40  # 小窗口高度
        x = (screen_width - mini_width) // 5
        y = 1  # 距离顶部像素
        self.mini_window.geometry(f"{mini_width}x{mini_height}+{x}+{y}")
        self.mini_window_position_set = True
    
    def reset_main_window_position(self):
        """重置主窗口位置和大小"""
        self.root.geometry("950x600")
        self.set_window_position()
        
    def show_ai_painting_window(self):
        """显示AI绘画窗口"""
        # 创建AI绘画窗口
        ai_window = tk.Toplevel(self.root)
        ai_window.title("AI绘画")
        ai_window.geometry("860x750")
        ai_window.resizable(True, True)
        
        # 设置窗口图标
        if hasattr(sys, '_MEIPASS'):
            icon_path = os.path.join(sys._MEIPASS, 'logo.ico')
        else:
            icon_path = './logo.ico'
        if os.path.exists(icon_path):
            ai_window.iconbitmap(icon_path)

        # 创建主画布
        main_canvas = tk.Canvas(ai_window, highlightthickness=0)
        main_canvas.pack(fill=tk.BOTH, expand=True)

        # 创建输入区域 - 移除style参数
        input_frame = ttk.Frame(main_canvas, padding="20")
        input_frame_id = main_canvas.create_window(0, 0, window=input_frame, anchor="nw")

        # 标题 - 直接应用样式参数
        title_label = ttk.Label(input_frame, text="AI 绘画", foreground="black", font=("Microsoft YaHei", 14, "bold"))
        title_label.pack(side=tk.TOP, pady=(0, 10), fill=tk.X)

        # 技术提供说明（保持不变）
        self.tech_support_frame = ttk.Frame(ai_window)
        self.tech_support_frame.place(relx=1.0, rely=1.0, anchor="se", x=-10)

        # 加载图片文件，确保路径正确
        if hasattr(sys, '_MEIPASS'):
            img_path = os.path.join(sys._MEIPASS, 'OIAPI.png')
        else:
            img_path = './OIAPI.png'
        if os.path.exists(img_path):
            self.tech_logo = tk.PhotoImage(file=img_path)

        # 创建图片标签
        logo_label = ttk.Label(self.tech_support_frame, image=self.tech_logo)
        logo_label.pack(side="left", padx=5)

        # 创建文字标签
        support_label = ttk.Label(self.tech_support_frame, text="提供技术支持，AI生成仅供参考", font=("Microsoft YaHei", 12))
        support_label.pack(side="bottom", pady=(0, 15))

        # 提示词输入 - 直接应用样式参数
        prompt_label = ttk.Label(input_frame, text="提示词:", foreground="black", font=("Microsoft YaHei", 10))
        prompt_label.pack(side=tk.TOP, anchor=tk.W, pady=(0, 5))

        self.prompt_var = tk.StringVar(value="")
        prompt_entry = ttk.Entry(input_frame, textvariable=self.prompt_var, font=("Microsoft YaHei", 10))
        prompt_entry.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))

        # 创建风格和尺寸选择的水平框架
        style_size_frame = ttk.Frame(input_frame)
        style_size_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 10))

        # 风格选择 - 直接应用样式参数
        style_label = ttk.Label(style_size_frame, text="绘画风格:", foreground="black", font=("Microsoft YaHei", 10))
        style_label.pack(side=tk.LEFT, anchor=tk.CENTER, padx=(0, 5))

        self.style_var = tk.IntVar(value=100)
        style_values = {100: "智能匹配", 1: "现代都市", 2: "古风武侠", 3: "水墨国风", 4: "梦幻异世", 5: "现代日漫"}
        style_combobox = ttk.Combobox(style_size_frame, values=list(style_values.values()), width=15)
        style_combobox.current(0)  # 默认选择智能匹配
        style_combobox.pack(side=tk.LEFT, anchor=tk.CENTER, padx=(0, 20))

        # 尺寸选择 - 直接应用样式参数
        size_label = ttk.Label(style_size_frame, text="图片尺寸(宽高比):", foreground="black", font=("Microsoft YaHei", 10))
        size_label.pack(side=tk.LEFT, anchor=tk.CENTER, padx=(0, 5))

        self.size_var = tk.IntVar(value=3)
        size_values = {1: "1:1", 2: "5:7", 3: "9:16"}
        size_combobox = ttk.Combobox(style_size_frame, values=list(size_values.values()), width=15)
        size_combobox.current(2)  # 默认选择1080*1512
        size_combobox.pack(side=tk.LEFT, anchor=tk.CENTER)

        # 智能补充复选框
        self.llm_var = tk.BooleanVar(value=True)
        llm_checkbox = ttk.Checkbutton(input_frame, text="智能补充提示词", variable=self.llm_var)
        llm_checkbox.pack(side=tk.TOP, anchor=tk.W, pady=(0, 10))

        # 生成按钮 - 直接应用样式参数
        generate_button = ttk.Button(input_frame, text="生成图片*4", 
                                command=lambda: self.generate_ai_image(ai_window, style_values, style_combobox, size_values, size_combobox))
        generate_button.pack(side=tk.TOP, pady=(0, 10))

        # 创建加载状态标签
        self.loading_label = ttk.Label(input_frame, text="", foreground="#008cff", font=("Microsoft YaHei", 10))
        self.loading_label.pack(side=tk.TOP, anchor=tk.W, pady=(0, 5))

        # 创建图片显示区域 - 移除style参数
        self.image_frame = ttk.Frame(main_canvas, padding="5")
        self.image_frame_id = main_canvas.create_window(0, 250, window=self.image_frame, anchor="nw")

        # 添加画布大小变化的绑定事件，实现横向自适应
        def on_canvas_configure(event):
            main_canvas.itemconfig(input_frame_id, width=event.width)
            main_canvas.itemconfig(self.image_frame_id, width=event.width)
        
        main_canvas.bind("<Configure>", on_canvas_configure)

        # 预创建四个图片显示区域和保存按钮
        self.precreated_image_widgets = []
        for i in range(4):
            # 创建图片框架
            img_frame = ttk.Frame(self.image_frame, padding="10")
            img_frame.pack(side="left", fill="y", padx=5, pady=5, expand=True)
            
            # 显示占位文本
            placeholder_label = ttk.Label(img_frame, text="图片待生成", font=("Microsoft YaHei", 12))
            placeholder_label.pack()
            
            # 保存按钮（初始隐藏）
            save_button = ttk.Button(img_frame, text="保存图片", state="disabled")
            save_button.pack(pady=5)
            
            # 存储组件引用
            self.precreated_image_widgets.append({
                "frame": img_frame,
                "image_label": placeholder_label,
                "save_button": save_button,
                "image_reference": None  # 用于保存图片引用，防止被垃圾回收
            })

        # 绑定组合框到变量的更新函数
        def update_style_var(event):
            selected_text = style_combobox.get()
            for key, value in style_values.items():
                if value == selected_text:
                    self.style_var.set(key)
                    break

        def update_size_var(event):
            selected_text = size_combobox.get()
            for key, value in size_values.items():
                if value == selected_text:
                    self.size_var.set(key)
                    break

        style_combobox.bind("<<ComboboxSelected>>", update_style_var)
        size_combobox.bind("<<ComboboxSelected>>", update_size_var)

        # 窗口关闭时的清理
        def on_close():
            # 清理可能的加载线程
            ai_window.destroy()

        ai_window.protocol("WM_DELETE_WINDOW", on_close)

    def generate_ai_image(self, ai_window, style_values, style_combobox, size_values, size_combobox):
        """生成AI图片"""
        # 更新选择的值
        selected_style_text = style_combobox.get()
        for key, value in style_values.items():
            if value == selected_style_text:
                self.style_var.set(key)
                break
        
        selected_size_text = size_combobox.get()
        for key, value in size_values.items():
            if value == selected_size_text:
                self.size_var.set(key)
                break
        
        # 显示加载状态
        self.loading_label.config(text="正在生成图片，请稍候...")
        
        # 在新线程中生成图片
        def generate_task():
            try:
                url = "https://oiapi.net/api/AiDrawImage"
                params = {
                    "prompt": self.prompt_var.get(),
                    "style": self.style_var.get(),
                    "size": self.size_var.get(),
                    "llm": str(self.llm_var.get()).lower(),
                    "type": "json"
                }
                
                response = requests.get(url, params=params, timeout=30)
                
                if response.status_code == 200:
                    data = response.json()
                    
                    if data.get('code') == 1 and data.get('data'):
                        # 清除加载状态
                        ai_window.after(0, lambda: self.loading_label.config(text="生成成功！"))
                        
                        # 移除冲突的清除和创建滚动区域的代码
                        # 直接使用预创建的组件来显示图片
                        
                        # 显示生成的图片（直接替换预创建组件的内容）
                        from PIL import Image, ImageTk
                        import io
                        
                        # 确保有足够的预创建组件
                        for i, (img_data, widget_set) in enumerate(zip(data['data'], self.precreated_image_widgets)):
                            try:
                                # 下载并显示图片
                                img_response = requests.get(img_data['url'], timeout=10)
                                img_data_bytes = img_response.content
                                image = Image.open(io.BytesIO(img_data_bytes))
                                
                                # 调整图片大小以适应显示
                                max_width = 200
                                width, height = image.size
                                ratio = max_width / width
                                new_height = int(height * ratio)
                                resized_image = image.resize((max_width, new_height), Image.LANCZOS)
                                
                                photo = ImageTk.PhotoImage(resized_image)
                                
                                # 清除原有的占位标签
                                widget_set["image_label"].destroy()
                                
                                # 创建新的图片标签
                                img_label = ttk.Label(widget_set["frame"], image=photo)
                                img_label.image = photo  # 保持引用
                                img_label.pack()
                                
                                # 更新存储的引用
                                widget_set["image_label"] = img_label
                                widget_set["image_reference"] = photo
                                
                                # 启用保存按钮
                                widget_set["save_button"].config(state="normal")
                                widget_set["save_button"].config(command=lambda img=image: self.save_ai_image(img))
                                
                            except Exception as img_error:
                                # 如果加载失败，显示错误信息
                                widget_set["image_label"].destroy()
                                
                                error_label = ttk.Label(widget_set["frame"], text=f"图片{i+1}加载失败", foreground="red", font=("Microsoft YaHei", 12))
                                print(f"图片{i+1}加载失败: {img_error}")
                                error_label.pack()
                                
                                widget_set["image_label"] = error_label
                                widget_set["save_button"].config(state="disabled")
                
                        
                    else:
                        ai_window.after(0, lambda: self.loading_label.config(text=f"生成失败: {data.get('message', '未知错误')}"))
                else:
                    ai_window.after(0, lambda: self.loading_label.config(text=f"API请求失败: {response.status_code}"))
                    
            except Exception as e:
                import traceback
                error_message = str(e)
                ai_window.after(0, lambda: self.loading_label.config(text=f"生成失败: {error_message}"))
        
        # 启动生成任务
        thread = threading.Thread(target=generate_task)
        thread.daemon = True
        thread.start()

    def save_ai_image(self, image):
        """保存AI生成的图片"""
        try:
            # 获取当前日期时间作为默认文件名
            import datetime
            now = datetime.datetime.now()
            default_filename = f"ai_drawing_{now.strftime('%Y%m%d_%H%M%S')}.png"
            
            # 打开保存对话框
            file_path = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG files", "*.png"), ("All files", "*.*")],
                initialfile=default_filename
            )
            
            if file_path:
                # 保存图片
                image.save(file_path)
                self.show_windows_notification("保存成功", f"图片已保存至:\n{file_path}")
                
        except Exception as e:
            self.show_windows_notification("保存失败", str(e))

    # 美化show_joke_window方法
    def show_joke_window(self):
        """显示笑话窗口"""
        # 创建笑话窗口
        joke_window = tk.Toplevel(self.root)
        joke_window.title("每日一笑")
        joke_window.geometry("600x450")
        joke_window.resizable(False, False)
        
        # 设置窗口图标
        if hasattr(sys, '_MEIPASS'):
            icon_path = os.path.join(sys._MEIPASS, 'logo.ico')
        else:
            icon_path = './logo.ico'
        if os.path.exists(icon_path):
            joke_window.iconbitmap(icon_path)
        
        # 创建ttk样式
        style = ttk.Style(joke_window)
        
        # 设置自定义样式
        style.configure("JokeTitle.TFrame", background="#2c3e50")
        style.configure("JokeTitle.TLabel", background="#2c3e50", foreground="white", font=("Microsoft YaHei", 14, "bold"))
        style.configure("TypeLabel.TLabel", background="#2c3e50", foreground="white", font=("Microsoft YaHei", 10))
        style.configure("Content.TFrame", background="#f8f9fa")
        style.configure("Joke.TLabel", background="#f8f9fa", foreground="#000000", font=("Microsoft YaHei", 16), wraplength=500, justify="center")
        style.configure("Button.TFrame", background="#ecf0f1")
        style.configure("Load.TButton", background="#2894ff", font=("Microsoft YaHei", 11))
        style.configure("Like.TButton", font=("Microsoft YaHei", 11))
        style.configure("Liked.TButton", background="#ffc400", foreground="#c29500", font=("Microsoft YaHei", 11))
        style.configure("Edit.TButton", font=("Microsoft YaHei", 11))
        style.configure("Exit.TButton", background="#ff0000", font=("Microsoft YaHei", 11))
        style.configure("Info.TFrame", height=30)
        style.configure("Info.TLabel", foreground="#7f8c8d", font=("Microsoft YaHei", 10))
        
        # 创建圆角效果的画布来作为窗口背景
        main_canvas = tk.Canvas(joke_window, bg="#ecf0f1", highlightthickness=0)
        main_canvas.pack(fill="both", expand=True)
        
        # 创建标题栏（自定义）
        title_frame = ttk.Frame(main_canvas, style="JokeTitle.TFrame", height=50)
        title_frame_id = main_canvas.create_window(0, 0, window=title_frame, anchor="nw", width=600)
        
        # 标题文本
        title_label = ttk.Label(title_frame, text="每日一笑", style="JokeTitle.TLabel")
        title_label.pack(side="left", padx=20, pady=12)
        
        # 添加笑话类型选择下拉列表
        self.joke_type_var = tk.StringVar(value="弱智吧")
        joke_type_frame = ttk.Frame(title_frame, style="Title.TFrame")
        joke_type_frame.pack(side="right", padx=20, pady=10)
        
        joke_type_label = ttk.Label(joke_type_frame, text="笑话类型:", style="TypeLabel.TLabel")
        joke_type_label.pack(side="left", padx=(0, 5))
        
        # 使用ttk.Combobox替代OptionMenu
        joke_type_combobox = ttk.Combobox(joke_type_frame, textvariable=self.joke_type_var, values=["弱智吧", "毒鸡汤"], width=8)
        joke_type_combobox.bind("<<ComboboxSelected>>", lambda _: self.update_joke(joke_window))
        joke_type_combobox.pack(side="left")
        joke_type_combobox.current(0)  # 默认选择第一项
        
        # 创建内容区域（卡片式设计）
        content_frame = ttk.Frame(main_canvas, style="Content.TFrame", padding="20")
        content_frame_id = main_canvas.create_window(0, 50, window=content_frame, anchor="nw", width=600, height=300)
        
        # 创建笑话内容标签（居中显示）
        self.joke_label = ttk.Label(content_frame, 
                                  text="加载中...", 
                                  style="Joke.TLabel",
                                  padding=(30, 40))
        self.joke_label.pack(expand=True, fill="both")
        
        # 创建按钮区域
        button_frame = ttk.Frame(main_canvas, style="Button.TFrame", height=60)
        button_frame_id = main_canvas.create_window(0, 350, window=button_frame, anchor="nw", width=600)
        
        # 创建刷新按钮
        refresh_btn = ttk.Button(button_frame, 
                               text="🔄 刷新", 
                               style="Load.TButton",
                               padding=(10, 5),
                               command=lambda: self.update_joke(joke_window))
        refresh_btn.pack(side="left", padx=20, pady=10)
        
        # 创建收藏按钮
        self.favorite_btn = ttk.Button(button_frame, 
                                     text="⭐ 收藏", 
                                     style="Like.TButton",
                                     padding=(10, 5),
                                     command=lambda: self.add_to_favorite(joke_window))
        self.favorite_btn.pack(side="left", padx=10, pady=10)
        
        # 创建查看收藏按钮
        view_favorite_btn = ttk.Button(button_frame, 
                                     text="📋 查看收藏", 
                                     style="Edit.TButton",
                                     padding=(10, 5),
                                     command=self.show_favorite_window)
        view_favorite_btn.pack(side="left", padx=10, pady=10)
        
        # 创建关闭按钮
        close_btn = ttk.Button(button_frame, 
                             text="关闭", 
                             style="Exit.TButton",
                             padding=(10, 5),
                             command=joke_window.destroy)
        close_btn.pack(side="right", padx=20, pady=10)
        
        # 创建信息显示区域
        info_frame = ttk.Frame(main_canvas, style="Info.TFrame", height=30)
        info_frame_id = main_canvas.create_window(0, 410, window=info_frame, anchor="nw", width=600)
        
        # 创建笑话计数显示标签
        self.joke_count_label = ttk.Label(info_frame, 
                                        text="累计显示 0 条", 
                                        style="Info.TLabel")
        self.joke_count_label.pack(side="right", padx=20, pady=5)
        
        # 免责声明
        disclaimer_label = ttk.Label(info_frame, 
                                   text="免责声明：互联网内容仅供参考，不代表任何个人立场", 
                                   style="Info.TLabel")
        disclaimer_label.pack(side="left", padx=20, pady=5)
        
        # 初始化笑话收藏状态
        self.current_joke = None
        
        # 首次加载笑话
        self.update_joke(joke_window)
        
        # 显示当前计数
        self.update_joke_count_display(joke_window)
    
    # 优化update_joke方法，添加加载动画效果
    def update_joke(self, window):
        """从API获取笑话并在新线程中更新显示"""
        # 显示加载提示
        self.loading_index = 0
        self.loading_states = ["联网获取中   ", "联网获取中.  ", "联网获取中.. ", "联网获取中..."]
        
        def update_loading_text():
            self.loading_index = (self.loading_index + 1) % 4
            self.root.after(0, lambda: self.joke_label.config(text=self.loading_states[self.loading_index]))
            self.loading_timer = self.root.after(300, update_loading_text)
        
        # 启动加载动画
        self.root.after(0, lambda: self.joke_label.config(text=self.loading_states[0]))
        self.loading_timer = self.root.after(300, update_loading_text)
        
        # 在新线程中获取笑话，避免阻塞UI
        thread = threading.Thread(target=lambda: self.fetch_joke(window))
        thread.daemon = True
        thread.start()
    
    # 修改fetch_joke方法，确保清除加载动画
    def fetch_joke(self, window):
        """从API获取笑话并解析JSON数据"""
        try:
            # 获取选中的笑话类型
            joke_type = self.joke_type_var.get() if hasattr(self, 'joke_type_var') else "弱智吧"
            
            # 根据笑话类型选择API
            if joke_type == "弱智吧":
                api_url = "https://www.7ed.net/ruozi/api"
                content_key = 'ruozi'
            elif joke_type == "毒鸡汤":
                api_url = "https://www.7ed.net/soup/api"
                content_key = 'badsoup'
            else:
                api_url = "https://www.7ed.net/ruozi/api"
                content_key = 'ruozi'
            
            # 调用笑话API
            response = requests.get(api_url, timeout=5)
            if response.status_code == 200:
                # 解析JSON数据
                data = response.json()
                joke_content = data.get(content_key, '获取失败').strip()
                
                # 保存当前笑话信息
                self.current_joke = {
                    'type': joke_type,
                    'content': joke_content
                }
                
                # 检查是否已收藏
                self.check_favorite_status()
                
                # 取消加载动画并更新UI
                self.root.after(0, lambda: self.root.after_cancel(self.loading_timer))
                self.root.after(0, lambda: self.joke_label.config(text=joke_content))
                
                # 更新笑话计数器
                self.update_joke_counter()
                
                # 刷新显示计数
                self.root.after(0, lambda: self.update_joke_count_display(window))
            else:
                # 请求失败处理
                self.root.after(0, lambda: self.joke_label.config(text=f"获取{joke_type}内容失败，请稍后重试"))
        except Exception as e:
            self.show_windows_notification("获取内容失败", str(e))
            # 异常处理
            self.root.after(0, lambda: self.root.after_cancel(self.loading_timer))
            self.root.after(0, lambda: self.joke_label.config(text="网络连接异常，请检查网络后重试"))
    
    # 添加更新笑话计数器的方法
    def update_joke_counter(self):
        """更新笑话计数器"""
        
        # 获取收藏文件路径
        favorite_file = os.path.join(self.get_appdata_path(), 'jokes.json')
        
        try:
            # 读取现有收藏
            if os.path.exists(favorite_file):
                with open(favorite_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
            else:
                data = []
            
            # 检查是否已有计数器信息
            counter_info = next((item for item in data if isinstance(item, dict) and item.get('id') == 'counter'), None)
            
            if counter_info:
                # 更新现有计数器
                counter_info['count'] = counter_info.get('count', 0) + 1
                counter_info['last_update'] = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            else:
                # 创建新的计数器
                counter_info = {
                    'id': 'counter',
                    'count': 1,
                    'last_update': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                }
                data.insert(0, counter_info)
            
            # 保存更新后的数据
            with open(favorite_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            self.show_windows_notification("更新计数器失败", str(e))

    # 添加获取笑话计数的方法
    def get_joke_count(self):
        """获取累计显示的笑话数量"""
        
        # 获取收藏文件路径
        favorite_file = os.path.join(self.get_appdata_path(), 'jokes.json')
        
        try:
            if os.path.exists(favorite_file):
                with open(favorite_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    counter_info = next((item for item in data if isinstance(item, dict) and item.get('id') == 'counter'), None)
                    if counter_info:
                        return counter_info.get('count', 0)
            return 0
        except Exception as e:
            self.show_windows_notification("获取计数失败", str(e))
            return 0

    # 添加更新笑话计数显示的方法
    def update_joke_count_display(self, window):
        """更新笑话窗口中的计数显示"""
        count = self.get_joke_count()
        if hasattr(self, 'joke_count_label'):
            self.joke_count_label.config(text=f"累计显示 {count} 条")

    def add_to_favorite(self, window):
        """将当前笑话添加到收藏"""
        if not self.current_joke:
            return
        
        # 获取收藏文件路径
        favorite_file = os.path.join(self.get_appdata_path(), 'jokes.json')
        
        # 读取现有收藏
        favorites = []
        if os.path.exists(favorite_file):
            try:
                with open(favorite_file, 'r', encoding='utf-8') as f:
                    favorites = json.load(f)
            except:
                favorites = []

        # 分离计数器信息和实际收藏项
        counter_info = None
        actual_favorites = []
        for item in favorites:
            if isinstance(item, dict) and item.get('id') == 'counter':
                counter_info = item
            else:
                actual_favorites.append(item)
        
        # 检查是否已收藏
        is_already_favorite = any(fav.get('content') == self.current_joke['content'] and fav.get('type') == self.current_joke['type'] for fav in actual_favorites if 'content' in fav and 'type' in fav)
        
        if not is_already_favorite:
            # 添加新收藏项
            new_favorite = {
                'id': len(favorites) + 1,
                'type': self.current_joke['type'],
                'content': self.current_joke['content'],
                'date': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            }
            actual_favorites.append(new_favorite)
            
            # 重新组合数据（计数器+收藏项）
            updated_favorites = []
            if counter_info:
                updated_favorites.append(counter_info)
            updated_favorites.extend(actual_favorites)
            
            # 保存到文件
            try:
                with open(favorite_file, 'w', encoding='utf-8') as f:
                    json.dump(updated_favorites, f, ensure_ascii=False, indent=4)
                
                # 更新收藏按钮状态
                self.root.after(0, lambda: self.update_favorite_button(True))
            except Exception as e:
                self.show_windows_notification("收藏失败", str(e))
        else:
            # 移除已收藏的笑话
            filtered_favorites = [fav for fav in favorites if not (('content' in fav and 'type' in fav and 'content' in self.current_joke and 'type' in self.current_joke) and (fav['content'] == self.current_joke['content'] and fav['type'] == self.current_joke['type']))]
            
            # 重新编号（仅对非计数器项）
            counter_item = None
            actual_items = []
            for fav in filtered_favorites:
                if isinstance(fav, dict) and fav.get('id') == 'counter':
                    counter_item = fav
                else:
                    actual_items.append(fav)
            
            # 为实际收藏项重新编号
            for i, fav in enumerate(actual_items):
                fav['id'] = i + 1
            
            # 重新组合数据（计数器+收藏项）
            updated_favorites = []
            if counter_item:
                updated_favorites.append(counter_item)
            updated_favorites.extend(actual_items)
            
            # 保存到文件
            try:
                with open(favorite_file, 'w', encoding='utf-8') as f:
                    json.dump(updated_favorites, f, ensure_ascii=False, indent=4)
                
                # 更新收藏按钮状态
                self.root.after(0, lambda: self.update_favorite_button(False))
            except Exception as e:
                self.show_windows_notification("取消收藏失败", str(e))
    def update_favorite_button(self, is_favorited):
        """统一处理收藏按钮的状态更新"""
        if not hasattr(self, 'favorite_btn') or not self.favorite_btn.winfo_exists():
            return
        
        self.favorite_btn.config(text="⭐ 已收藏" if is_favorited else "⭐ 收藏", style="Liked.TButton" if is_favorited else "Like.TButton")

    def check_favorite_status(self):
        """检查当前笑话是否已收藏"""
        if not self.current_joke:
            return
            
        # 获取收藏文件路径
        favorite_file = os.path.join(self.get_appdata_path(), 'jokes.json')
        
        # 读取现有收藏
        favorites = []
        if os.path.exists(favorite_file):
            try:
                with open(favorite_file, 'r', encoding='utf-8') as f:
                    favorites = json.load(f)
            except:
                favorites = []

        # 过滤计数器
        actual_favorites = [item for item in favorites if isinstance(item, dict) and 'content' in item and 'type' in item]
        
        # 检查是否已收藏
        is_already_favorite = False
        if self.current_joke and 'content' in self.current_joke and 'type' in self.current_joke:
            is_already_favorite = any(fav['content'] == self.current_joke['content'] and fav['type'] == self.current_joke['type'] for fav in actual_favorites)
        
        # 更新收藏按钮状态 - 使用ttk样式方法
        self.root.after(0, lambda: self.update_favorite_button(is_already_favorite))

    def show_favorite_window(self):
        # 创建收藏窗口
        favorite_window = tk.Toplevel(self.root)
        favorite_window.title("我的收藏")
        favorite_window.geometry("800x600")
        
        # 设置ttk主题样式
        style = ttk.Style(favorite_window)
        style.configure("Treeview.Heading", font=('Microsoft YaHei', 10, 'bold'), foreground='#333333', background='#f0f0f0')
        style.configure("Treeview", font=('Microsoft YaHei', 10), rowheight=30)
        style.configure("Like.TLabel", font=('Microsoft YaHei', 14, 'bold'), foreground='#333333')
        style.configure("Empty.TLabel", font=('Microsoft YaHei', 12), foreground='#999999')
        
        # 获取收藏文件路径
        favorite_file = os.path.join(self.get_appdata_path(), 'jokes.json')

        # 读取现有收藏
        favorites = []
        if os.path.exists(favorite_file):
            try:
                with open(favorite_file, 'r', encoding='utf-8') as f:
                    favorites = json.load(f)
            except:
                favorites = []

        # 过滤计数器
        actual_favorites = [item for item in favorites if isinstance(item, dict) and 'content' in item and 'type' in item]
        
        # 创建主框架
        main_frame = ttk.Frame(favorite_window)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # 创建标题
        title_label = ttk.Label(main_frame, text=f"我的收藏（{len(actual_favorites)}）", style="Like.TLabel")
        title_label.pack(fill="x", pady=(0, 20))
        
        # 创建表格框架
        table_frame = ttk.Frame(main_frame)
        table_frame.pack(fill="both", expand=True)
        
        # 创建滚动条
        scrollbar = ttk.Scrollbar(table_frame)
        scrollbar.pack(side="right", fill="y")
        
        # 创建Treeview作为表格
        tree = ttk.Treeview(table_frame, yscrollcommand=scrollbar.set, show="headings", height=12)
        
        # 定义列
        tree['columns'] = ('id', 'type', 'content', 'date')
        
        # 设置列属性
        tree.column('id', width=20, anchor='center')
        tree.column('type', width=50, anchor='center')
        tree.column('content', width=450, anchor='w')
        tree.column('date', width=150, anchor='center')
        
        # 设置表头
        tree.heading('id', text='序号')
        tree.heading('type', text='板块')
        tree.heading('content', text='内容')
        tree.heading('date', text='收藏时间')
        
        # 绑定滚动条
        scrollbar.config(command=tree.yview)
        
        # 填充表格内容
        for i, favorite in enumerate(actual_favorites):
            # 在Treeview中插入一行数据
            item_id = tree.insert('', 'end', values=(favorite['id']-1, favorite['type'], favorite['content'], favorite['date'], ''))
            
            # 这里我们使用tag来标记每一行，便于后续操作
            tree.item(item_id, tags=('favorite_row',))
        
        # 如果需要实现行选择和删除功能，可以添加以下代码
        # 创建一个删除按钮
        if actual_favorites:
            delete_frame = ttk.Frame(main_frame)
            delete_frame.pack(fill="x", pady=(10, 0))
            
            delete_btn = ttk.Button(delete_frame, text="删除选中", command=lambda: self.delete_selected_favorite(tree, favorite_window))
            delete_btn.pack(side="right")
        
        # 如果没有收藏内容，显示提示
        if not actual_favorites:
            # 创建一个空数据提示框架
            empty_frame = ttk.Frame(table_frame)
            empty_frame.pack(fill="both", expand=True)
            
            # 使用ttk.Label显示提示信息
            no_data_label = ttk.Label(empty_frame, text="暂无收藏内容", style="Empty.TLabel")
            no_data_label.pack(pady=40)
        else:
            # 只有当有数据时才显示treeview
            tree.pack(fill="both", expand=True)
        
    def delete_selected_favorite(self, tree, favorite_window):
        """删除选中的收藏内容"""
        # 获取选中的项目
        selected_items = tree.selection()
        
        if not selected_items:
            self.show_windows_notification("请先选择要删除的内容", "")
            return
        
        # 获取收藏文件路径
        favorite_file = os.path.join(self.get_appdata_path(), 'jokes.json')
        
        # 读取现有收藏
        favorites = []
        if os.path.exists(favorite_file):
            try:
                with open(favorite_file, 'r', encoding='utf-8') as f:
                    favorites = json.load(f)
            except:
                favorites = []
        
        # 过滤有效的收藏
        actual_favorites = [item for item in favorites if isinstance(item, dict) and 'content' in item and 'type' in item]
        
        # 收集要删除的id
        to_delete_ids = []
        for item in selected_items:
            item_id = int(tree.item(item, 'values')[0]) + 1  # 转换回原始id
            to_delete_ids.append(item_id)
        
        # 删除选中的项目
        new_favorites = [item for item in actual_favorites if item['id'] not in to_delete_ids]
        
        # 分离计数器和实际收藏项
        counter_item = None
        actual_items = []
        for fav in favorites:
            if isinstance(fav, dict) and fav.get('id') == 'counter':
                counter_item = fav
            else:
                # 只保留不在删除列表中的实际收藏项
                if isinstance(fav, dict) and 'content' in fav and 'type' in fav and fav['id'] not in to_delete_ids:
                    actual_items.append(fav)

        # 只为实际收藏项重新编号
        for i, fav in enumerate(actual_items):
            fav['id'] = i + 1

        # 重新组合数据
        updated_favorites = []
        if counter_item:
            updated_favorites.append(counter_item)
        updated_favorites.extend(actual_items)

        # 保存更新后的收藏 - 使用包含计数器的updated_favorites
        try:
            with open(favorite_file, 'w', encoding='utf-8') as f:
                json.dump(updated_favorites, f, ensure_ascii=False, indent=4)
            
            # 修复：重新创建并刷新Treeview，确保序号正确
            # 先清空现有数据
            for item in tree.get_children():
                tree.delete(item)
            
            # 重新加载数据到Treeview
            for i, favorite in enumerate(updated_favorites):
                # 跳过计数器项
                if favorite.get('id') == 'counter':
                    continue
                tree.insert('', 'end', values=(favorite['id']-1, favorite['type'], favorite['content'], favorite['date'], ''))
            
             # 更新标题中的数量和计数标签
            for child in favorite_window.winfo_children():
                if isinstance(child, ttk.Frame):  # 遍历主框架
                    for sub_child in child.winfo_children():
                        if isinstance(sub_child, ttk.Label):
                            # 更新标题中的数量
                            if "我的收藏" in str(sub_child['text']):
                                sub_child.config(text=f"我的收藏（{len(actual_items)}）")
            
            # 检查是否还有数据，如果没有则显示空数据提示
            if not actual_items:
                # 找到表格框架
                table_frame = None
                for child in favorite_window.winfo_children():
                    if isinstance(child, ttk.Frame):  # 遍历主框架
                        for sub_child in child.winfo_children():
                            if "table_frame" in str(sub_child):
                                table_frame = sub_child
                                break
                
                if table_frame:
                    # 清空表格框架内容
                    for child in table_frame.winfo_children():
                        child.destroy()
                    
                    # 显示空数据提示
                    empty_frame = ttk.Frame(table_frame)
                    empty_frame.pack(fill="both", expand=True)
                    no_data_label = ttk.Label(empty_frame, text="暂无收藏内容", style="Empty.TLabel")
                    no_data_label.pack(pady=40)
                
                # 删除删除按钮框架
                for child in favorite_window.winfo_children():
                    if isinstance(child, ttk.Frame):  # 遍历主框架
                        for sub_child in child.winfo_children():
                            if "delete_frame" in str(sub_child):
                                sub_child.destroy()
                                break 
        except Exception as e:
            self.show_windows_notification("删除失败", str(e))

    def show_schedule_window(self):
        """显示完整课程表窗口，支持编辑"""
        # 创建课程表窗口
        schedule_window = tk.Toplevel(self.root)
        schedule_window.title("课程表")
        schedule_window.geometry("950x600")
        schedule_window.configure(bg="#f0f0f0")
        
        # 设置窗口居中显示
        schedule_window.update_idletasks()
        width = schedule_window.winfo_width()
        height = schedule_window.winfo_height()
        x = (schedule_window.winfo_screenwidth() // 2) - (width // 2)
        y = (schedule_window.winfo_screenheight() // 2) - (height // 2)
        schedule_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))

        # 主窗口
        main_frame = ttk.Frame(schedule_window, padding="20")
        main_frame.pack(fill="both", expand=True)
        
        
        # 创建标题
        title_label = ttk.Label(main_frame, text="课程表", font=('Microsoft YaHei UI', 16, 'bold'))
        title_label.pack(fill="x", pady=(0, 20))
        
        # 创建课程表框架
        schedule_frame = ttk.Frame(main_frame)
        schedule_frame.pack(fill="both", expand=True)
        
        # 获取星期几和时间段
        weekdays = list(self.schedule_data.get('school_days', {}).keys())
        time_slots = self.schedule_data.get('time_slots', [])
        
        # 确保weekdays是按顺序排列的
        weekday_order = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        sorted_weekdays = [day for day in weekday_order if day in weekdays]
        
        # 创建表头（时间段）
        time_header = ttk.Label(schedule_frame, text="时间段", font=('Microsoft YaHei UI', 12, 'bold'), width=15, anchor="center")
        time_header.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        
        # 创建表头（星期几）
        for col_idx, weekday in enumerate(sorted_weekdays):
            # 获取星期的中文名称
            weekday_dict = {'Monday': '周一', 'Tuesday': '周二', 'Wednesday': '周三', 
                           'Thursday': '周四', 'Friday': '周五', 'Saturday': '周六', 'Sunday': '周日'}
            weekday_zh = weekday_dict.get(weekday, weekday)
            
            header = ttk.Label(schedule_frame, text=weekday_zh, font=('Microsoft YaHei UI', 12, 'bold'), width=10, anchor="center")
            header.grid(row=0, column=col_idx+1, sticky="nsew", padx=1, pady=1)
        
        # 存储所有课程标签的引用
        self.course_labels = {}
        
        # 显示每节课
        for row_idx, time_slot in enumerate(time_slots):
            slot_id = time_slot['slot_id']
            start_time = time_slot['start_time']
            end_time = time_slot['end_time']
            
            # 显示时间段
            time_label = ttk.Label(schedule_frame, text=f"{start_time}-{end_time}", font=('Microsoft YaHei UI', 10), width=10, anchor="center")
            time_label.grid(row=row_idx+1, column=0, sticky="nsew", padx=1, pady=1)
            
            # 显示每一天的课程
            for col_idx, weekday in enumerate(sorted_weekdays):
                # 获取该时间段的课程
                course_name = ""
                if weekday in self.schedule_data['school_days']:
                    for course in self.schedule_data['school_days'][weekday]:
                        if course['slot_id'] == slot_id:
                            course_name = course['name']
                            break
                
                # 创建课程标签，支持点击编辑
                label_frame = ttk.Frame(schedule_frame, padding=(5, 5))
                label_frame.grid(row=row_idx+1, column=col_idx+1, sticky="nsew", padx=1, pady=1)
                
                # 创建点击事件处理器函数
                def on_click(event, wd=weekday, slot=slot_id):
                    self.on_course_click(event, wd, slot)
                
                # 为标签框架绑定点击事件
                label_frame.bind("<Button-1>", on_click)
                
                # 创建课程标签
                course_label = ttk.Label(label_frame, text=course_name, font=('Microsoft YaHei UI', 10), width=10, anchor="center", wraplength=100)
                course_label.pack(fill="both", expand=True)
                
                # 为课程标签也绑定点击事件，确保点击标签本身也能触发编辑
                course_label.bind("<Button-1>", on_click)
                
                # 存储标签引用，便于后续更新
                key = (weekday, slot_id)
                self.course_labels[key] = course_label
        
        # 创建保存按钮
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(20, 0))
        
        def save_changes():
            # 保存课程表数据到JSON文件
            try:
                json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'schedule.json')
                
                # 保存数据
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(self.schedule_data, f, ensure_ascii=False, indent=2)

                # 保存后刷新主界面的课表显示
                self.load_schedule()
                self.display_todays_schedule()
                self.update_class_status()
                
                self.show_windows_notification("搞定！", "课程表已保存")
                schedule_window.destroy()
            except Exception as e:
                self.show_windows_notification("保存失败", str(e))
        
        save_button = ttk.Button(button_frame, text="保存", command=save_changes, width=15)
        save_button.pack(side="right", padx=10)
        
        cancel_button = ttk.Button(button_frame, text="取消", command=schedule_window.destroy, width=15)
        cancel_button.pack(side="right", padx=10)
    
    def on_course_click(self, event, weekday, slot_id):
        """处理课程单元格点击事件，插入输入框进行编辑"""
        # 获取点击的组件
        widget = event.widget
        
        # 如果点击的是标签，找到它的父框架
        if isinstance(widget, ttk.Label):
            frame = widget.master
        else:
            frame = widget
        
        # 检查是否已经有输入框在编辑
        for child in frame.winfo_children():
            if isinstance(child, ttk.Entry):
                return
        
        # 获取当前课程名称
        current_course = ""
        if weekday in self.schedule_data['school_days']:
            for course in self.schedule_data['school_days'][weekday]:
                if course['slot_id'] == slot_id:
                    current_course = course['name']
                    break
        
        # 清空框架
        for child in frame.winfo_children():
            child.destroy()
        
        # 创建输入框
        entry_var = tk.StringVar(value=current_course)
        entry = ttk.Entry(frame, textvariable=entry_var, font=('Microsoft YaHei UI', 10), width=10, justify="center")
        entry.pack(fill="both", expand=True)
        entry.focus_set()
        
        # 绑定回车键保存
        def on_enter_press(event):
            new_course = entry_var.get().strip()
            
            # 更新数据
            if weekday not in self.schedule_data['school_days']:
                self.schedule_data['school_days'][weekday] = []
            
            # 检查是否已存在该时间段的课程
            course_exists = False
            for i, course in enumerate(self.schedule_data['school_days'][weekday]):
                if course['slot_id'] == slot_id:
                    if new_course:
                        # 更新课程
                        self.schedule_data['school_days'][weekday][i]['name'] = new_course
                    else:
                        # 删除课程
                        del self.schedule_data['school_days'][weekday][i]
                    course_exists = True
                    break
            
            # 如果不存在且有新课程，则添加
            if not course_exists and new_course:
                self.schedule_data['school_days'][weekday].append({
                    'slot_id': slot_id,
                    'name': new_course
                })
            
            # 清空框架并显示新的课程标签
            for child in frame.winfo_children():
                child.destroy()
            
            course_label = ttk.Label(frame, text=new_course, font=('Microsoft YaHei UI', 10), width=10, anchor="center", wraplength=100)
            course_label.pack(fill="both", expand=True)
            
            # 为新创建的标签绑定点击事件
            def on_new_label_click(event):
                self.on_course_click(event, weekday, slot_id)
            
            course_label.bind("<Button-1>", on_new_label_click)
            
            # 更新标签引用
            key = (weekday, slot_id)
            self.course_labels[key] = course_label
        
        entry.bind("<Return>", on_enter_press)
        
        # 绑定失去焦点事件
        def on_focus_out(event):
            # 优先处理点击其他地方的事件
            frame.after(100, lambda: on_enter_press(None))
        
        entry.bind("<FocusOut>", on_focus_out)

    # 创建NotepadWindow实例
    def show_notepad_window(self):
        """创建桌面便签窗口"""
        # 增加便签计数
        self.notepad_count += 1
        
        # 创建便签窗口实例
        notepad = self.NotepadWindow(self, self.notepad_count)
        
        # 将实例添加到列表中管理
        self.notepads.append(notepad)

    class NotepadWindow:
        def __init__(self, parent, count):
            self.parent = parent
            self.count = count
            self.window = None
            self.operation_frame = None
            self.content_frame = None
            self.title_label = None
            self.title_entry = None
            self.close_button = None
            self.notepad_label = None
            self.notepad_entry = None
            self.notepad_text = None
            self.notepad_title = None
            self.drag_data = None
            self.font_size = 12  # 字体大小变量
            self.window_width = 300
            self.window_height = 300
            self.operation_height = 30
            self.notepad_setting_height = 30
            self.content_height = self.window_height - self.operation_height - self.notepad_setting_height - 50
            
            self.create_window()

        def create_window(self):
            # 创建便签窗口
            self.window = tk.Toplevel(self.parent.root)
            self.window.overrideredirect(True)  # 无边框
            self.window.configure(bg="#fff599")  # 黄色背景
            self.window.title(f"便签#{self.count}")  # 设置窗口标题
            self.window.attributes("-alpha", 0.9)

            
            # 设置窗口大小和位置
            screen_width = self.window.winfo_screenwidth()
            screen_height = self.window.winfo_screenheight()
            x = (screen_width - self.window_width) // 2 + (self.count - 1) * 50
            y = (screen_height - self.window_height) // 2 + (self.count - 1) * 50
            self.window.geometry(f"{self.window_width}x{self.window_height}+{x}+{y}")
            
            # 将窗口置于最底层
            try:
                hwnd = win32gui.FindWindow(None, "")
                win32gui.SetWindowPos(hwnd, win32con.HWND_BOTTOM, 0, 0, 0, 0, 
                                    win32con.SWP_NOSIZE | win32con.SWP_NOMOVE | win32con.SWP_NOACTIVATE)
            except:
                pass  # 如果没有安装pywin32库，就不执行此操作
            
            # 窗口布局：使用grid布局确保各区域位置固定
            self.window.grid_rowconfigure(0, minsize=self.operation_height, weight=0)
            self.window.grid_rowconfigure(1, minsize=self.content_height, weight=1)
            self.window.grid_rowconfigure(2, minsize=self.notepad_setting_height, weight=0)
            self.window.grid_columnconfigure(0, weight=1)
            
            # 创建操作区框架（标题栏）
            self.operation_frame = tk.Frame(self.window, bg="#fff599", height=self.operation_height)
            self.operation_frame.grid(row=0, column=0, sticky="ew")
            self.operation_frame.pack_propagate(False)  # 防止框架大小被内容改变
            
            # 操作区布局
            self.operation_frame.grid_columnconfigure(0, weight=1)
            self.operation_frame.grid_columnconfigure(1, minsize=20, weight=0)
            
            # 标题标签
            self.notepad_title = tk.StringVar(value=f"便签#{self.count}")
            self.title_label = tk.Label(self.operation_frame, textvariable=self.notepad_title, 
                                    font=('Microsoft YaHei UI', 15, 'bold'), bg="#fff599",
                                    anchor='w', cursor='hand2')
            self.title_label.grid(row=0, column=0, sticky="ew", padx=10, pady=(5,0))
            
            # 拖拽按钮
            self.drag_button = tk.Label(self.operation_frame, text="≡", 
                                    font=('Microsoft YaHei UI', 15), bg="#fff599",
                                    cursor='hand2')
            self.drag_button.grid(row=0, column=1, sticky="e", padx=5, pady=0)

            # 关闭按钮
            self.close_button = tk.Label(self.operation_frame, text="✕", 
                                    font=('Microsoft YaHei UI', 15), bg="#fff599",
                                    fg="red", cursor='hand2')
            self.close_button.grid(row=0, column=2, sticky="e", padx=5, pady=0)

            # 创建内容区框架
            self.content_frame = tk.Frame(self.window, bg="#fff599")
            self.content_frame.grid(row=1, column=0, sticky="nsew", padx=0, pady=0)
            self.content_frame.pack_propagate(False)

            # 配置内容区的网格布局，为滚动条留出空间
            self.content_frame.grid_rowconfigure(0, weight=1)
            self.content_frame.grid_columnconfigure(0, weight=1)
            self.content_frame.grid_columnconfigure(1, minsize=15, weight=0)

            # 创建共享滚动条
            self.scrollbar = tk.Scrollbar(self.content_frame, 
                                          orient=tk.VERTICAL, 
                                          troughcolor="#fff599",
                                          takefocus=False
                                          )
            self.scrollbar.grid(row=0, column=1, sticky="ns")

            # 创建便签内容标签（放入Canvas实现滚动）
            self.notepad_text = tk.StringVar(value="点击输入文本")

            # 创建Canvas用于包装Label实现滚动
            self.notepad_canvas = tk.Canvas(self.content_frame, bg="#fff599", bd=0, highlightthickness=0,
                                            yscrollcommand=self.scrollbar.set)
            self.notepad_canvas.grid(row=0, column=0, sticky="nsew")

            # 创建Label并放入Canvas
            self.notepad_label = tk.Label(self.notepad_canvas, textvariable=self.notepad_text, 
                                        font=('Microsoft YaHei UI', self.font_size), bg='#fff599', 
                                        wraplength=self.window_width-10, justify='left', 
                                        anchor='nw')

            # 将Label添加到Canvas
            self.label_window = self.notepad_canvas.create_window((12, 12), window=self.notepad_label, 
                                                                anchor="nw", tags="label")

            # 设置Canvas滚动区域
            self.notepad_text.trace("w", self.update_canvas_scrollregion)

            # 绑定滚动条到Canvas
            self.scrollbar.config(command=self.notepad_canvas.yview)
            
            # 创建设置框架
            self.notepad_setting_frame = tk.Frame(self.window, bg="#fff599", height=self.notepad_setting_height)
            self.notepad_setting_frame.grid(row=2, column=0, sticky="ew")
            self.notepad_setting_frame.pack_propagate(False)  # 防止框架大小被内容改变
            
            # 字号调整区布局
            self.notepad_setting_frame.grid_columnconfigure(0, minsize=50, weight=0)
            self.notepad_setting_frame.grid_columnconfigure(1, minsize=60, weight=0)
            self.notepad_setting_frame.grid_columnconfigure(2, weight=1)
            
            # 添加字号标签
            ttk.Label(self.notepad_setting_frame, text="字号:", font=('Microsoft YaHei UI', 10), 
                    background="#fff599").grid(row=0, column=0, sticky="w", padx=5, pady=5)
            
            # 添加ttk步进器
            self.font_size_var = tk.IntVar(value=self.font_size)
            self.font_size_spinbox = ttk.Spinbox(self.notepad_setting_frame, from_=8, to=50, 
                                            textvariable=self.font_size_var, 
                                            command=self.on_font_size_change,
                                            width=5)
            self.font_size_spinbox.grid(row=0, column=1, sticky="w", padx=5, pady=5)

            # 设置ttk复选框样式，使其背景与便签背景一致
            style = ttk.Style()
            style.configure("Notepad.TCheckbutton", background="#fff599", fieldbackground="#fff599")

            # "始终置顶"复选框
            self.always_on_top_var = tk.BooleanVar(value=False)
            self.always_on_top_checkbox = ttk.Checkbutton(self.notepad_setting_frame, text="始终置顶",
                                                        variable=self.always_on_top_var,
                                                        command=self.on_always_on_top_change,
                                                        style="Notepad.TCheckbutton",
                                                        takefocus=False)
            self.always_on_top_checkbox.grid(row=0, column=2, sticky="e", padx=5, pady=5)
            
            # 绑定事件
            self.title_label.bind("<Button-1>", lambda e: self.on_title_click(e))  # 标题点击事件
            self.close_button.bind("<Button-1>", lambda e: self.on_close_click(e))  # 关闭按钮点击事件
            self.notepad_label.bind("<Button-1>", lambda e: self.on_notepad_click(e))  # 内容点击事件
            
            # 绑定拖动事件
            self.drag_button.bind("<Button-1>", lambda e: self.on_drag_start(e))
            self.drag_button.bind("<B1-Motion>", lambda e: self.on_drag_motion(e))
            
            # 发送通知
            self.parent.show_windows_notification(f"便签#{self.count} 已创建", "点击便签可以添加内容\n点击标题可以修改标题\n按住 ≡ 可以移动便签位置")

        def update_canvas_scrollregion(self, *args):
            """更新Canvas的滚动区域"""
            self.window.after_idle(lambda: self.notepad_canvas.configure(
                scrollregion=self.notepad_canvas.bbox("all")
            ))

        def update_label_wraplength(self):
            """更新Label的wraplength，确保不被滚动条遮挡"""
            # 计算可用宽度：窗口宽度 - 左右内边距(12*2) - 滚动条宽度(如果显示的话)
            # 由于滚动条宽度是15px(minsize=15)，我们统一减去这个宽度以确保内容不被遮挡
            available_width = self.window_width - 24 - 15  # 24=12*2内边距，15=滚动条宽度
            self.notepad_label.configure(wraplength=available_width)
            self.update_canvas_scrollregion()
        
        def on_always_on_top_change(self):
            """处理始终置顶复选框状态变化"""
            if self.always_on_top_var.get():
                self.window.attributes("-topmost", True)
            else:
                self.window.attributes("-topmost", False)

        def on_font_size_change(self):
            """处理字号调整事件"""
            self.font_size = self.font_size_var.get()
            # 更新标签的字体大小
            self.notepad_label.configure(font=('Microsoft YaHei UI', self.font_size))
            # 更新Label的wraplength
            self.update_label_wraplength()
            # 更新Canvas滚动区域
            self.update_canvas_scrollregion()
            # 如果当前显示的是输入框，也更新输入框的字体大小
            if hasattr(self, 'notepad_entry') and self.notepad_entry:
                self.notepad_entry.configure(font=('Microsoft YaHei UI', self.font_size))

        def on_title_click(self, event):
            """处理标题点击事件，将标题变为输入框"""
            # 如果已有输入框，先销毁
            if hasattr(self, 'title_entry') and self.title_entry:
                self.title_entry.destroy()
                self.title_entry = None
            
            # 获取当前标题
            current_title = self.notepad_title.get()
            
            # 创建临时输入框替换标题标签
            self.title_entry = tk.Entry(self.operation_frame, font=('Microsoft YaHei UI', 15, 'bold'),
                                    bg="#fff599", bd=0)
            self.title_entry.grid(row=0, column=0, sticky="ew", padx=10, pady=(5,5))
            
            # 设置输入框内容
            self.title_entry.insert(0, current_title)
            self.title_entry.select_range(0, tk.END)
            
            # 隐藏标题标签
            self.title_label.grid_remove()
            
            # 绑定事件
            self.title_entry.bind("<FocusOut>", lambda e: self.on_title_save(e))
            self.title_entry.bind("<Return>", lambda e: self.on_title_save(e))
            # 防止点击输入框时触发父容器的点击事件
            self.title_entry.bind("<Button-1>", lambda e: "break")
            
            # 获取焦点
            self.title_entry.focus_set()

        def on_title_save(self, event):
            """保存标题并切换回标签"""
            # 获取新标题
            new_title = self.title_entry.get().strip()
            if not new_title:
                new_title = f"便签#{self.count}"
            
            # 更新标题
            self.notepad_title.set(new_title)
            
            # 销毁输入框
            self.title_entry.destroy()
            self.title_entry = None
            
            # 显示标题标签
            self.title_label.grid(row=0, column=0, sticky="ew", padx=10, pady=(5,0))

        def on_close_click(self, event):
            """处理关闭按钮点击事件，显示确认提示框"""
            # 安全获取当前标题，避免AttributeError
            if hasattr(self, 'title_entry') and self.title_entry:
                title = self.title_entry.get().strip()
            else:
                title = self.notepad_title.get()
            
            if messagebox.askyesno("确认关闭", f"确定要关闭“{title}”吗？\n注意：关闭后内容不会保存"):
                # 确认关闭，销毁便签窗口
                self.window.destroy()
                # 从主应用的便签列表中移除
                if self in self.parent.notepads:
                    self.parent.notepads.remove(self)

        def on_drag_start(self, event):
            """开始拖动窗口"""
            self.drag_data = {"x": event.x_root - self.window.winfo_x(),
                            "y": event.y_root - self.window.winfo_y()}

        def on_drag_motion(self, event):
            """拖动窗口"""
            x = event.x_root - self.drag_data["x"]
            y = event.y_root - self.drag_data["y"]
            self.window.geometry(f"+{x}+{y}")

        def on_notepad_click(self, event):
            """处理便签点击事件，将标签变为输入框"""
            # 如果已有输入框，先销毁
            if hasattr(self, 'notepad_entry') and self.notepad_entry:
                self.notepad_entry.destroy()
                self.notepad_entry = None
            
            # 获取当前便签内容
            current_text = self.notepad_text.get()
            if current_text == "点击输入文本":
                current_text = ""
            
            # 隐藏Canvas（包含Label）
            self.notepad_canvas.grid_remove()
            
            # 创建多行输入框
            self.notepad_entry = tk.Text(self.content_frame, font=('Microsoft YaHei UI', self.font_size), 
                                        bg='#fff599', wrap=tk.WORD, bd=0, 
                                        height=5, yscrollcommand=self.scrollbar.set)
            self.notepad_entry.grid(row=0, column=0, sticky="nsew", padx=12, pady=12)
            
            # 设置滚动条绑定到Text
            self.scrollbar.config(command=self.notepad_entry.yview)
            
            # 设置输入框内容
            self.notepad_entry.insert(tk.END, current_text)
            self.notepad_entry.see(tk.END)  # 滚动到文本末尾
            
            # 绑定事件
            self.notepad_entry.bind("<FocusOut>", lambda e: self.on_notepad_save(e))  # 失去焦点保存
            self.notepad_entry.bind("<KeyRelease>", lambda e: self.on_notepad_text_changed(e))  # 文本变化时保存
            # 防止点击输入框时触发父容器的点击事件
            self.notepad_entry.bind("<Button-1>", lambda e: "break")
            
            # 使用after方法延迟获取焦点，确保输入框已完全创建
            self.window.after(50, lambda: self.notepad_entry.focus_set())

        def on_notepad_text_changed(self, event):
            """处理便签文本变化事件，实时保存"""
            self.save_notepad_content()

        def on_notepad_save(self, event):
            """保存便签内容并切换回标签"""
            self.save_notepad_content()
            
            # 彻底销毁输入框实例
            self.notepad_entry.destroy()
            self.notepad_entry = None
            
            # 显示Canvas（包含Label）
            self.notepad_canvas.grid(row=0, column=0, sticky="nsew")
            
            # 设置滚动条绑定到Canvas
            self.scrollbar.config(command=self.notepad_canvas.yview)

            # 更新Label的wraplength（确保内容不被滚动条遮挡）
            self.update_label_wraplength()
            
            # 更新Canvas滚动区域
            self.update_canvas_scrollregion()

        def save_notepad_content(self):
            """保存便签内容"""
            # 获取输入框内容
            if hasattr(self, 'notepad_entry') and self.notepad_entry:
                content = self.notepad_entry.get("1.0", tk.END).strip()
                
                # 如果内容为空，显示默认提示文字
                if not content:
                    content = "点击输入文本"
                
                # 更新标签内容
                self.notepad_text.set(content)

    def show_context_menu(self, event):
        """显示右键菜单"""
        if event:
            self.context_menu.post(event.x_root, event.y_root)
        else:
            # 获取按钮在屏幕上的位置
            x = self.right_click_menu_button.winfo_rootx()
            y = self.right_click_menu_button.winfo_rooty() + self.right_click_menu_button.winfo_height()
            self.context_menu.post(x, y)
        
    def run(self):
        """运行程序"""
        self.root.mainloop()

if __name__ == "__main__":
    # 创建一个唯一的命名互斥体
    mutex_name = "DesktopClockCountdownMutex"
    mutex = win32event.CreateMutex(None, True, mutex_name)
    
    # 检查是否已存在该互斥体（即程序是否已在运行）
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        # 程序已经在运行，显示提示并退出
        root = tk.Tk()
        root.withdraw()  # 隐藏主窗口
        root.destroy()
        sys.exit(0)
    
    # 首次运行正常启动
    app = CountdownApp()
    try:
        app.run()
    finally:
        # 程序退出，释放互斥体
        win32api.CloseHandle(mutex)