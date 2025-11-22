import tkinter as tk
from tkinter import ttk, Menu
import win32api
import win32con
import win32gui
import time
import logging
import os
import datetime

class CapsLockChecker:
    def __init__(self, root):
        self.root = root
        self.root.title("Caps Lock 状态检测")
        
        # 初始化日志功能
        self.setup_logging()
        self.logger.info("应用程序启动")
        
        self.logger.debug("初始化前 - 检查self属性")
        
        # 设置默认颜色值
        self.color_caps_on = "#fa6666"
        self.color_caps_off = "#4CAF50"
        self.color_titlebar = "#2c3e50"
        
        self.root.resizable(True, True)
        self.root.overrideredirect(True)  # 永久无边框窗口，避免overrideredirect切换导致的卡顿
        
        # 标题栏状态和拖动控制
        self.titlebar_visible = False
        self.dragging = False
        self.drag_offset = (0, 0)
        self.leave_hide_timer = None  # 鼠标离开后延迟隐藏的计时器
        self.last_mouse_move_time = 0  # 记录上次鼠标移动时间
        
        self.logger.debug("创建标题栏前 - 检查self属性")
        # 创建自定义标题栏
        self.titlebar = tk.Frame(self.root, bg=self.color_titlebar, height=30)
        self.titlebar.pack_propagate(False)  # 防止标题栏高度被内部组件改变
        self.logger.debug("创建标题栏后 - self.titlebar = %s", self.titlebar)
        
        # 添加关闭按钮
        self.close_button = tk.Label(
            self.titlebar,
            text="×",
            font= ("Arial", 12, "bold"),
            fg="white",
            bg=self.color_titlebar,
            cursor="hand2"
        )
        self.close_button.pack(side=tk.RIGHT, padx=5, pady=2)
        self.close_button.bind("<Button-1>", self.on_close_click)
        
        # 添加刷新按钮
        self.refresh_button = tk.Label(
            self.titlebar,
            text="⟳",
            font= ("Arial", 10, "bold"),
            fg="white",
            bg=self.color_titlebar,
            cursor="hand2"
        )
        self.refresh_button.pack(side=tk.RIGHT, padx=5, pady=2)
        self.refresh_button.bind("<Button-1>", lambda e: self.refresh_config())
        
        # 添加设置按钮
        self.settings_button = tk.Label(
            self.titlebar,
            text="⚙",
            font= ("Arial", 10, "bold"),
            fg="white",
            bg=self.color_titlebar,
            cursor="hand2"
        )
        self.settings_button.pack(side=tk.RIGHT, padx=5, pady=2)
        self.settings_button.bind("<Button-1>", lambda e: self.show_settings_window())
        
        # 标题栏拖动功能
        self.titlebar.bind("<ButtonPress-1>", self.on_titlebar_drag_start)
        self.titlebar.bind("<B1-Motion>", self.on_titlebar_drag_motion)
        
        # 绑定窗口拖动事件（标题栏隐藏时可拖动整个窗口）
        self.root.bind("<ButtonPress-1>", self.on_window_drag_start)
        self.root.bind("<B1-Motion>", self.on_window_drag_motion)
        self.root.bind("<ButtonRelease-1>", self.on_drag_stop)
        
        # 绑定鼠标事件
        self.root.bind("<Motion>", self.on_mouse_motion)
        self.root.bind("<Enter>", self.on_mouse_enter)
        self.root.bind("<Leave>", self.on_mouse_leave)
        
        # 绑定ESC键关闭窗口
        self.root.bind("<Escape>", self.on_escape)
        
        # 创建右键菜单
        self.right_click_menu = Menu(self.root, tearoff=False)
        self.right_click_menu.add_command(label="设置", command=self.show_settings_window)
        self.right_click_menu.add_command(label="刷新", command=self.refresh_config)
        self.right_click_menu.add_separator()
        self.right_click_menu.add_command(label="关闭", command=self.on_menu_close)
        self.root.bind("<Button-3>", self.show_right_click_menu)
        
        # 初始化Caps Lock状态
        self.caps_lock_on = win32api.GetKeyState(win32con.VK_CAPITAL) & 1 != 0
        self.last_hwnd = None
        
        self.logger.debug("创建主框架前")
        # 创建主框架，填充整个窗口
        self.main_frame = tk.Frame(self.root, bg=self.color_caps_on if self.caps_lock_on else self.color_caps_off)
        self.main_frame.place(x=0, y=0, relwidth=1, relheight=1)
        
        # 创建容器帧放置状态文本
        self.label_container = tk.Frame(self.main_frame, bg=self.main_frame['bg'])
        self.label_container.pack(expand=True, fill=tk.BOTH)

        # 创建内部容器用于居中显示
        self.inner_frame = tk.Frame(self.label_container, bg=self.main_frame['bg'])
        self.inner_frame.pack(expand=True)

        # 添加固定文本标签（"Caps Lock"）
        self.caps_text_label = tk.Label(
            self.inner_frame,
            text="Caps Lock",
            font=("Arial", 22, "bold"),
            fg="white",
            bg=self.main_frame['bg']
        )
        self.caps_text_label.pack(side=tk.LEFT, padx=(0, 5))  # 固定文本右侧留5px间距

        # 添加动态状态标签（"ON"/"OFF"）
        self.status_value_label = tk.Label(
            self.inner_frame,
            text="ON" if self.caps_lock_on else "OFF",
            font=("Arial", 22, "bold"),
            fg="white",
            bg=self.main_frame['bg'],
        )
        self.status_value_label.pack(side=tk.LEFT)  # 动态状态标签紧随其后

        # 设置状态标签的固定宽度，确保ON/OFF切换时Caps Lock文本不动
        self.status_value_label.configure(width=4)  # 设置固定宽度为4个字符，足够容纳"OFF"

        self.logger.debug("读取配置文件前")
        # 读取配置文件
        self.read_config()
        
        # 开始定期检查Caps Lock状态
        self.check_caps_lock()
        
        self.logger.debug("应用配置前 - 检查self.titlebar属性")
        if hasattr(self, 'titlebar'):
            self.logger.debug("应用配置前 - self.titlebar存在: %s", self.titlebar)
        else:
            self.logger.error("应用配置前 - self.titlebar不存在!")
        
        # 应用配置
        self.apply_config()
        
        # 隐藏标题栏
        self.hide_titlebar()
    
    def update_status(self):
        """更新界面显示的Caps Lock状态"""
        if self.caps_lock_on:
            status_text = "ON"
        else:
            status_text = "OFF"

        # 更新状态文本
        self.status_value_label.configure(text=status_text)
    
    def check_caps_lock(self):
        """检测Caps Lock状态变化并根据配置的软件列表自动切换"""
        hwnd = win32gui.GetForegroundWindow()
        window_title = win32gui.GetWindowText(hwnd)
        is_target_software_active = any(software in window_title for software in self.config['software_list'])
        if not self.config['software_list']:
            self.config['software_list'] = ['CAXA']
        current_status = win32api.GetKeyState(win32con.VK_CAPITAL) & 1 != 0
        is_switch = hwnd != self.last_hwnd
        if is_switch:
            desired_status = is_target_software_active
            if current_status != desired_status:
                win32api.keybd_event(win32con.VK_CAPITAL, 0, win32con.KEYEVENTF_EXTENDEDKEY | 0, 0)
                win32api.keybd_event(win32con.VK_CAPITAL, 0, win32con.KEYEVENTF_EXTENDEDKEY | win32con.KEYEVENTF_KEYUP, 0)
                self.caps_lock_on = desired_status
            else:
                self.caps_lock_on = current_status
            self.last_hwnd = hwnd
        else:
            self.caps_lock_on = current_status
        self.update_main_frame_bg()
        self.update_status()
        self.root.after(100, self.check_caps_lock)
    
    def hide_titlebar(self):
        """隐藏自定义标题栏"""
        if self.titlebar_visible:
            self.titlebar.place_forget()
            self.titlebar_visible = False
            self.update_main_frame_bg()  # 更新背景色
    
    def show_titlebar(self):
        """显示自定义标题栏"""
        if not self.titlebar_visible:
            self.titlebar.place(x=0, y=0, relwidth=1, height=30)
            self.titlebar.lift()  # 确保标题栏在主框架上方
            self.titlebar_visible = True
            self.update_main_frame_bg()  # 更新背景色
    
    def on_titlebar_drag_start(self, event):
        """开始拖动自定义标题栏"""
        self.dragging = True
        self.drag_offset = (event.x_root - self.root.winfo_x(), event.y_root - self.root.winfo_y())
    
    def on_titlebar_drag_motion(self, event):
        """拖动自定义标题栏"""
        if self.dragging:
            new_x = event.x_root - self.drag_offset[0]
            new_y = event.y_root - self.drag_offset[1]
            self.root.geometry(f"+{new_x}+{new_y}")
    
    def on_window_drag_start(self, event):
        """标题栏隐藏时拖动整个窗口"""
        if not self.titlebar_visible:
            self.dragging = True
            self.drag_offset = (event.x_root - self.root.winfo_x(), event.y_root - self.root.winfo_y())
    
    def on_window_drag_motion(self, event):
        """拖动整个窗口"""
        if self.dragging and not self.titlebar_visible:
            new_x = event.x_root - self.drag_offset[0]
            new_y = event.y_root - self.drag_offset[1]
            self.root.geometry(f"+{new_x}+{new_y}")
    
    def on_drag_stop(self, event):
        """停止拖动窗口"""
        self.dragging = False
    
    def on_mouse_motion(self, event):
        """鼠标移动事件"""
        current_time = time.time()
        
        # 添加时间戳检查，限制标题栏显示的频率
        if current_time - self.last_mouse_move_time > 0.05:  # 50ms间隔
            self.last_mouse_move_time = current_time
            
            # 鼠标靠近顶部时显示标题栏
            if not self.titlebar_visible and event.y < 30:
                self.show_titlebar()
        
        # 取消延迟隐藏计时器
        if self.leave_hide_timer:
            self.root.after_cancel(self.leave_hide_timer)
            self.leave_hide_timer = None
    
    def on_mouse_enter(self, event):
        """鼠标进入窗口时显示标题栏"""
        if not self.dragging:
            self.show_titlebar()
    
    def on_mouse_leave(self, event):
        """鼠标离开窗口时隐藏标题栏"""
        if self.titlebar_visible and not self.dragging:
            # 延迟隐藏标题栏，确保用户有足够时间操作
            self.leave_hide_timer = self.root.after(500, self.hide_titlebar)
    
    def close_application(self):
        """关闭应用程序"""
        self.save_window_position()
        self.logger.info("应用程序退出")
        self.root.destroy()
        
    def on_close_click(self, event):
        """关闭按钮点击事件"""
        self.close_application()
    
    def update_main_frame_bg(self):
        """更新主框架背景色"""
        current_color = self.color_caps_on if self.caps_lock_on else self.color_caps_off
        self.main_frame.configure(bg=current_color)
        self.label_container.configure(bg=current_color)
        self.inner_frame.configure(bg=current_color)
        self.caps_text_label.configure(bg=current_color)
        self.status_value_label.configure(bg=current_color)
    
    def center_window(self):
        """将窗口居中显示在屏幕上"""
        # 强制更新窗口尺寸信息
        self.root.update_idletasks()
        
        # 获取屏幕尺寸
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # 获取窗口尺寸
        window_width = self.root.winfo_width()
        window_height = self.root.winfo_height()
        
        # 计算中心位置
        center_x = int((screen_width - window_width) / 2)
        center_y = int((screen_height - window_height) / 2)
        
        # 设置窗口位置
        self.root.geometry(f"+{center_x}+{center_y}")
    
    def on_escape(self, event):
        """ESC键关闭窗口"""
        self.close_application()
    
    def on_menu_close(self):
        """右键菜单关闭事件"""
        self.close_application()
    
    def show_right_click_menu(self, event):
        """显示右键菜单"""
        self.right_click_menu.post(event.x_root, event.y_root)
    
    def show_settings_window(self):
        """显示设置窗口"""
        try:
            if os.path.exists('config.txt'):
                os.startfile('config.txt')
                self.logger.info("配置文件已打开")
            else:
                # 如果文件不存在，创建文件后打开
                with open('config.txt', 'w', encoding='utf-8') as f:
                    f.write('')
                os.startfile('config.txt')
                self.logger.info("配置文件已创建并打开")
        except Exception as e:
            self.logger.error(f"打开配置文件失败: {e}")
    
    def apply_config(self):
        try:
            # 设置窗口大小和位置
            width = self.config.get("window_width", 250)
            height = self.config.get("window_height", 150)
            x = self.config.get("window_x", -1)
            y = self.config.get("window_y", -1)
            
            # 如果位置是有效坐标，则合并大小和位置
            if x != -1 and y != -1:
                self.root.geometry(f"{width}x{height}+{x}+{y}")
            else:
                # 设置大小并居中显示
                self.root.geometry(f"{width}x{height}")
                self.center_window()
            
            # 应用颜色设置
            self.color_caps_on = self.config.get("color_caps_on", "#fa6666")
            self.color_caps_off = self.config.get("color_caps_off", "#4CAF50")
            self.color_titlebar = self.config.get("color_titlebar", "#2c3e50")
            
            # 更新标题栏颜色
            self.titlebar.configure(bg=self.color_titlebar)
            self.close_button.configure(bg=self.color_titlebar)
            
            # 更新主框架颜色
            self.update_main_frame_bg()
            
            # 设置窗口总是在最前
            self.root.wm_attributes("-topmost", self.config.get("always_on_top", 0))
        except Exception as e:
            self.logger.error(f"应用配置时出错: {str(e)}", exc_info=0)
    
    def read_config(self):
        """读取配置文件"""
        # 默认配置
        self.config = {
            'color_caps_on': '#fa6666',
            'color_caps_off': '#4CAF50',
            'color_titlebar': '#2c3e50',
            'window_width': 250,
            'window_height': 150,
            'window_x': -1,
            'window_y': -1,
            'always_on_top': 0,
            'software_list': ['CAXA']  # 默认检测软件列表
        }
        
        try:
            if os.path.exists('config.txt'):
                with open('config.txt', 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    
                # 使用字典映射处理配置键值，提高效率
                config_handlers = {
                    'color_caps_on': str,
                    'color_caps_off': str,
                    'color_titlebar': str,
                    'window_width': int,
                    'window_height': int,
                    'window_x': int,
                    'window_y': int,
                    'always_on_top': lambda v: v.strip().lower() in ['true', '1', 'yes', 'on'],
                    'software_list': lambda v: [sw.strip() for sw in v.split(',')]  # 解析软件列表
                }
                
                for line in lines:
                    line = line.strip()
                    if not line or line.startswith('#'):
                        continue
                    
                    if '=' in line:
                        key, value = line.split('=', 1)
                        key = key.strip()
                        value = value.strip()
                        
                        # 处理注释
                        if ';' in value:
                            value = value.split(';')[0].strip()
                        
                        if key in config_handlers:
                            self.config[key] = config_handlers[key](value)
                            
                self.logger.info("配置文件读取成功")
                self.logger.debug(f"配置内容: {self.config}")
            else:
                # 配置文件不存在，生成默认配置文件
                self.write_config_file()
                self.logger.info("配置文件不存在，已生成默认配置")
        except Exception as e:
            self.logger.error(f"读取配置文件失败: {str(e)}", exc_info=True)
    
    def write_config_file(self):
        """将配置写入文件"""
        try:
            with open('config.txt', 'w', encoding='utf-8') as f:
                f.write('# Caps Lock 检测器配置文件\n')
                f.write('# 颜色设置\n')
                f.write(f"color_caps_on = {self.config['color_caps_on']}\n")
                f.write(f"color_caps_off = {self.config['color_caps_off']}\n")
                f.write(f"color_titlebar = {self.config['color_titlebar']}\n")
                f.write('\n# 窗口设置\n')
                f.write(f"window_width = {self.config['window_width']}\n")
                f.write(f"window_height = {self.config['window_height']}\n")
                f.write(f"window_x = {self.config['window_x']}\n")
                f.write(f"window_y = {self.config['window_y']}\n")
                f.write('\n# 其他设置\n')
                f.write(f"always_on_top = {'true' if self.config['always_on_top'] else 'false'}\n")
                f.write(f"software_list = {','.join(self.config['software_list'])}\n")  # 写入软件列表
            self.logger.info("默认配置文件已生成")
        except Exception as e:
            self.logger.error(f"写入配置文件失败: {str(e)}")

    def refresh_config(self):
        """刷新配置文件"""
        try:
            self.logger.info("刷新配置文件")
            self.read_config()
            self.apply_config()
            self.logger.info("配置文件刷新成功")
        except Exception as e:
            self.logger.error(f"刷新配置文件失败: {str(e)}", exc_info=True)
    
    def save_window_position(self):
        """保存当前窗口位置和大小到配置文件"""
        try:
            # 获取当前窗口位置和大小
            geo = self.root.geometry()
            if 'x' in geo and '+' in geo:
                # 解析宽度、高度和位置
                size_part, pos_part = geo.split('+', 1)
                width, height = map(int, size_part.split('x'))
                x, y = map(int, pos_part.split('+', 1))
                
                # 读取最新配置
                self.read_config()
                
                # 读取现有配置内容
                existing_lines = []
                if os.path.exists('config.txt'):
                    with open('config.txt', 'r', encoding='utf-8') as f:
                        existing_lines = [line.strip() for line in f if line.strip()]
                
                # 构建新的配置内容
                new_lines = []
                
                # 创建配置项映射
                config_map = {
                    'color_caps_on': self.config.get('color_caps_on', '#fa6666'),
                    'color_caps_off': self.config.get('color_caps_off', '#4CAF50'),
                    'color_titlebar': self.config.get('color_titlebar', '#2c3e50'),
                    'window_width': str(width),
                    'window_height': str(height),
                    'window_x': str(x),
                    'window_y': str(y),
                    'always_on_top': 'true' if self.config.get('always_on_top', False) else 'false'
                }
                
                # 处理现有配置项
                for line in existing_lines:
                    if not line or line.startswith('#'):
                        new_lines.append(line)
                        continue
                    
                    if '=' in line:
                        key, value = line.split('=', 1)
                        key = key.strip()
                        
                        # 更新需要保存的配置项
                        if key in config_map:
                            # 保留原有注释
                            if ';' in value:
                                comment = value.split(';', 1)[1]
                                new_lines.append(f'{key} = {config_map[key]};{comment}')
                            else:
                                new_lines.append(f'{key} = {config_map[key]}')
                            
                            # 从映射中移除已处理的项
                            del config_map[key]
                        else:
                            # 保留未定义的配置项
                            new_lines.append(line)
                
                # 添加剩余的配置项
                for key, value in config_map.items():
                    new_lines.append(f'{key} = {value}')
                
                # 写入配置文件
                with open('config.txt', 'w', encoding='utf-8') as f:
                    for line in new_lines:
                        f.write(f'{line}\n')
                
                self.logger.info(f"窗口位置已保存: {geo}")
        except Exception as e:
            self.logger.error(f"保存窗口位置失败: {str(e)}", exc_info=True)
    
    def setup_logging(self):
        """设置日志系统"""
        # 创建logs文件夹
        log_folder = 'logs'
        if not os.path.exists(log_folder):
            os.makedirs(log_folder)
            self.logger.info("日志文件夹创建成功")
        
        # 生成日志文件名
        today = datetime.date.today()
        log_filename = os.path.join(log_folder, 'log_{}.txt'.format(today.strftime('%Y-%m-%d')))
        
        # 配置日志
        logging.basicConfig(
            filename=log_filename,
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s',
            encoding='utf-8'
        )
        self.logger = logging.getLogger(__name__)

if __name__ == "__main__":
    root = tk.Tk()
    app = CapsLockChecker(root)
    root.mainloop()