# Caps Lock 检测器

一个智能的Caps Lock状态检测和自动切换工具，专为特定软件设计。

## 功能特性

- 🎯 **智能检测**: 自动检测当前活动窗口
- 🔄 **自动切换**: 在特定软件中自动切换Caps Lock状态
- ⚡ **实时监控**: 100ms间隔实时检测窗口切换
- 🎨 **美观界面**: 简洁的GUI状态显示
- 📊 **日志记录**: 完整的操作日志记录

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 直接运行
```bash
python caps_lock_checker.py
```

### 打包为EXE
```bash
python -m PyInstaller --noconsole --onefile --icon caps_lock_checker.ico caps_lock_checker.py
```

## 配置文件

编辑 `config.txt` 文件来配置需要监控的软件：

```
[software_list]
CAXA
SolidWorks
AutoCAD
```

## 项目结构

```
├── caps_lock_checker.py    # 主程序
├── config.txt              # 配置文件
├── caps_lock_checker.ico   # 应用程序图标
├── draw_ico_v2.py          # ICO生成脚本
├── requirements.txt         # 依赖文件
├── setup.py                # 安装脚本
└── logs/                   # 日志目录
```

## 技术栈

- Python 3.11+
- Tkinter (GUI界面)
- pywin32 (Windows API调用)
- PyInstaller (打包工具)
- Pillow (图标生成)

## 开发说明

### 代码架构

1. **CapsLockChecker类**: 主应用程序类
2. **GUI界面**: 使用Tkinter构建的状态显示窗口
3. **窗口检测**: 使用win32gui检测当前活动窗口
4. **状态管理**: 跟踪Caps Lock状态和窗口切换

### 核心逻辑

- 每100ms检测一次当前活动窗口
- 只在窗口切换时改变Caps Lock状态
- 避免在用户手动操作时干扰

## 许可证

MIT License