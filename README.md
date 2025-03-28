# Shortcut Generator Tool / 快捷方式生成工具

## Overview / 概述

**Shortcut Generator Tool** is a Python-based application that allows you to create shortcuts with custom icons for local and remote files. The tool can generate `.lnk` files that you can use to quickly access files or applications on your desktop or other locations. It also supports adding execution parameters for more advanced use cases.  
**快捷方式生成工具**是一个基于Python的应用程序，允许您为本地和远程文件创建带有自定义图标的快捷方式。该工具可以生成`.lnk`文件，您可以用它们快速访问桌面或其他位置的文件或应用程序。它还支持添加执行参数以满足更高级的使用需求。

## Features / 功能

- **Create Shortcuts**: Create shortcuts for local files or remote links.  
  **创建快捷方式**：为本地文件或远程链接生成快捷方式。
- **Custom Icon Support**: Choose a default icon or select a custom icon from your local machine.  
  **自定义图标**：选择默认图标或从本地计算机中选择自定义图标。
- **Execution Parameters**: Add execution parameters to your shortcut (e.g., command line arguments).  
  **执行参数**：为快捷方式设置执行参数（例如命令行参数）。
- **Folder Management**: Organize the shortcut and associated files into a hidden folder.  
  **文件夹管理**：将快捷方式和关联文件组织到隐藏的文件夹中。
- **Remote Link Support**: Supports creating shortcuts for remote links, allowing you to directly enter a URL.  
  **远程链接加载**：支持为远程链接创建快捷方式，可以直接输入URL进行链接。
- **Cross-platform Compatibility**: The tool is designed for Windows and uses `win32com.client` to create shortcuts.  
  **跨平台兼容性**：本工具设计用于Windows，利用`win32com.client`来创建快捷方式。

## Requirements / 环境要求

- **Python 3.x**  
- **Libraries**:  
  - `tkinter`: For the graphical user interface (GUI).  
  - `win32com.client`: For creating Windows shortcuts.  
  - `shutil`: For file operations.  
  - `os`: For path and directory management.  
  - **安装依赖**：确保您的Windows机器上已安装`pywin32`：
```bash
pip install pywin32

