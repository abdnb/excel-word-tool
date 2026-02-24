#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Linux 打包脚本 - 将程序打包成 Linux 可执行文件
使用方法：在 Linux 环境下运行 python build_linux.py
"""

import PyInstaller.__main__
import os
import sys

# 检查是否在 Linux 环境
if sys.platform != 'linux':
    print("警告：当前不是 Linux 环境！")
    print(f"当前系统：{sys.platform}")
    print("建议在 Linux 环境下运行此脚本以生成 Linux 可执行文件")
    response = input("是否继续？(y/n): ")
    if response.lower() != 'y':
        sys.exit(0)

print("=" * 60)
print("开始打包 Linux 版本...")
print("=" * 60)

# 打包配置
params = [
    'gui_app.py',                    # 主程序文件
    '--name=Excel-Word工具',          # 可执行文件名
    '--onefile',                     # 打包成单个文件
    '--windowed',                    # 不显示控制台窗口（GUI 应用）
    '--clean',                       # 清理临时文件
    '--noconfirm',                   # 不询问确认
    # Linux 下不使用 .ico 图标，可以使用 .png
    # '--icon=icon.png',             # 如果有 PNG 图标可以取消注释
]

# 运行打包
PyInstaller.__main__.run(params)

print("\n" + "=" * 60)
print("打包完成！")
print("=" * 60)
print(f"可执行文件位置: dist/Excel-Word工具")
print("\n使用说明：")
print("1. 将 dist/Excel-Word工具 复制到目标 Linux 系统")
print("2. 添加执行权限：chmod +x Excel-Word工具")
print("3. 双击运行或命令行执行：./Excel-Word工具")
print("\n注意事项：")
print("- 确保目标 Linux 系统已安装图形界面（X11 或 Wayland）")
print("- 如果无法运行，可能需要安装依赖：sudo apt install python3-tk")
print("=" * 60)
