#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Smart Excel Extractor 启动器
用于快速启动Excel数据提取服务
"""

import subprocess
import sys
import os
from main import app
import uvicorn
import socket


def check_port(port):
    """检查端口是否可用"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex(('localhost', port)) != 0


def check_dependencies():
    """检查依赖是否安装"""
    try:
        import pandas
        import fastapi
        import uvicorn
        import openpyxl
        import xlrd
        import jinja2
        return True
    except ImportError as e:
        print(f"❌ 缺少依赖: {e}")
        print("请运行: pip install -r requirements.txt")
        return False


def main():
    print("🚀 Smart Excel Extractor 启动器")
    print("="*40)
    
    if not check_dependencies():
        return
    
    # 固定使用8000端口
    selected_port = 8000
    
    # 检查端口是否可用
    if not check_port(selected_port):
        print(f"❌ 端口 {selected_port} 已被占用")
        print("💡 请关闭占用端口的程序后再试，或修改launch.py中的端口号")
        return
    
    print(f"✅ 依赖检查通过")
    print(f"🌐 服务将在 http://localhost:{selected_port} 启动")
    print(f"📌 按 Ctrl+C 可停止服务")
    print()
    
    try:
        print("正在启动服务...")
        uvicorn.run(app, host="127.0.0.1", port=selected_port, reload=False)
    except KeyboardInterrupt:
        print("\n👋 服务已停止")
    except Exception as e:
        print(f"❌ 启动失败: {e}")


if __name__ == "__main__":
    main()