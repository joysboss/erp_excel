"""
Smart Excel Extractor - 服务器启动脚本
快速启动Web服务进行Excel数据提取
"""
import subprocess
import sys
import os
import webbrowser
import time
import threading


def start_server():
    """启动FastAPI服务器"""
    try:
        # 使用uvicorn启动服务
        import uvicorn
        from main import app
        
        print("=" * 60)
        print("🚀 Smart Excel Extractor 启动中...")
        print("=" * 60)
        print("📍 服务地址: http://localhost:8000")
        print("📖 API文档: http://localhost:8000/docs")
        print("📊 Web界面: http://localhost:8000")
        print("=" * 60)
        
        uvicorn.run(app, host="127.0.0.1", port=8000, reload=False)
        
    except ImportError:
        print("❌ 未安装uvicorn，请运行: pip install uvicorn")
        sys.exit(1)
    except Exception as e:
        print(f"❌ 服务启动失败: {e}")
        sys.exit(1)


def main():
    """主函数"""
    print("🌟 Smart Excel Extractor - Excel数据智能提取系统")
    print("💡 基于列结构映射，无需AI，100%准确率")
    print()
    
    choice = input("是否立即启动Web服务? (y/n): ").strip().lower()
    
    if choice in ['y', 'yes', '是', '']:
        # 在新线程中启动浏览器
        def open_browser():
            time.sleep(3)  # 等待服务器启动
            webbrowser.open("http://localhost:8000")
        
        browser_thread = threading.Thread(target=open_browser)
        browser_thread.daemon = True
        browser_thread.start()
        
        # 启动服务器
        start_server()
    else:
        print()
        print("📚 使用方法:")
        print("1. 直接运行: python main.py")
        print("2. 或使用: uvicorn main:app --host 127.0.0.1 --port 8000")
        print("3. 访问: http://localhost:8000")


if __name__ == "__main__":
    main()