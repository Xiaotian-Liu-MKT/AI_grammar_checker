#!/usr/bin/env python3
"""
AI语法检查器启动脚本
自动检查依赖并启动Streamlit应用
"""

import subprocess
import sys
import os
from pathlib import Path

def check_and_install_requirements():
    """检查并安装依赖"""
    requirements_file = Path(__file__).parent / "requirements.txt"
    
    if not requirements_file.exists():
        print("❌ requirements.txt 文件不存在")
        return False
    
    try:
        print("🔍 检查依赖...")
        # 尝试导入主要依赖
        import streamlit
        import docx
        import pandas
        import litellm
        import openpyxl
        print("✅ 所有依赖已安装")
        return True
        
    except ImportError as e:
        print(f"📦 正在安装依赖: {e}")
        try:
            subprocess.check_call([
                sys.executable, "-m", "pip", "install", "-r", str(requirements_file)
            ])
            print("✅ 依赖安装完成")
            return True
        except subprocess.CalledProcessError:
            print("❌ 依赖安装失败，请手动运行:")
            print(f"pip install -r {requirements_file}")
            return False

def launch_streamlit():
    """启动Streamlit应用"""
    app_file = Path(__file__).parent / "app.py"
    
    if not app_file.exists():
        print("❌ app.py 文件不存在")
        return False
    
    try:
        print("🚀 启动AI语法检查器...")
        print("📱 应用将在浏览器中打开")
        print("🔗 如果没有自动打开，请访问: http://localhost:8501")
        print("\n按 Ctrl+C 停止应用\n")
        
        subprocess.run([
            sys.executable, "-m", "streamlit", "run", str(app_file),
            "--server.port", "8501",
            "--server.headless", "false",
            "--browser.gatherUsageStats", "false"
        ])
        
    except KeyboardInterrupt:
        print("\n👋 应用已停止")
    except Exception as e:
        print(f"❌ 启动失败: {e}")
        return False
    
    return True

def main():
    print("=" * 50)
    print("📝 AI语法检查器")
    print("=" * 50)
    
    # 检查Python版本
    if sys.version_info < (3, 8):
        print("❌ 需要Python 3.8或更高版本")
        return
    
    # 检查并安装依赖
    if not check_and_install_requirements():
        return
    
    # 启动应用
    launch_streamlit()

if __name__ == "__main__":
    main()
