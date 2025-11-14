#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
運行 Streamlit 儀表板
"""

import subprocess
import sys
import os

print("=" * 60)
print("啟動 Streamlit 儀表板")
print("=" * 60)
print()

# 確保在正確的目錄
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

print(f"工作目錄: {os.getcwd()}")
print(f"儀表板文件: visualization_dashboard.py")
print()

# 檢查文件是否存在
if not os.path.exists('visualization_dashboard.py'):
    print("錯誤: 找不到 visualization_dashboard.py")
    sys.exit(1)

# 檢查必要文件
required_files = ['彙總表.xlsx']
missing_files = [f for f in required_files if not os.path.exists(f)]

if missing_files:
    print("警告: 以下文件不存在:")
    for f in missing_files:
        print(f"  - {f}")
    print()
    print("請先運行 execute_prompt.py 生成必要文件")
    print()

print("正在啟動 Streamlit...")
print("瀏覽器將自動打開，如果沒有，請訪問: http://localhost:8501")
print()
print("按 Ctrl+C 停止服務器")
print("=" * 60)
print()

try:
    # 運行 streamlit
    subprocess.run([sys.executable, "-m", "streamlit", "run", "visualization_dashboard.py"])
except KeyboardInterrupt:
    print("\n\n服務器已停止")
except Exception as e:
    print(f"\n錯誤: {e}")
    print("\n請確保已安裝 streamlit:")
    print("  pip install streamlit")





