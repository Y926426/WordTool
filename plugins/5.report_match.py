# -*- coding: utf-8 -*-
import os
import webbrowser

NAME = "报告匹配助手"

def run(doc):
    # 工具根目录（main.py 所在目录）
    tool_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    html_path = os.path.join(tool_dir, "report_match.html")
    
    if os.path.exists(html_path):
        # 使用默认浏览器打开
        webbrowser.open(f"file:///{os.path.abspath(html_path)}")
        return True, "已打开报告匹配助手（浏览器）"
    else:
        return False, f"未找到 report_match.html，请将文件放在 {tool_dir} 目录下"
