# -*- coding: utf-8 -*-
import os
import sys
import importlib.util
import tkinter as tk
from tkinter import scrolledtext, messagebox
import threading
import subprocess
import pythoncom
import win32com.client as win32

def load_plugins(plugins_dir="plugins"):
    """扫描 plugins 目录，加载所有插件"""
    plugins = []
    if not os.path.exists(plugins_dir):
        os.makedirs(plugins_dir)
    for fname in os.listdir(plugins_dir):
        if fname.endswith(".py") and fname != "__init__.py":
            mod_name = fname[:-3]
            try:
                spec = importlib.util.spec_from_file_location(mod_name, os.path.join(plugins_dir, fname))
                mod = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)
                if hasattr(mod, "NAME") and hasattr(mod, "run"):
                    plugins.append((mod.NAME, mod.run))
                else:
                    print(f"⚠️ 插件 {fname} 缺少 NAME 或 run 函数")
            except Exception as e:
                print(f"❌ 加载插件 {fname} 失败: {e}")
    return plugins

class WordToolApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Word 格式处理工具")
        self.root.geometry("560x520")
        self.root.resizable(True, True)

        # 按钮区域
        self.btn_frame = tk.Frame(self.root)
        self.btn_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        # 日志区域
        self.log = scrolledtext.ScrolledText(self.root, height=15, width=70)
        self.log.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # 加载插件并生成按钮
        self.plugins = load_plugins()
        self.create_buttons()

        self.root.mainloop()

    def create_buttons(self):
        # 清除旧按钮（如果有）
        for widget in self.btn_frame.winfo_children():
            widget.destroy()

        # 1. 功能按钮（每个插件一个）
        if not self.plugins:
            lbl = tk.Label(self.btn_frame, text="未发现任何插件，请将插件放入 plugins 文件夹", fg="red")
            lbl.pack(pady=20)
        else:
            for name, func in self.plugins:
                btn = tk.Button(self.btn_frame, text=name, command=lambda f=func: self.run_plugin(f),
                                width=30, pady=2)
                btn.pack(pady=3)

        # 2. 分隔线
        tk.Label(self.btn_frame, text="").pack()

        # 3. 更新按钮（只有点击此按钮才会调用 updater.py）
        update_btn = tk.Button(self.btn_frame, text="📥 获取最新版本", command=self.run_update,
                               width=30, pady=2, bg="#d9ead3")
        update_btn.pack(pady=5)

    def run_plugin(self, plugin_func):
        """在新线程中运行插件，避免界面卡死"""
        def task():
            pythoncom.CoInitialize()
            word = None
            try:
                word = win32.gencache.EnsureDispatch('Word.Application')
                doc = word.ActiveDocument
                if doc is None:
                    self.log_msg("❌ 请先在 Word 中打开一个文档！")
                    return
                success, msg = plugin_func(doc)
                self.log_msg(f"{'✅' if success else '❌'} {msg}")
            except Exception as e:
                self.log_msg(f"❌ 运行出错: {e}")
            finally:
                if word:
                    word.Quit()
                pythoncom.CoUninitialize()
        threading.Thread(target=task).start()

    def run_update(self):
        """执行更新操作（仅当用户点击按钮时）"""
        if not messagebox.askyesno("更新确认",
                                   "是否从 GitHub 获取最新版本？\n主程序将关闭并自动更新，更新后会自动重启。"):
            return

        self.log_msg("📡 正在启动更新程序...")
        updater_path = os.path.join(os.path.dirname(__file__), "updater.py")
        if not os.path.exists(updater_path):
            self.log_msg("❌ 错误：找不到 updater.py 文件，请确保它与 main.py 在同一目录。")
            return

        try:
            # 启动独立更新进程，不等待
            subprocess.Popen([sys.executable, updater_path])
            # 关闭当前主程序窗口
            self.root.quit()
        except Exception as e:
            self.log_msg(f"❌ 启动更新失败: {e}")

    def log_msg(self, msg):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.root.update()

if __name__ == "__main__":
    WordToolApp()
