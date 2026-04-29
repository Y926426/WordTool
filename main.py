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
import urllib.request

# ---------- 版本检测 ----------
REMOTE_VERSION_URL = "https://raw.githubusercontent.com/Y926426/WordTool/main/version.txt"
LOCAL_VERSION_FILE = "version.txt"

def get_local_version():
    if os.path.exists(LOCAL_VERSION_FILE):
        with open(LOCAL_VERSION_FILE, 'r', encoding='utf-8') as f:
            return f.read().strip()
    return "0.0.0"

def check_remote_version():
    try:
        with urllib.request.urlopen(REMOTE_VERSION_URL, timeout=5) as resp:
            remote_version = resp.read().decode('utf-8').strip()
            return remote_version
    except:
        return None

def load_plugins(plugins_dir="plugins"):
    plugins = []
    if not os.path.exists(plugins_dir):
        os.makedirs(plugins_dir)
        return plugins
    for fname in os.listdir(plugins_dir):
        if fname.endswith(".py") and fname != "__init__.py":
            mod_name = fname[:-3]
            filepath = os.path.join(plugins_dir, fname)
            try:
                spec = importlib.util.spec_from_file_location(mod_name, filepath)
                mod = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(mod)
                name = getattr(mod, "NAME", None)
                run_func = getattr(mod, "run", None)
                if name and callable(run_func):
                    plugins.append((name, run_func))
                else:
                    print(f"⚠️ 插件 {fname} 缺少 NAME 或 run 函数")
            except Exception as e:
                print(f"❌ 加载插件 {fname} 失败: {e}")
    return plugins

class WordToolApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Word 格式处理工具")
        self.root.geometry("560x600")
        self.root.resizable(True, True)

        # 状态栏（显示当前文档，可选）
        self.status_frame = tk.Frame(self.root)
        self.status_frame.pack(fill=tk.X, padx=10, pady=5)
        self.current_doc_label = tk.Label(self.status_frame, text="当前文档：未检测", fg="blue", anchor="w")
        self.current_doc_label.pack(side=tk.LEFT)

        # 按钮区域
        self.btn_frame = tk.Frame(self.root)
        self.btn_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        # 日志区域
        self.log = scrolledtext.ScrolledText(self.root, height=15, width=70)
        self.log.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        script_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(script_dir)

        self.plugins = load_plugins()
        self.create_buttons()
        self.refresh_current_doc()   # 定时刷新文档名
        self.check_for_updates()     # 后台检查更新
        self.root.mainloop()

    def create_buttons(self):
        for widget in self.btn_frame.winfo_children():
            widget.destroy()
        if not self.plugins:
            lbl = tk.Label(self.btn_frame, text="未发现任何插件，请将插件放入 plugins 文件夹", fg="red")
            lbl.pack(pady=20)
        else:
            for name, func in self.plugins:
                btn = tk.Button(self.btn_frame, text=name, command=lambda f=func: self.run_plugin(f),
                                width=30, pady=2)
                btn.pack(pady=3)
        tk.Label(self.btn_frame, text="").pack()

        # 更新按钮（可能高亮）
        self.update_btn = tk.Button(self.btn_frame, text="📥 获取最新版本", command=self.run_update,
                                    width=30, pady=2, bg="#d9ead3")
        self.update_btn.pack(pady=5)

    def refresh_current_doc(self):
        """安全获取当前活动文档名称（支持Word和WPS）"""
        try:
            word = win32.gencache.EnsureDispatch('Word.Application')
            if word.Documents.Count == 0:
                self.current_doc_label.config(text="当前文档：未打开任何文档", fg="orange")
            else:
                doc = word.ActiveDocument
                if doc:
                    self.current_doc_label.config(text=f"当前文档：{doc.Name}", fg="green")
                else:
                    self.current_doc_label.config(text="当前文档：未知", fg="red")
        except Exception as e:
            # 可能WPS没有注册，忽略
            self.current_doc_label.config(text="当前文档：无法连接", fg="red")
        finally:
            if self.root.winfo_exists():
                self.root.after(1000, self.refresh_current_doc)

    def check_for_updates(self):
        """后台检查远程版本，如有更新则提示并高亮更新按钮"""
        def task():
            local_ver = get_local_version()
            remote_ver = check_remote_version()
            if remote_ver and remote_ver != local_ver:
                def update_ui():
                    self.log_msg(f"✨ 发现新版本 {remote_ver}（当前 {local_ver}），请点击「获取最新版本」更新。")
                    self.update_btn.config(bg="#ffa500", text="📥 有新版本！")
                self.root.after(0, update_ui)
        threading.Thread(target=task, daemon=True).start()

    def run_plugin(self, plugin_func):
        def task():
            pythoncom.CoInitialize()
            word = None
            try:
                word = win32.gencache.EnsureDispatch('Word.Application')
                doc = word.ActiveDocument
                if doc is None:
                    self.log_msg("❌ 请先在 Word 中打开一个文档！")
                    return
                self.log_msg(f"📄 正在处理文档：{doc.Name}")
                success, msg = plugin_func(doc)
                self.log_msg(f"{'✅' if success else '❌'} {msg}")
            except Exception as e:
                self.log_msg(f"❌ 运行出错: {e}")
            finally:
                if word:
                    doc = None
                    word = None
                pythoncom.CoUninitialize()
        threading.Thread(target=task).start()

    def run_update(self):
        if not messagebox.askyesno("更新确认",
                                   "是否从 GitHub 获取最新版本？\n主程序将关闭并自动更新，更新后会自动重启。"):
            return
        self.log_msg("📡 正在启动更新程序...")
        updater_path = os.path.join(os.path.dirname(__file__), "updater.py")
        if not os.path.exists(updater_path):
            self.log_msg("❌ 错误：找不到 updater.py 文件，请确保它与 main.py 在同一目录。")
            return
        try:
            subprocess.Popen([sys.executable, updater_path])
            self.root.quit()
        except Exception as e:
            self.log_msg(f"❌ 启动更新失败: {e}")

    def log_msg(self, msg):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.root.update()

if __name__ == "__main__":
    WordToolApp()
