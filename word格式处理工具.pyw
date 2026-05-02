# -*- coding: utf-8 -*-
import os
import sys
import importlib.util
import tkinter as tk
from tkinter import scrolledtext, messagebox, filedialog
import threading
import subprocess
import pythoncom
import win32com.client as win32
import urllib.request
import win32gui
import re

# ---------- 版本检测 ----------
REMOTE_VERSION_URL = "https://raw.githubusercontent.com/Y926426/WordTool/main/version.txt"
LOCAL_VERSION_FILE = "version.txt"

def get_local_version():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    version_file = os.path.join(script_dir, "version.txt")
    try:
        if os.path.exists(version_file):
            with open(version_file, 'r', encoding='utf-8') as f:
                content = f.read().strip()
                if content:
                    return content
        return "0.0.0"
    except:
        return "0.0.0"

def check_remote_version():
    try:
        with urllib.request.urlopen(REMOTE_VERSION_URL, timeout=5) as resp:
            return resp.read().decode('utf-8').strip()
    except:
        return None

def load_plugins(plugins_dir="plugins"):
    plugins = []
    script_dir = os.path.dirname(os.path.abspath(__file__))
    full_plugins_dir = os.path.join(script_dir, plugins_dir)
    if not os.path.exists(full_plugins_dir):
        os.makedirs(full_plugins_dir)
        return plugins
    for fname in os.listdir(full_plugins_dir):
        if fname.endswith(".py") and fname != "__init__.py":
            mod_name = fname[:-3]
            filepath = os.path.join(full_plugins_dir, fname)
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

def get_wps_window_title():
    def enum_callback(hwnd, titles):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            class_name = win32gui.GetClassName(hwnd)
            title = win32gui.GetWindowText(hwnd)
            if ('WPS' in class_name or 'Kwps' in class_name) and title:
                titles.append(title)
        return True
    titles = []
    win32gui.EnumWindows(enum_callback, titles)
    if titles:
        return titles[0]
    return None

def get_document_name_from_title(title):
    if not title:
        return "未知"
    parts = re.split(r' - | \[|\]', title)
    return parts[0].strip()

def get_word_app_and_doc():
    for progid in ["Word.Application", "Kwps.Application"]:
        try:
            app = win32.GetActiveObject(progid)
            if app.Documents.Count > 0:
                doc = app.ActiveDocument
                if doc:
                    return app, doc
        except:
            pass
    return None, None

def get_document_path_via_file_dialog():
    file_path = filedialog.askopenfilename(
        title="请选择要处理的 Word 文档",
        filetypes=[("Word文档", "*.docx *.doc *.wps"), ("所有文件", "*.*")]
    )
    return file_path if file_path else None

class WordToolApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Word 格式处理工具")
        self.root.geometry("560x620")
        self.root.resizable(True, True)

        self.local_version = get_local_version()

        self.status_frame = tk.Frame(self.root)
        self.status_frame.pack(fill=tk.X, padx=10, pady=5)

        self.current_doc_label = tk.Label(self.status_frame, text="当前文档：未检测", fg="blue", anchor="w")
        self.current_doc_label.pack(side=tk.LEFT)

        self.version_label = tk.Label(self.status_frame, text=f"v{self.local_version}", fg="gray", anchor="e")
        self.version_label.pack(side=tk.RIGHT)

        self.btn_frame = tk.Frame(self.root)
        self.btn_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        self.log = scrolledtext.ScrolledText(self.root, height=15, width=70)
        self.log.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        script_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(script_dir)

        self.plugins = load_plugins()
        self.create_buttons()
        self.refresh_doc_status()
        self.check_for_updates()
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
        self.update_btn = tk.Button(self.btn_frame, text="📥 获取最新版本", command=self.run_update,
                                    width=30, pady=2, bg="#d9ead3")
        self.update_btn.pack(pady=5)

    def refresh_doc_status(self):
        title = get_wps_window_title()
        if title:
            doc_name = get_document_name_from_title(title)
            self.current_doc_label.config(text=f"当前文档：{doc_name}", fg="green")
        else:
            _, doc = get_word_app_and_doc()
            if doc:
                self.current_doc_label.config(text=f"当前文档：{doc.Name}", fg="green")
            else:
                self.current_doc_label.config(text="当前文档：未检测到打开的文档", fg="orange")
        if self.root.winfo_exists():
            self.root.after(2000, self.refresh_doc_status)

    def check_for_updates(self):
        def task():
            local_ver = self.local_version
            remote_ver = check_remote_version()
            if remote_ver and remote_ver != local_ver:
                self.root.after(0, lambda: self.log_msg(f"✨ 发现新版本 {remote_ver}（当前 {local_ver}），请点击「获取最新版本」更新。"))
                self.root.after(0, lambda: self.update_btn.config(bg="#ffa500", text="📥 有新版本！"))
        threading.Thread(target=task, daemon=True).start()

    def get_active_document_for_processing(self):
        app, doc = get_word_app_and_doc()
        if doc:
            return app, doc
        self.log_msg("⚠️ 无法自动获取已打开的文档，请手动选择一个文件（将复制一份临时处理，不影响原文件）")
        file_path = get_document_path_via_file_dialog()
        if not file_path:
            return None, None
        try:
            new_app = win32.gencache.EnsureDispatch("Word.Application")
            new_app.Visible = False
            new_doc = new_app.Documents.Open(file_path)
            return new_app, new_doc
        except Exception as e:
            self.log_msg(f"❌ 打开文件失败: {e}")
            return None, None

    def run_plugin(self, plugin_func, plugin_name):
        """执行插件，支持无文档插件（白名单）"""
        # 定义不需要打开文档的插件名称列表
        no_doc_plugins = ["报告匹配助手"]  # 可按需增加其他插件名
        if plugin_name in no_doc_plugins:
            # 此类插件无需文档，直接调用，传入 None 或任意占位
            try:
                success, msg = plugin_func(None)  # 插件内部应忽略 doc 参数
                self.log_msg(f"{'✅' if success else '❌'} {msg}")
            except Exception as e:
                self.log_msg(f"❌ 运行出错: {e}")
            return

        # 以下为需要文档的插件原有逻辑
        def task():
            pythoncom.CoInitialize()
            app = None
            doc = None
            need_cleanup = False
            try:
                app, doc = self.get_active_document_for_processing()
                if doc is None:
                    self.log_msg("❌ 没有可处理的文档，请先打开一个文档或手动选择文件。")
                    return
                self.log_msg(f"📄 正在处理文档：{doc.Name}")
                success, msg = plugin_func(doc)
                self.log_msg(f"{'✅' if success else '❌'} {msg}")
                if need_cleanup and doc:
                    try:
                        doc.Close(SaveChanges=0)
                    except:
                        pass
            except Exception as e:
                self.log_msg(f"❌ 运行出错: {e}")
                import traceback
                traceback.print_exc()
            finally:
                if need_cleanup and app:
                    try:
                        app.Quit()
                    except:
                        pass
                pythoncom.CoUninitialize()
        threading.Thread(target=task).start()

    def run_update(self):
        if not messagebox.askyesno("更新确认",
                                   "是否从 GitHub 获取最新版本？\n主程序将关闭并自动更新。"):
            return
        self.log_msg("📡 正在启动更新程序...")
        updater_path = os.path.join(os.path.dirname(__file__), "updater.py")
        if not os.path.exists(updater_path):
            self.log_msg("❌ 错误：找不到 updater.py 文件，请确保它与 main.py 在同一目录。")
            return
        try:
            python_exe = sys.executable
            if "pythonw.exe" in python_exe.lower():
                base_dir = os.path.dirname(python_exe)
                possible = os.path.join(base_dir, "python.exe")
                if os.path.exists(possible):
                    python_exe = possible
            subprocess.Popen([python_exe, updater_path])
            self.root.quit()
        except Exception as e:
            self.log_msg(f"❌ 启动更新失败: {e}")

    def log_msg(self, msg):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.root.update()

if __name__ == "__main__":
    WordToolApp()
