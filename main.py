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
import win32gui
import re

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
            return resp.read().decode('utf-8').strip()
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

def get_active_document_path_via_window():
    """通过窗口标题获取WPS当前活动文档的路径（如果文档已保存）"""
    def enum_windows_callback(hwnd, result):
        if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
            title = win32gui.GetWindowText(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            # WPS文字窗口类名通常包含 "WPS" 或 "KWPS"，标题包含扩展名
            if ('WPS' in class_name or 'KWPS' in class_name) and any(ext in title for ext in ['.docx', '.doc', '.wps', '.docm']):
                # 提取文件名，标题格式可能是 "文档名 - WPS文字" 或 "[文档名]"
                parts = re.split(r' - | \[|\]', title)
                if parts:
                    doc_name = parts[0].strip()
                    # 在常见目录中搜索
                    for search_dir in [os.path.expanduser("~\\Documents"), os.path.expanduser("~\\Desktop")]:
                        for ext in ['.docx', '.doc', '.wps', '.docm']:
                            full_path = os.path.join(search_dir, doc_name + ext)
                            if os.path.exists(full_path):
                                result.append(full_path)
                                return False
                    # 如果标题可能包含完整路径
                    if os.path.exists(doc_name):
                        result.append(doc_name)
                        return False
        return True

    windows = []
    win32gui.EnumWindows(enum_windows_callback, windows)
    if windows:
        return windows[0]
    return None

def get_active_document_path():
    """获取当前活动文档路径（先COM，后窗口）"""
    # 尝试通过 COM 获取 Word 活动文档
    for progid in ["Word.Application", "Kwps.Application"]:
        try:
            app = win32.GetActiveObject(progid)
            if app.Documents.Count > 0:
                doc = app.ActiveDocument
                if doc and doc.FullName:
                    return doc.FullName
        except:
            pass
    # 尝试通过窗口标题获取（适用于WPS）
    return get_active_document_path_via_window()

class WordToolApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Word 格式处理工具")
        self.root.geometry("560x600")
        self.root.resizable(True, True)

        self.btn_frame = tk.Frame(self.root)
        self.btn_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        self.log = scrolledtext.ScrolledText(self.root, height=15, width=70)
        self.log.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        script_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(script_dir)

        self.plugins = load_plugins()
        self.create_buttons()
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

    def check_for_updates(self):
        def task():
            local_ver = get_local_version()
            remote_ver = check_remote_version()
            if remote_ver and remote_ver != local_ver:
                self.root.after(0, lambda: self.log_msg(f"✨ 发现新版本 {remote_ver}（当前 {local_ver}），请点击「获取最新版本」更新。"))
                self.root.after(0, lambda: self.update_btn.config(bg="#ffa500", text="📥 有新版本！"))
        threading.Thread(target=task, daemon=True).start()

    def run_plugin(self, plugin_func):
        def task():
            pythoncom.CoInitialize()
            word_app = None
            doc = None
            doc_path = None
            need_close_doc = False
            try:
                # 获取文档路径
                doc_path = get_active_document_path()
                if not doc_path:
                    self.log_msg("❌ 无法自动检测到打开的文档（请确保文档已保存），请手动打开Word或使用Microsoft Word。")
                    return
                self.log_msg(f"📄 正在处理文档：{os.path.basename(doc_path)}")
                # 打开文档（使用 Word 对象模型，兼容 WPS）
                word_app = win32.gencache.EnsureDispatch("Word.Application")
                # 设置为不可见，避免闪烁
                word_app.Visible = False
                doc = word_app.Documents.Open(doc_path)
                need_close_doc = True
                # 执行插件
                success, msg = plugin_func(doc)
                self.log_msg(f"{'✅' if success else '❌'} {msg}")
            except Exception as e:
                self.log_msg(f"❌ 运行出错: {e}")
                import traceback
                traceback.print_exc()
            finally:
                if need_close_doc and doc:
                    try:
                        doc.Close(SaveChanges=0)  # 不保存
                    except:
                        pass
                if word_app:
                    try:
                        word_app.Quit()
                    except:
                        pass
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
