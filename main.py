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
import win32gui
import time

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
        self.root.geometry("560x650")
        self.root.resizable(True, True)

        # 状态栏：显示当前文档
        self.status_frame = tk.Frame(self.root)
        self.status_frame.pack(fill=tk.X, padx=10, pady=5)
        self.doc_status_label = tk.Label(self.status_frame, text="当前文档：未检测", fg="blue", anchor="w")
        self.doc_status_label.pack(side=tk.LEFT)
        self.refresh_btn = tk.Button(self.status_frame, text="刷新", command=self.refresh_doc_status)
        self.refresh_btn.pack(side=tk.RIGHT)

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

        # 启动定时刷新
        self.refresh_doc_status()
        self.root.after(2000, self.auto_refresh)  # 每2秒自动刷新

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
        update_btn = tk.Button(self.btn_frame, text="📥 获取最新版本", command=self.run_update,
                               width=30, pady=2, bg="#d9ead3")
        update_btn.pack(pady=5)

    def get_active_document_info(self):
        """返回当前活动文档的名称和路径（如果可能）"""
        # 先尝试COM方式（Word）
        try:
            word = win32.GetActiveObject("Word.Application")
            if word.Documents.Count > 0:
                doc = word.ActiveDocument
                if doc:
                    return doc.Name, doc.FullName
        except:
            pass
        # 尝试WPS窗口标题
        def enum_callback(hwnd, extra):
            if win32gui.IsWindowVisible(hwnd) and win32gui.IsWindowEnabled(hwnd):
                title = win32gui.GetWindowText(hwnd)
                class_name = win32gui.GetClassName(hwnd)
                if class_name.startswith("KWPS") and any(ext in title for ext in ['.docx', '.doc', '.wps', '.docm']):
                    # 标题格式 "文档名 - WPS文字"
                    parts = title.split(' - ')
                    if parts:
                        doc_name = parts[0].strip()
                        extra.append((doc_name, None))
                        return False
            return True
        windows = []
        win32gui.EnumWindows(enum_callback, windows)
        if windows:
            return windows[0][0], None
        return None, None

    def refresh_doc_status(self):
        """刷新状态栏显示"""
        name, path = self.get_active_document_info()
        if name:
            self.doc_status_label.config(text=f"当前文档：{name}", fg="green")
        else:
            self.doc_status_label.config(text="当前文档：未检测到Word/WPS文档", fg="orange")
        self.root.update()

    def auto_refresh(self):
        """定时自动刷新"""
        self.refresh_doc_status()
        self.root.after(2000, self.auto_refresh)

    def get_word_app_and_doc(self):
        """获取应用程序和文档对象（用于执行插件）"""
        # 先尝试COM（Word）
        try:
            word = win32.GetActiveObject("Word.Application")
            if word.Documents.Count > 0:
                doc = word.ActiveDocument
                if doc:
                    return word, doc
        except:
            pass
        # 如果COM失败，尝试通过窗口标题获取文档路径，然后新打开
        name, path = self.get_active_document_info()
        if path and os.path.exists(path):
            try:
                word = win32.gencache.EnsureDispatch("Word.Application")
                word.Visible = False
                doc = word.Documents.Open(path)
                return word, doc
            except Exception as e:
                print(f"打开文档失败: {e}")
        return None, None

    def run_plugin(self, plugin_func):
        def task():
            pythoncom.CoInitialize()
            app = None
            doc = None
            need_close = False
            try:
                app, doc = self.get_word_app_and_doc()
                if doc is None:
                    self.log_msg("❌ 无法获取打开的文档，请确保Word/WPS已打开一个文档（已保存）")
                    return
                self.log_msg(f"📄 正在处理文档：{doc.Name}")
                success, msg = plugin_func(doc)
                self.log_msg(f"{'✅' if success else '❌'} {msg}")
            except Exception as e:
                self.log_msg(f"❌ 运行出错: {e}")
                import traceback
                traceback.print_exc()
            finally:
                if need_close and doc:
                    try:
                        doc.Close(SaveChanges=0)
                    except:
                        pass
                if app:
                    # 如果是新打开的Word实例且没有其他文档，退出
                    try:
                        if app.Documents.Count == 0:
                            app.Quit()
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