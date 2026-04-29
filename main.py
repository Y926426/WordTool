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

        # ========== 新增：当前文档状态栏 ==========
        self.status_frame = tk.Frame(self.root)
        self.status_frame.pack(fill=tk.X, padx=10, pady=5)
        self.current_doc_label = tk.Label(self.status_frame, text="当前文档：未检测", fg="blue", anchor="w")
        self.current_doc_label.pack(side=tk.LEFT)

        # 刷新按钮（手动刷新）
        self.refresh_btn = tk.Button(self.status_frame, text="刷新", command=self.refresh_current_doc)
        self.refresh_btn.pack(side=tk.RIGHT)

        # 按钮区域
        self.btn_frame = tk.Frame(self.root)
        self.btn_frame.pack(pady=10, fill=tk.BOTH, expand=True)

        # 日志区域
        self.log = scrolledtext.ScrolledText(self.root, height=15, width=70)
        self.log.pack(pady=10, padx=10, fill=tk.BOTH, expand=True)

        # 强制将工作目录切换到 main.py 所在目录
        script_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(script_dir)
        print(f"当前工作目录已切换至: {os.getcwd()}")

        self.plugins = load_plugins()
        self.create_buttons()

        # 启动定时刷新当前文档（每秒一次）
        self.refresh_current_doc()
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

    def refresh_current_doc(self):
        """刷新显示当前激活的 Word 文档名称（在主线程中安全调用）"""
        try:
            word = win32.gencache.EnsureDispatch('Word.Application')
            doc = word.ActiveDocument
            if doc:
                self.current_doc_label.config(text=f"当前文档：{doc.Name}", fg="green")
            else:
                self.current_doc_label.config(text="当前文档：未打开任何文档", fg="orange")
            # 注意：不调用 word.Quit()，避免干扰用户正在使用的 Word
        except Exception as e:
            self.current_doc_label.config(text=f"当前文档：获取失败 ({str(e)})", fg="red")
        finally:
            # 每1秒刷新一次（仅当窗口未销毁时）
            if self.root.winfo_exists():
                self.root.after(1000, self.refresh_current_doc)

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
                # 记录当前操作的文档名
                self.log_msg(f"📄 正在处理文档：{doc.Name}")
                success, msg = plugin_func(doc)
                self.log_msg(f"{'✅' if success else '❌'} {msg}")
            except Exception as e:
                self.log_msg(f"❌ 运行出错: {e}")
            finally:
                if word:
                    # 不关闭 Word，仅释放引用
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