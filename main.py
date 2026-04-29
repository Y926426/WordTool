import os
import sys
import tkinter as tk
from tkinter import scrolledtext, messagebox
import threading
import importlib.util

def diagnose_and_load_plugins(plugins_dir="plugins"):
    print("=== 诊断信息 ===")
    print("当前工作目录:", os.getcwd())
    print("脚本所在目录:", os.path.dirname(os.path.abspath(__file__)))
    # 确保 plugins 路径基于脚本所在目录
    base_dir = os.path.dirname(os.path.abspath(__file__))
    plugins_path = os.path.join(base_dir, plugins_dir)
    print("plugins 绝对路径:", plugins_path)
    print("目录是否存在:", os.path.exists(plugins_path))
    if os.path.exists(plugins_path):
        files = os.listdir(plugins_path)
        print("plugins 目录下的文件:", files)
    else:
        print("plugins 目录不存在！")
    
    plugins = []
    if os.path.exists(plugins_path):
        for fname in os.listdir(plugins_path):
            if fname.endswith(".py") and fname != "__init__.py":
                mod_name = fname[:-3]
                full_path = os.path.join(plugins_path, fname)
                print(f"\n尝试加载 {fname}...")
                try:
                    spec = importlib.util.spec_from_file_location(mod_name, full_path)
                    mod = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(mod)
                    if hasattr(mod, "NAME") and hasattr(mod, "run"):
                        plugins.append((mod.NAME, mod.run))
                        print(f"  成功: NAME={mod.NAME}")
                    else:
                        print(f"  失败: 缺少 NAME 或 run 函数")
                        if hasattr(mod, "NAME"):
                            print(f"    NAME存在: {mod.NAME}")
                        else:
                            print(f"    NAME不存在")
                        if hasattr(mod, "run"):
                            print(f"    run存在")
                        else:
                            print(f"    run不存在")
                except Exception as e:
                    print(f"  加载异常: {e}")
    print(f"\n总共加载了 {len(plugins)} 个插件")
    return plugins

class WordToolApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Word 格式处理工具")
        self.root.geometry("500x450")

        self.plugins = diagnose_and_load_plugins()
        
        # 创建按钮框架
        self.btn_frame = tk.Frame(self.root)
        self.btn_frame.pack(pady=10)
        
        if not self.plugins:
            label = tk.Label(self.btn_frame, text="未找到任何插件！请检查 plugins 文件夹。", fg="red")
            label.pack()
        
        for name, func in self.plugins:
            btn = tk.Button(self.btn_frame, text=name, command=lambda f=func: self.run_plugin(f), width=25)
            btn.pack(pady=3)
        
        self.log = scrolledtext.ScrolledText(self.root, height=15, width=65)
        self.log.pack(pady=10)
        
        self.root.mainloop()

    def run_plugin(self, plugin_func):
        def task():
            # COM初始化将在实际调用时由插件run内部处理？你的插件run是直接接收doc的，所以需要在这里初始化Word
            # 简化：直接调用插件run(doc)前先获取doc
            import pythoncom
            import win32com.client as win32
            pythoncom.CoInitialize()
            word = None
            try:
                word = win32.gencache.EnsureDispatch('Word.Application')
                doc = word.ActiveDocument
                if doc is None:
                    self.log_msg("请先在Word中打开一个文档！")
                    return
                success, msg = plugin_func(doc)
                self.log_msg(f"{'✅' if success else '❌'} {msg}")
            except Exception as e:
                self.log_msg(f"❌ 运行出错：{e}")
            finally:
                if word:
                    word.Quit()
                pythoncom.CoUninitialize()
        threading.Thread(target=task).start()

    def log_msg(self, msg):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.root.update()

if __name__ == "__main__":
    WordToolApp()
