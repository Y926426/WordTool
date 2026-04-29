import os
import sys
import shutil
import urllib.request
import zipfile
import tempfile
import subprocess
import time
import ctypes

REPO_URL = "https://github.com/Y926426/WordTool/archive/refs/heads/main.zip"
CURRENT_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

def download_and_update():
    print("📥 正在从 GitHub 下载最新版本...")
    temp_zip = None
    extract_dir = None
    try:
        temp_zip = os.path.join(tempfile.gettempdir(), "wordtool_latest.zip")
        urllib.request.urlretrieve(REPO_URL, temp_zip)
        print("✅ 下载完成，正在解压...")
        extract_dir = tempfile.mkdtemp()
        with zipfile.ZipFile(temp_zip, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        source_dir = None
        for item in os.listdir(extract_dir):
            full = os.path.join(extract_dir, item)
            if os.path.isdir(full) and item.startswith("WordTool"):
                source_dir = full
                break
        if not source_dir:
            print("❌ 未找到解压后的 WordTool 目录")
            return False
        print("📂 正在更新文件...")
        for item in os.listdir(source_dir):
            s = os.path.join(source_dir, item)
            d = os.path.join(CURRENT_DIR, item)
            if item == "updater.py":
                continue
            if os.path.isdir(s):
                if os.path.exists(d):
                    shutil.rmtree(d)
                shutil.copytree(s, d)
            else:
                shutil.copy2(s, d)
        print("✅ 更新完成！")
        return True
    except Exception as e:
        print(f"❌ 更新失败: {e}")
        return False
    finally:
        try:
            if temp_zip and os.path.exists(temp_zip):
                os.remove(temp_zip)
            if extract_dir and os.path.exists(extract_dir):
                shutil.rmtree(extract_dir, ignore_errors=True)
        except:
            pass

def show_message_box(title, message):
    """显示 Windows 消息框，等待用户点击确定"""
    ctypes.windll.user32.MessageBoxW(0, message, title, 0)

if __name__ == "__main__":
    time.sleep(1)
    success = download_and_update()
    if success:
        show_message_box("更新完成", "Word格式处理工具已更新成功！\n点击确定后将自动启动。")
        # 重新启动主程序（使用 pythonw.exe 无控制台）
        main_py = os.path.join(CURRENT_DIR, "main.py")
        # 查找 pythonw.exe
        pythonw_exe = None
        if sys.executable.endswith("python.exe"):
            pythonw_exe = sys.executable.replace("python.exe", "pythonw.exe")
        else:
            # 在 PATH 中查找 pythonw
            pythonw_exe = shutil.which("pythonw")
        if pythonw_exe and os.path.exists(pythonw_exe):
            subprocess.Popen([pythonw_exe, main_py])
        else:
            # 回退到 python.exe（会有控制台）
            subprocess.Popen([sys.executable, main_py])
    else:
        show_message_box("更新失败", "更新失败，请检查网络或手动下载更新。")
