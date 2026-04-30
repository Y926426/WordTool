import os
import sys
import shutil
import zipfile
import tempfile
import subprocess
import time
import ctypes
import requests

REPO_URL = "https://github.com/Y926426/WordTool/archive/refs/heads/main.zip"
CURRENT_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))
LOG_FILE = os.path.join(CURRENT_DIR, "updater_log.txt")

def log(msg):
    with open(LOG_FILE, 'a', encoding='utf-8') as f:
        f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {msg}\n")
    print(msg)

def download_and_update():
    log("📥 正在从 GitHub 下载最新版本...")
    temp_zip = None
    extract_dir = None
    try:
        resp = requests.get(REPO_URL, stream=True, verify=False)
        resp.raise_for_status()
        temp_zip = os.path.join(tempfile.gettempdir(), "wordtool_latest.zip")
        with open(temp_zip, 'wb') as f:
            for chunk in resp.iter_content(chunk_size=8192):
                f.write(chunk)
        log("✅ 下载完成，正在解压...")
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
            log("❌ 未找到解压后的 WordTool 目录")
            return False
        log("📂 正在更新文件...")
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
        log("✅ 更新完成！")
        return True
    except Exception as e:
        log(f"❌ 更新失败: {e}")
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
    ctypes.windll.user32.MessageBoxW(0, message, title, 0)

if __name__ == "__main__":
    time.sleep(1)
    success = download_and_update()
    if success:
        show_message_box("更新完成", "Word格式处理工具已更新成功！\n请手动重新启动工具。")
        # 不再自动重启
    else:
        show_message_box("更新失败", "更新失败，请检查网络或手动下载更新。")
