import os
import sys
import shutil
import zipfile
import tempfile
import time
import ctypes
import requests

REPO_URL = "https://github.com/Y926426/WordTool/archive/refs/heads/main.zip"
CURRENT_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

def log(msg):
    print(msg)
    with open(os.path.join(CURRENT_DIR, "updater_log.txt"), 'a', encoding='utf-8') as f:
        f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} - {msg}\n")

def download_and_update():
    log("📥 正在从 GitHub 下载最新版本...")
    temp_zip = os.path.join(tempfile.gettempdir(), "wordtool_latest.zip")
    extract_dir = tempfile.mkdtemp()
    try:
        # 下载
        resp = requests.get(REPO_URL, stream=True, verify=False)
        resp.raise_for_status()
        with open(temp_zip, 'wb') as f:
            for chunk in resp.iter_content(chunk_size=8192):
                f.write(chunk)
        log("✅ 下载完成，正在解压...")
        with zipfile.ZipFile(temp_zip, 'r') as zf:
            zf.extractall(extract_dir)
        # 找到解压后的根目录（可能是 WordTool-main 或 WordTool-xxx）
        source_dir = None
        for name in os.listdir(extract_dir):
            full = os.path.join(extract_dir, name)
            if os.path.isdir(full) and name.startswith("WordTool"):
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
        # 额外检查 version.txt 是否被更新
        new_version_file = os.path.join(CURRENT_DIR, "version.txt")
        if os.path.exists(new_version_file):
            with open(new_version_file, 'r') as f:
                ver = f.read().strip()
            log(f"✅ 更新后的版本号: {ver}")
        else:
            log("⚠️ warning: version.txt 不存在")
        return True
    except Exception as e:
        log(f"❌ 更新失败: {e}")
        return False
    finally:
        try:
            os.remove(temp_zip)
            shutil.rmtree(extract_dir, ignore_errors=True)
        except:
            pass

def show_msg(title, msg):
    ctypes.windll.user32.MessageBoxW(0, msg, title, 0)

if __name__ == "__main__":
    time.sleep(1)
    if download_and_update():
        show_msg("更新完成", "Word格式处理工具已更新成功！\n请手动重新启动工具。")
    else:
        show_msg("更新失败", "更新失败，请检查网络或手动下载更新。")
