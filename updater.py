import os
import sys
import shutil
import urllib.request
import zipfile
import tempfile
import subprocess
import time

REPO_URL = "https://github.com/Y926426/WordTool/archive/refs/heads/main.zip"
CURRENT_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

def download_and_update():
    print("正在下载最新版本...")
    try:
        temp_zip = os.path.join(tempfile.gettempdir(), "wordtool_latest.zip")
        urllib.request.urlretrieve(REPO_URL, temp_zip)
        extract_dir = tempfile.mkdtemp()
        with zipfile.ZipFile(temp_zip, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        # 查找解压出的文件夹（可能是 WordTool-main 或 WordTool-xxx）
        source_dir = None
        for item in os.listdir(extract_dir):
            full = os.path.join(extract_dir, item)
            if os.path.isdir(full) and item.startswith("WordTool"):
                source_dir = full
                break
        if not source_dir:
            print("错误：未找到解压后的 WordTool 目录")
            return False

        # 复制文件到 CURRENT_DIR（覆盖）
        for item in os.listdir(source_dir):
            s = os.path.join(source_dir, item)
            d = os.path.join(CURRENT_DIR, item)
            if os.path.isdir(s):
                if os.path.exists(d):
                    shutil.rmtree(d)
                shutil.copytree(s, d)
            else:
                shutil.copy2(s, d)
        print("更新完成！")
        return True
    except Exception as e:
        print(f"更新失败: {e}")
        return False
    finally:
        try:
            os.remove(temp_zip)
        except:
            pass

if __name__ == "__main__":
    time.sleep(1)  # 等待原程序完全退出
    success = download_and_update()
    if success:
        subprocess.Popen([sys.executable, os.path.join(CURRENT_DIR, "main.py")])
    else:
        input("更新失败，按回车键退出...")
