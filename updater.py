import os
import sys
import shutil
import urllib.request
import zipfile
import tempfile
import subprocess
import time

# 配置
REPO_URL = "https://github.com/Y926426/WordTool/archive/refs/heads/main.zip"
# 当前程序所在目录（即 WordTool 文件夹）
CURRENT_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

def download_and_update():
    print("正在下载最新版本...")
    try:
        # 下载 ZIP 到临时文件
        temp_zip = os.path.join(tempfile.gettempdir(), "wordtool_latest.zip")
        urllib.request.urlretrieve(REPO_URL, temp_zip)
        # 解压到临时目录
        extract_dir = tempfile.mkdtemp()
        with zipfile.ZipFile(temp_zip, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
        # 解压后的文件夹名可能是 WordTool-main
        source_dir = os.path.join(extract_dir, "WordTool-main")
        if not os.path.exists(source_dir):
            # 可能是其他命名，尝试查找
            for item in os.listdir(extract_dir):
                if os.path.isdir(os.path.join(extract_dir, item)) and item.startswith("WordTool"):
                    source_dir = os.path.join(extract_dir, item)
                    break
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
        # 清理临时文件
        try:
            os.remove(temp_zip)
        except:
            pass

if __name__ == "__main__":
    # 等待一小段时间，确保主程序已关闭
    time.sleep(1)
    success = download_and_update()
    if success:
        # 重新启动主程序
        subprocess.Popen([sys.executable, os.path.join(CURRENT_DIR, "main.py")])
    else:
        input("更新失败，按回车键退出...")