import os
import sys
import shutil
import urllib.request
import zipfile
import tempfile
import subprocess
import time

# GitHub 仓库的 ZIP 包地址（公开仓库）
REPO_URL = "https://github.com/Y926426/WordTool/archive/refs/heads/main.zip"

# 当前程序所在目录（即 WordTool 文件夹）
CURRENT_DIR = os.path.dirname(os.path.abspath(sys.argv[0]))

def download_and_update():
    print("📥 正在从 GitHub 下载最新版本...")
    temp_zip = None
    extract_dir = None
    try:
        # 1. 下载 ZIP 到临时文件
        temp_zip = os.path.join(tempfile.gettempdir(), "wordtool_latest.zip")
        urllib.request.urlretrieve(REPO_URL, temp_zip)
        print("✅ 下载完成，正在解压...")

        # 2. 解压到临时目录
        extract_dir = tempfile.mkdtemp()
        with zipfile.ZipFile(temp_zip, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)

        # 3. 找到解压后的最顶层文件夹（通常是 WordTool-main 或类似）
        source_dir = None
        for item in os.listdir(extract_dir):
            full = os.path.join(extract_dir, item)
            if os.path.isdir(full) and item.startswith("WordTool"):
                source_dir = full
                break
        if not source_dir:
            print("❌ 未找到解压后的 WordTool 目录")
            return False

        # 4. 复制文件到 CURRENT_DIR（跳过 updater.py，避免覆盖正在运行的自己）
        print("📂 正在更新文件...")
        for item in os.listdir(source_dir):
            s = os.path.join(source_dir, item)
            d = os.path.join(CURRENT_DIR, item)

            # 跳过 updater.py 自身，以免覆盖导致冲突
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
        # 清理临时文件
        try:
            if temp_zip and os.path.exists(temp_zip):
                os.remove(temp_zip)
            if extract_dir and os.path.exists(extract_dir):
                shutil.rmtree(extract_dir, ignore_errors=True)
        except:
            pass

if __name__ == "__main__":
    # 等待 1 秒，确保主程序已经彻底关闭
    time.sleep(1)
    success = download_and_update()
    if success:
        print("🚀 正在重新启动主程序...")
        # 启动新版的 main.py
        subprocess.Popen([sys.executable, os.path.join(CURRENT_DIR, "main.py")])
    else:
        print("❌ 更新失败，请检查网络或手动下载。")
        input("按回车键退出...")
