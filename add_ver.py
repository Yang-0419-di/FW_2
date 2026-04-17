import openpyxl
from datetime import datetime
import argparse
import sys
import os
import time

def ask_input(timeout=10):
    print(f"請輸入版本號（{timeout} 秒內輸入，否則自動填入 MMDDHHMM）：")
    sys.stdout.write("> ")
    sys.stdout.flush()

    if sys.platform == "win32":
        import msvcrt, time as t
        start = t.time()
        input_str = ""
        while True:
            if msvcrt.kbhit():
                char = msvcrt.getwch()
                if char in ("\r", "\n"):
                    break
                elif char == "\b":
                    input_str = input_str[:-1]
                    sys.stdout.write("\b \b")
                    sys.stdout.flush()
                else:
                    input_str += char
                    sys.stdout.write(char)
                    sys.stdout.flush()
            if (t.time() - start) > timeout:
                return None
        return input_str.strip()
    else:
        import select
        i, _, _ = select.select([sys.stdin], [], [], timeout)
        if i:
            return sys.stdin.readline().strip()
        return None

def safe_save(wb, file_path):
    tmp_path = file_path + ".tmp"

    wb.save(tmp_path)
    wb.close()

    time.sleep(1)  # ⭐ 關鍵：避免 IO 還沒寫完

    os.replace(tmp_path, file_path)

def get_version(timeout=10, auto_only=False):
    version = None

    if not auto_only:
        user_input = ask_input(timeout)
        if user_input:
            version = user_input

    if not version:
        version = datetime.now().strftime("%m%d%H%M")
        if not auto_only:
            print(f"\n[超時] 自動使用版本號 {version}")

    file_path = "data.xlsx"

    wb = openpyxl.load_workbook(file_path)
    sheet = wb["首頁"]
    sheet["G1"] = version

    # ⭐ 改成安全寫入
    safe_save(wb, file_path)

    print(f"[完成] 已寫入版本號 {version}")
    print(version)
    return version

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--auto", action="store_true")
    args = parser.parse_args()

    try:
        get_version(auto_only=args.auto)
        sys.exit(0)
    except Exception as e:
        print(f"❌ 錯誤：{e}")
        sys.exit(1)