import openpyxl
from datetime import datetime
import argparse
import sys

# 適用 Windows 與 Linux 的 timeout 輸入
def ask_input(timeout=10):
    print(f"請輸入版本號（{timeout} 秒內輸入，否則自動填入 MMDDHHMM）：")
    sys.stdout.write("> ")
    sys.stdout.flush()

    if sys.platform == "win32":
        import msvcrt, time
        start = time.time()
        input_str = ""
        while True:
            if msvcrt.kbhit():
                char = msvcrt.getwch()
                if char in ("\r", "\n"):  # 按 Enter 結束
                    break
                elif char == "\b":  # Backspace
                    input_str = input_str[:-1]
                    sys.stdout.write("\b \b")
                    sys.stdout.flush()
                else:
                    input_str += char
                    sys.stdout.write(char)
                    sys.stdout.flush()
            if (time.time() - start) > timeout:
                return None
        return input_str.strip()
    else:
        # Linux / Mac
        import select
        i, _, _ = select.select([sys.stdin], [], [], timeout)
        if i:
            return sys.stdin.readline().strip()
        else:
            return None

def get_version(timeout=10, auto_only=False):
    version = None

    if not auto_only:
        user_input = ask_input(timeout)
        if user_input:
            version = user_input

    if not version:
        version = datetime.now().strftime("%m%d%H%M")
        if not auto_only:
            print(f"\n[超時] 超過 {timeout} 秒未輸入，自動使用版本號 {version}")

    # 寫入 Excel
    file_path = "data.xlsx"
    wb = openpyxl.load_workbook(file_path)
    sheet = wb["首頁"]
    sheet["G1"] = version
    wb.save(file_path)

    print(f"[完成] 已將版本號 {version} 寫入 G1")
    # 輸出給 BAT
    print(version)
    return version

if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--auto", action="store_true", help="Git-only 自動版本號")
    args = parser.parse_args()
    get_version(auto_only=args.auto)
