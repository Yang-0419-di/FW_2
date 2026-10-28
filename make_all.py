import glob
import hashlib
import os
import platform
from datetime import datetime
from importlib import metadata

base = os.path.dirname(os.path.abspath(__file__))

# 1. 遞迴搜尋所有的 .py 檔案（排除 make_all.py）
py_files = sorted(
    f for f in glob.glob(os.path.join(base, "**", "*.py"), recursive=True)
    if os.path.basename(f) != "make_all.py"
)

# 2. 遞迴搜尋 templates 資料夾內所有的 .html 檔案
html_files = sorted(
    glob.glob(os.path.join(base, "templates", "**", "*.html"), recursive=True)
)

# 3. 合併檔案列表（Python 檔 + HTML 檔 + requirements.txt）
files = py_files + html_files
files.append(os.path.join(base, "requirements.txt"))

lines = [
    f"===== SNAPSHOT {datetime.now():%Y-%m-%d %H:%M:%S} =====",
    f"Python {platform.python_version()} / {platform.system()} {platform.release()}"
]

for pkg in ("yt-dlp", "bgutil-ytdlp-pot-provider", "curl_cffi", "playwright", "requests"):
    try:
        lines.append(f"{pkg}=={metadata.version(pkg)}")
    except metadata.PackageNotFoundError:
        lines.append(f"{pkg} (未安裝)")

lines.append("")

for path in files:
    # 取得相對於 base 的路徑（例如 templates/index.html 或 utils/helper.py）
    name = os.path.relpath(path, base)
    
    if not os.path.isfile(path):
        lines.append(f"===== FILE: {name} (不存在) =====\n")
        continue
    
    # 讀取檔案內容（使用 utf-8 編碼）
    with open(path, "r", encoding="utf-8") as f:
        text = f.read()
        
    digest = hashlib.sha256(text.encode("utf-8")).hexdigest()[:8]
    lines.append(f"===== FILE: {name} (lines={len(text.splitlines())}, sha={digest}) =====")
    lines.append(text)
    lines.append("")

with open(os.path.join(base, "ALL.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

print("已產生包含子資料夾及 templates 內 HTML 檔的 ALL.txt")