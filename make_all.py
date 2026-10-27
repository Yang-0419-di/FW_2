import glob
import hashlib
import os
import platform
from datetime import datetime
from importlib import metadata

base = os.path.dirname(os.path.abspath(__file__))

# 改用遞迴搜尋 **/*.py，並設定 recursive=True
files = sorted(
    f for f in glob.glob(os.path.join(base, "**", "*.py"), recursive=True)
    if os.path.basename(f) != "make_all.py"
)
files.append(os.path.join(base, "requirements.txt"))

lines = [f"===== SNAPSHOT {datetime.now():%Y-%m-%d %H:%M:%S} =====",
         f"Python {platform.python_version()} / {platform.system()} {platform.release()}"]

for pkg in ("yt-dlp", "bgutil-ytdlp-pot-provider", "curl_cffi", "playwright", "requests"):
    try:
        lines.append(f"{pkg}=={metadata.version(pkg)}")
    except metadata.PackageNotFoundError:
        lines.append(f"{pkg} (未安裝)")

lines.append("")

for path in files:
    # 取得相對於 base 的路徑（例如 utils/helper.py），若在同一層就只顯示檔名
    name = os.path.relpath(path, base)
    
    if not os.path.isfile(path):
        lines.append(f"===== FILE: {name} (不存在) =====\n")
        continue
    with open(path, "r", encoding="utf-8") as f:
        text = f.read()
    digest = hashlib.sha256(text.encode("utf-8")).hexdigest()[:8]
    lines.append(f"===== FILE: {name} (lines={len(text.splitlines())}, sha={digest}) =====")
    lines.append(text)
    lines.append("")

with open(os.path.join(base, "ALL.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

print("已產生包含子資料夾的 ALL.txt")