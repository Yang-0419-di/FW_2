import hashlib
import os
import platform
from datetime import datetime
from importlib import metadata

base = os.path.dirname(os.path.abspath(__file__))

# ===== 設定 =====
# 要收錄的副檔名（整個資料夾含子資料夾）
INCLUDE_EXT = {".py", ".html", ".css", ".js"}
# 一定要收錄的單一檔案（不存在時會標示「不存在」）
EXTRA_FILES = ["requirements.txt"]
# 不收錄的檔案
EXCLUDE_FILES = {"make_all.py"}
# 不進入的資料夾（套件、快取、版本控制等，與程式內容無關）
EXCLUDE_DIRS = {".git", "__pycache__", "venv", ".venv", "env", "node_modules", ".idea", ".vscode", "dist", "build"}
# 是否收錄備份資料夾（名稱含「備份」或 _fix_backup）；改成 False 可讓 ALL.txt 小很多
INCLUDE_BACKUPS = False
# 壓縮過的第三方套件（*.min.js / *.min.css）或超過此大小的檔案不收錄，避免 ALL.txt 暴增
MAX_BYTES = 300 * 1024


def is_backup_dir(name):
    return "備份" in name or name.lower().startswith("_fix_backup")


def read_text(path):
    """依序嘗試 UTF-8、Big5（cp950）；都失敗就用 UTF-8 並把無法解碼的字元替換掉，不讓整個工具中斷。"""
    raw = open(path, "rb").read()
    for enc in ("utf-8-sig", "cp950"):
        try:
            return raw.decode(enc), enc
        except UnicodeDecodeError:
            pass
    return raw.decode("utf-8", errors="replace"), "utf-8(部分字元無法解碼)"


# 1. 遞迴搜尋所有要收錄的檔案
files, skipped = [], []
for root, dirs, names in os.walk(base):
    dirs[:] = sorted(d for d in dirs
                     if d not in EXCLUDE_DIRS and (INCLUDE_BACKUPS or not is_backup_dir(d)))
    for n in sorted(names):
        ext = os.path.splitext(n)[1].lower()
        if ext not in INCLUDE_EXT or n in EXCLUDE_FILES:
            continue
        path = os.path.join(root, n)
        rel = os.path.relpath(path, base)
        if n.lower().endswith((".min.js", ".min.css")):
            skipped.append(f"{rel}（壓縮過的第三方檔案）")
            continue
        size = os.path.getsize(path)
        if size > MAX_BYTES:
            skipped.append(f"{rel}（{size // 1024} KB，超過 {MAX_BYTES // 1024} KB）")
            continue
        files.append(path)

# 排序：與原本相同，Python 檔在前，其次 HTML、CSS、JS（各自依路徑排序）
order = {".py": 0, ".html": 1, ".css": 2, ".js": 3}
files.sort(key=lambda p: (order.get(os.path.splitext(p)[1].lower(), 9), os.path.relpath(p, base).lower()))
files += [os.path.join(base, f) for f in EXTRA_FILES]

lines = [
    f"===== SNAPSHOT {datetime.now():%Y-%m-%d %H:%M:%S} =====",
    f"Python {platform.python_version()} / {platform.system()} {platform.release()}"
]

for pkg in ("yt-dlp", "bgutil-ytdlp-pot-provider", "curl_cffi", "playwright", "requests"):
    try:
        lines.append(f"{pkg}=={metadata.version(pkg)}")
    except metadata.PackageNotFoundError:
        lines.append(f"{pkg} (未安裝)")

if skipped:
    lines.append("")
    lines.append("===== 未收錄的檔案 =====")
    lines.extend("  " + s for s in skipped)

lines.append("")

count = {}
for path in files:
    # 取得相對於 base 的路徑（例如 templates\index.html 或 static\css\mobile.css）
    name = os.path.relpath(path, base)

    if not os.path.isfile(path):
        lines.append(f"===== FILE: {name} (不存在) =====\n")
        continue

    text, enc = read_text(path)
    # 統一換行為 \n（寫檔時 Windows 會自動轉成 \r\n；不先統一會變成 \r\r\n，每行多一個空行）
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    digest = hashlib.sha256(text.encode("utf-8")).hexdigest()[:8]
    note = "" if enc in ("utf-8-sig",) else f", encoding={enc}"
    lines.append(f"===== FILE: {name} (lines={len(text.splitlines())}, sha={digest}{note}) =====")
    lines.append(text)
    lines.append("")
    ext = os.path.splitext(name)[1].lower() or name
    count[ext] = count.get(ext, 0) + 1

with open(os.path.join(base, "ALL.txt"), "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

size_kb = os.path.getsize(os.path.join(base, "ALL.txt")) // 1024
print("已產生 ALL.txt（" + "、".join(f"{k} {v} 個" for k, v in sorted(count.items())) + f"，共 {size_kb} KB）")
if skipped:
    print("未收錄：")
    for s in skipped:
        print("  " + s)