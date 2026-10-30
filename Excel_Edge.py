from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service as EdgeService
from webdriver_manager.microsoft import EdgeChromiumDriverManager
from selenium.common.exceptions import TimeoutException
import os
import time
import shutil
import glob
from datetime import datetime
import sys

EIP_USER = os.environ.get("EIP_USER", "")
EIP_PASS = os.environ.get("EIP_PASS", "")
if not EIP_USER or not EIP_PASS:
    sys.exit("❌ 請先設定環境變數 EIP_USER / EIP_PASS（例：setx EIP_USER 帳號、setx EIP_PASS 密碼，重開終端機生效）")

# ===============================
# 時間與資料夾設定
# ===============================
yyyymm = datetime.now().strftime("%Y%m")

# Edge 實際下載位置（IM 保留）
download_path = r"D:\flask2\IM"

# Synology 同步位置（複製一份）
synology_im_path = r"D:\SynologyDrive\TOSHIBA\HL\保養\IM"

os.makedirs(download_path, exist_ok=True)
os.makedirs(synology_im_path, exist_ok=True)

# 檔案樣式
pos_pattern = os.path.join(download_path, f"{yyyymm}_HL_Maintain_Report*.xlsx")
mfp_pattern = os.path.join(download_path, f"{yyyymm}_Service_Count_Report*.xlsx")

pos_final = os.path.join(download_path, f"{yyyymm}_HL_Maintain_Report.xlsx")
mfp_final = os.path.join(download_path, f"{yyyymm}_Service_Count_Report.xlsx")

pos_synology = os.path.join(synology_im_path, f"{yyyymm}_HL_Maintain_Report.xlsx")
mfp_synology = os.path.join(synology_im_path, f"{yyyymm}_Service_Count_Report.xlsx")

# ===============================
# Edge 選項
# ===============================
options = webdriver.EdgeOptions()
options.use_chromium = True
options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})
# 穩定性優化參數
options.add_argument('--disable-gpu')
options.add_argument('--no-sandbox')

failures = []

try:
    # 使用 webdriver_manager 自動管理驅動版本，避免未來手動更新麻煩
    driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 30)

    # ===============================
    # 登入
    # ===============================
    driver.get("https://eip.toshibatec.com.tw/Main.aspx")
    wait.until(EC.presence_of_element_located((By.NAME, "AccountID"))).send_keys(EIP_USER)
    driver.find_element(By.NAME, "PassWord").send_keys(EIP_PASS)
    driver.find_element(By.NAME, "login_SubmitBtn").click()

    # ===============================
    # 進入系統
    # ===============================
    wait.until(EC.element_to_be_clickable((By.XPATH, '//td[text()="內部系統"]'))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.XPATH, '//td[contains(text(),"EIP 分析系統")]'))).click()
    time.sleep(3)

    driver.switch_to.window(driver.window_handles[-1])
    driver.refresh()
    wait.until(EC.title_contains("台芝技術服務分析系統"))

    # ==================================================
    # POS 服務工作統計表 (含查無資料重試機制)
    # ==================================================
    while True:
        driver.switch_to.default_content()
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "服務資料查詢"))).click()
        time.sleep(1)
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "POS服務工作統計表"))).click()
        driver.switch_to.frame("iframe")
        time.sleep(2)

        if "Warning: mysql" in driver.page_source:
            raise Exception("Warning: mysql detected")

        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//option[contains(text(),"萊爾富")]'))).click()
        time.sleep(0.5)
        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//option[contains(text(),"新北勤務一部")]'))).click()

        driver.find_element(By.XPATH, '//input[@type="submit" and @value="查詢"]').click()
        time.sleep(2)

        # 檢查是否出現「查無資料」
        if "查無資料" in driver.page_source:
            print("⚠️ POS 查詢結果顯示「查無資料」，正在返回重新執行...")
            driver.switch_to.default_content()
            # 嘗試返回上一頁或重新點選選單重試
            try:
                driver.execute_script("if(typeof(back) === 'function'){ back(); } else { window.history.back(); }")
            except Exception:
                pass
            time.sleep(2)
            continue
        else:
            # 確認已進入正確結果頁面（出現匯出按鈕）
            wait.until(EC.presence_of_element_located(
                (By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]')))
            break

    # 刪除舊檔
    for f in glob.glob(pos_pattern):
        os.remove(f)

    driver.find_element(By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]').click()

    # 偵測下載是否完成
    downloaded_pos = None
    for _ in range(30):
        files = glob.glob(pos_pattern)
        files = [f for f in files if not f.endswith('.crdownload')]
        if files:
            downloaded_pos = max(files, key=os.path.getctime)
            break
        time.sleep(1)

    if downloaded_pos:
        shutil.move(downloaded_pos, pos_final)
        try:
            shutil.copy2(pos_final, pos_synology)
        except Exception as e:
            print(f"⚠️ POS 複製至 Synology 失敗：{e}")

        print("✅ POS 報表完成")
        print(f"   IM：{pos_final}")
        print(f"   Synology：{pos_synology}")
    else:
        print("❌ POS 報表未下載完成")
        failures.append("POS")

    # ==================================================
    # 強力返回主選單機制
    # ==================================================
    print("🔄 正在嘗試返回主選單...")
    driver.switch_to.default_content()
    time.sleep(1)

    try:
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "back"))).click()
        print("▶️ 透過 ID 成功點擊返回")
    except Exception:
        try:
            back_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//input[@value="返回" or @value="回上一頁" or contains(@value,"回")]'))
            )
            back_btn.click()
            print("▶️ 透過 XPATH 中文按鈕成功點擊返回")
        except Exception:
            try:
                driver.execute_script("if(typeof(back) === 'function'){ back(); } else { window.history.back(); }")
                print("▶️ 透過 JavaScript 強制返回成功")
            except Exception as e:
                print(f"❌ 所有返回方法皆失敗，錯誤資訊: {e}")
                driver.get("https://eip.toshibatec.com.tw/Main.aspx")
                time.sleep(3)

    time.sleep(2)

    # ==================================================
    # MFP 勤務工作統計表 (含查無資料重試機制)
    # ==================================================
    while True:
        driver.switch_to.default_content()
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "服務資料查詢"))).click()
        time.sleep(1)
        wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "勤務工作統計表"))).click()
        driver.switch_to.frame("iframe")
        time.sleep(2)

        if "Warning: mysql" in driver.page_source:
            raise Exception("Warning: mysql detected")

        wait.until(EC.element_to_be_clickable(
            (By.XPATH, '//option[contains(text(),"新北勤務一部")]'))).click()

        driver.find_element(By.XPATH, '//input[@type="submit" and @value="查詢"]').click()
        time.sleep(2)

        # 檢查是否出現「查無資料」
        if "查無資料" in driver.page_source:
            print("⚠️ MFP 查詢結果顯示「查無資料」，正在返回重新執行...")
            driver.switch_to.default_content()
            try:
                driver.execute_script("if(typeof(back) === 'function'){ back(); } else { window.history.back(); }")
            except Exception:
                pass
            time.sleep(2)
            continue
        else:
            wait.until(EC.presence_of_element_located(
                (By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]')))
            break

    for f in glob.glob(mfp_pattern):
        os.remove(f)

    driver.find_element(By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]').click()

    # 偵測下載是否完成
    downloaded_mfp = None
    for _ in range(30):
        files = glob.glob(mfp_pattern)
        files = [f for f in files if not f.endswith('.crdownload')]
        if files:
            downloaded_mfp = max(files, key=os.path.getctime)
            break
        time.sleep(1)

    if downloaded_mfp:
        shutil.move(downloaded_mfp, mfp_final)
        try:
            shutil.copy2(mfp_final, mfp_synology)
        except Exception as e:
            print(f"⚠️ MFP 複製至 Synology 失敗：{e}")

        print("✅ MFP 報表完成")
        print(f"   IM：{mfp_final}")
        print(f"   Synology：{mfp_synology}")
    else:
        print("❌ MFP 報表未下載完成")
        failures.append("MFP")

    driver.quit()
    if failures:
        print(f"❌ 下載失敗項目：{', '.join(failures)}")
        sys.exit(1)

except Exception as e:
    print(f"❌ 發生錯誤：{e}")
    if 'driver' in locals():
        driver.quit()
    sys.exit(1)