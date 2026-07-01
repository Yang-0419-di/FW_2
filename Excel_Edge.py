from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import os
import time
import shutil
import glob
from datetime import datetime

# ===============================
# 時間與資料夾設定
# ===============================
yyyymm = datetime.now().strftime("%Y%m")

download_path = r"D:\flask2\IM"
synology_im_path = r"D:\SynologyDrive\TOSHIBA\HL\保養\IM"

os.makedirs(download_path, exist_ok=True)
os.makedirs(synology_im_path, exist_ok=True)

# 💡 修正：瀏覽器剛下載時的可能檔名關鍵字（請依據網頁實際下載的檔名修改）
pos_download_pattern = os.path.join(download_path, "*Report*.xlsx") 
mfp_download_pattern = os.path.join(download_path, "*統計表*.xlsx") 

pos_final = os.path.join(download_path, f"{yyyymm}_HL_Maintain_Report.xlsx")
mfp_final = os.path.join(download_path, f"{yyyymm}_Service_Count_Report.xlsx")

pos_synology = os.path.join(synology_im_path, f"{yyyymm}_HL_Maintain_Report.xlsx")
mfp_synology = os.path.join(synology_im_path, f"{yyyymm}_Service_Count_Report.xlsx")

# ===============================
# Edge 選項
# ===============================
options = webdriver.EdgeOptions()
options.use_chromium = True

# 💡 修正：隱藏 DevTools 提示與過濾無用 ERROR 日誌
options.add_experimental_option("excludeSwitches", ["enable-logging"])
options.add_argument('--log-level=3')

options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
    "safebrowsing.enabled": True
})

def wait_for_download_ready(directory, pattern, timeout=30):
    """確保檔案下載完成，且沒有 .crdownload 暫存檔"""
    for _ in range(timeout):
        # 檢查是否有未完成的暫存檔
        if glob.glob(os.path.join(directory, "*.crdownload")):
            time.sleep(1)
            continue
        files = glob.glob(pattern)
        if files:
            return max(files, key=os.path.getctime)
        time.sleep(1)
    return None

try:
    driver = webdriver.Edge(options=options)
    wait = WebDriverWait(driver, 30)

    # ===============================
    # 登入
    # ===============================
    driver.get("http://eip.toshibatec.com.tw/Main.aspx")
    wait.until(EC.presence_of_element_located((By.NAME, "AccountID"))).send_keys("yang.di")
    driver.find_element(By.NAME, "PassWord").send_keys("foxdie789")
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
    # POS 服務工作統計表
    # ==================================================
    driver.switch_to.default_content()
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "服務資料查詢"))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "POS服務工作統計表"))).click()
    driver.switch_to.frame("iframe")
    time.sleep(2)

    if "Warning: mysql" in driver.page_source:
        raise Exception("Warning: mysql detected")

    wait.until(EC.element_to_be_clickable((By.XPATH, '//option[contains(text(),"萊爾富")]'))).click()
    time.sleep(0.5)
    wait.until(EC.element_to_be_clickable((By.XPATH, '//option[contains(text(),"新北勤務一部")]'))).click()

    driver.find_element(By.XPATH, '//input[@type="submit" and @value="查詢"]').click()
    wait.until(EC.presence_of_element_located((By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]')))

    # 刪除舊的 final 檔案防止衝突
    if os.path.exists(pos_final): os.remove(pos_final)

    driver.find_element(By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]').click()

    # 💡 使用改良後的下載等待機制
    downloaded_pos = wait_for_download_ready(download_path, pos_download_pattern)

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

    wait.until(EC.element_to_be_clickable((By.ID, "back"))).click()
    time.sleep(2)

    # ==================================================
    # MFP 勤務工作統計表
    # ==================================================
    driver.switch_to.default_content()
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "服務資料查詢"))).click()
    time.sleep(1)
    wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "勤務工作統計表"))).click()
    driver.switch_to.frame("iframe")
    time.sleep(2)

    if "Warning: mysql" in driver.page_source:
        raise Exception("Warning: mysql detected")

    wait.until(EC.element_to_be_clickable((By.XPATH, '//option[contains(text(),"新北勤務一部")]'))).click()

    driver.find_element(By.XPATH, '//input[@type="submit" and @value="查詢"]').click()
    wait.until(EC.presence_of_element_located((By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]')))

    if os.path.exists(mfp_final): os.remove(mfp_final)

    driver.find_element(By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]').click()

    # 💡 使用改良後的下載等待機制
    downloaded_mfp = wait_for_download_ready(download_path, mfp_download_pattern)

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

    driver.quit()

except Exception as e:
    print(f"❌ 發生錯誤：{e}")
    if 'driver' in locals():
        driver.quit()