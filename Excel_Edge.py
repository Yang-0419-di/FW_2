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

    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//option[contains(text(),"萊爾富")]'))).click()
    time.sleep(0.5)
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//option[contains(text(),"新北勤務一部")]'))).click()

    driver.find_element(By.XPATH, '//input[@type="submit" and @value="查詢"]').click()
    wait.until(EC.presence_of_element_located(
        (By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]')))

    # 刪除舊檔
    for f in glob.glob(pos_pattern):
        os.remove(f)

    driver.find_element(By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]').click()

    downloaded_pos = None
    for _ in range(30):
        files = glob.glob(pos_pattern)
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

    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//option[contains(text(),"新北勤務一部")]'))).click()

    driver.find_element(By.XPATH, '//input[@type="submit" and @value="查詢"]').click()
    wait.until(EC.presence_of_element_located(
        (By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]')))

    for f in glob.glob(mfp_pattern):
        os.remove(f)

    driver.find_element(By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]').click()

    downloaded_mfp = None
    for _ in range(30):
        files = glob.glob(mfp_pattern)
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

    driver.quit()

except Exception as e:
    print(f"❌ 發生錯誤：{e}")
    if 'driver' in locals():
        driver.quit()
