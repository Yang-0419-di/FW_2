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

try:
    # 使用 webdriver_manager 自動管理驅動版本，避免未來手動更新麻煩
    driver = webdriver.Edge(service=EdgeService(EdgeChromiumDriverManager().install()), options=options)
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

    # 偵測下載是否完成
    downloaded_pos = None
    for _ in range(30):
        files = glob.glob(pos_pattern)
        # 排除 Edge 的暫存檔以確保下載完成
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

    # ==================================================
    # 強力返回主選單機制
    # ==================================================
    print("🔄 正在嘗試返回主選單...")
    driver.switch_to.default_content() # 確保跳出 iframe
    time.sleep(1)

    try:
        # 嘗試方法 1：使用原本的 ID 定位（縮短等待到 5 秒，失敗立刻換方法）
        WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.ID, "back"))).click()
        print("▶️ 透過 ID 成功點擊返回")
    except Exception:
        try:
            # 嘗試方法 2：如果它是 <input type="button" value="返回"> 或類似的中文字按鈕
            back_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, '//input[@value="返回" or @value="回上一頁" or contains(@value,"回")]'))
            )
            back_btn.click()
            print("▶️ 透過 XPATH 中文按鈕成功點擊返回")
        except Exception:
            try:
                # 嘗試方法 3：使用 JavaScript 強行引發網頁的 back 函數（通常後台系統都有掛 window.history.back 或自訂 back 函式）
                driver.execute_script("if(typeof(back) === 'function'){ back(); } else { window.history.back(); }")
                print("▶️ 透過 JavaScript 強制返回成功")
            except Exception as e:
                print(f"❌ 所有返回方法皆失敗，錯誤資訊: {e}")
                # 如果真的都回不去，直接讓瀏覽器重新填寫網址進入「台芝技術服務分析系統」的首頁
                driver.get("http://eip.toshibatec.com.tw/Main.aspx") # 或者是分析系統的那個新視窗網址
                time.sleep(3)

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

    # 偵測下載是否完成
    downloaded_mfp = None
    for _ in range(30):
        files = glob.glob(mfp_pattern)
        # 排除 Edge 的暫存檔以確保下載完成
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

    driver.quit()

except Exception as e:
    print(f"❌ 發生錯誤：{e}")
    if 'driver' in locals():
        driver.quit()