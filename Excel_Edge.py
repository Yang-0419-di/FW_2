from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import os, time, shutil, glob
from datetime import datetime

# ===============================
# 基本設定
# ===============================
yyyymm = datetime.now().strftime("%Y%m")

download_path = r"D:\flask2\IM"
synology_im_path = r"D:\SynologyDrive\TOSHIBA\HL\保養\IM"

os.makedirs(download_path, exist_ok=True)
os.makedirs(synology_im_path, exist_ok=True)

# ===============================
# Edge 設定
# ===============================
options = webdriver.EdgeOptions()
options.use_chromium = True
options.add_experimental_option("prefs", {
    "download.default_directory": download_path,
    "download.prompt_for_download": False,
})

driver = webdriver.Edge(options=options)
wait = WebDriverWait(driver, 30)

# ===============================
# 共用工具
# ===============================
def wait_for_excel_button():
    """等待匯出按鈕 + 檢查是否空資料"""
    try:
        wait.until(EC.presence_of_element_located(
            (By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]')
        ))

        # 判斷是否空資料
        if "查無資料" in driver.page_source:
            return False

        return True

    except TimeoutException:
        return False


def wait_download(pattern, timeout=30):
    """等待下載完成（避免抓到 .crdownload）"""
    for _ in range(timeout):
        files = glob.glob(pattern)
        if files:
            latest = max(files, key=os.path.getctime)

            if not latest.endswith(".crdownload"):
                return latest

        time.sleep(1)

    return None


def download_report(menu_text, pattern, final_path, synology_path, pre_actions=None):
    """通用下載流程（含 retry + refresh）"""

    success = False

    for attempt in range(3):
        try:
            driver.switch_to.default_content()

            # 進入報表
            wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "服務資料查詢"))).click()
            time.sleep(1)
            wait.until(EC.element_to_be_clickable((By.LINK_TEXT, menu_text))).click()

            # iframe
            wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "iframe")))
            time.sleep(2)

            # 額外選項（POS 用）
            if pre_actions:
                pre_actions()

            # 查詢
            driver.find_element(By.XPATH, '//input[@type="submit" and @value="查詢"]').click()

            # 等結果
            if not wait_for_excel_button():
                raise Exception("查詢結果為空或未載入")

            # 刪舊檔
            for f in glob.glob(pattern):
                os.remove(f)

            # 下載
            driver.find_element(By.XPATH, '//input[@type="submit" and @value="匯出成EXCEL"]').click()

            downloaded = wait_download(pattern)

            if not downloaded:
                raise Exception("下載失敗")

            # 改名
            shutil.move(downloaded, final_path)

            # 備份
            try:
                shutil.copy2(final_path, synology_path)
            except Exception as e:
                print(f"⚠️ 複製失敗：{e}")

            print(f"✅ {menu_text} 完成")
            success = True
            break

        except Exception as e:
            print(f"⚠️ {menu_text} 第 {attempt+1} 次失敗：{e}")
            driver.refresh()
            time.sleep(3)

    if not success:
        raise Exception(f"{menu_text} 最終失敗")


# ===============================
# 登入流程
# ===============================
driver.get("http://eip.toshibatec.com.tw/Main.aspx")

wait.until(EC.presence_of_element_located((By.NAME, "AccountID"))).send_keys("yang.di")
driver.find_element(By.NAME, "PassWord").send_keys("foxdie789")
driver.find_element(By.NAME, "login_SubmitBtn").click()

# 進系統
wait.until(EC.element_to_be_clickable((By.XPATH, '//td[text()="內部系統"]'))).click()
time.sleep(1)
wait.until(EC.element_to_be_clickable((By.XPATH, '//td[contains(text(),"EIP 分析系統")]'))).click()

driver.switch_to.window(driver.window_handles[-1])
driver.refresh()
wait.until(EC.title_contains("台芝技術服務分析系統"))

# ===============================
# POS
# ===============================
pos_pattern = os.path.join(download_path, f"{yyyymm}_HL_Maintain_Report*.xlsx")
pos_final = os.path.join(download_path, f"{yyyymm}_HL_Maintain_Report.xlsx")
pos_synology = os.path.join(synology_im_path, f"{yyyymm}_HL_Maintain_Report.xlsx")

def pos_actions():
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//option[contains(text(),"萊爾富")]'))).click()
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//option[contains(text(),"新北勤務一部")]'))).click()

download_report("POS服務工作統計表", pos_pattern, pos_final, pos_synology, pos_actions)

# ===============================
# MFP
# ===============================
mfp_pattern = os.path.join(download_path, f"{yyyymm}_Service_Count_Report*.xlsx")
mfp_final = os.path.join(download_path, f"{yyyymm}_Service_Count_Report.xlsx")
mfp_synology = os.path.join(synology_im_path, f"{yyyymm}_Service_Count_Report.xlsx")

def mfp_actions():
    wait.until(EC.element_to_be_clickable(
        (By.XPATH, '//option[contains(text(),"新北勤務一部")]'))).click()

download_report("勤務工作統計表", mfp_pattern, mfp_final, mfp_synology, mfp_actions)

# ===============================
# 結束
# ===============================
driver.quit()
print("🎉 全部完成")