import gspread
import os
from google.oauth2.service_account import Credentials

# ====== Google Sheet 認證（環境變數唯一來源） ======
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

DEFAULT_RENDER_SECRET = '/etc/secrets/disk-485810-82346bf9389a.json'

def get_google_client():
    secret_path = os.getenv('GOOGLE_SERVICE_ACCOUNT_FILE', DEFAULT_RENDER_SECRET)

    if not os.path.exists(secret_path):
        raise FileNotFoundError(
            f'❌ 找不到 Google Service Account JSON：{secret_path}'
        )

    print(f'🔐 使用 Service Account：{secret_path}')
    creds = Credentials.from_service_account_file(secret_path, scopes=SCOPES)
    return gspread.authorize(creds)

client = get_google_client()

# ====== googlesheet設定 ======

# Google Sheet ID
SHEET_ID = '1cFPw7C97a_xoqodcmvlWKPZJ2aBFvSBPqoE_PGPmxw0'  # ← 換成你的 ID

# 開啟工作表
sheet = client.open_by_key(SHEET_ID).sheet1  # 預設第一個工作表

def get_person_worksheet(person_name):
    sh = client.open_by_key(SHEET_ID)
    return sh.worksheet(person_name)
