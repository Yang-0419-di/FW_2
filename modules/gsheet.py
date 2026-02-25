import gspread
import os
from google.oauth2.service_account import Credentials

# ====== Google Sheet 認證 ======
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
DEFAULT_RENDER_SECRET = '/etc/secrets/disk-485810-82346bf9389a.json'

def get_google_client():
    secret_path = os.getenv('GOOGLE_SERVICE_ACCOUNT_FILE', DEFAULT_RENDER_SECRET)
    if not os.path.exists(secret_path):
        raise FileNotFoundError(f'❌ 找不到 Google Service Account JSON：{secret_path}')
    print(f'🔐 使用 Service Account：{secret_path}')
    creds = Credentials.from_service_account_file(secret_path, scopes=SCOPES)
    return gspread.authorize(creds)

client = get_google_client()
SHEET_ID = '1cFPw7C97a_xoqodcmvlWKPZJ2aBFvSBPqoE_PGPmxw0'

def get_person_worksheet(person_name):
    sh = client.open_by_key(SHEET_ID)
    return sh.worksheet(person_name)

# ====== contracts 分頁 ======
def get_contract(device_id):
    sh = client.open_by_key(SHEET_ID)
    ws = sh.worksheet("contracts")
    records = ws.get_all_records()
    for r in records:
        if str(r.get("device_id")) == str(device_id):
            return r
    return None

# ====== customers 分頁 ======
def get_customer(device_id):
    sh = client.open_by_key(SHEET_ID)
    ws = sh.worksheet("customers")
    records = ws.get_all_records()
    for r in records:
        if str(r.get("device_id")) == str(device_id):
            return r
    return None

# ====== 模糊搜尋 customer_name ======
def search_customers_by_name(keyword):
    sh = client.open_by_key(SHEET_ID)
    ws = sh.worksheet("customers")
    records = ws.get_all_records()
    keyword_lower = keyword.lower()
    return [r for r in records if keyword_lower in str(r.get("customer_name", "")).lower()]

# 回傳 Customers worksheet
def get_customer_worksheet():
    sh = client.open_by_key(SHEET_ID)
    return sh.worksheet("customers")

# 回傳 Contracts worksheet
def get_contract_worksheet():
    sh = client.open_by_key(SHEET_ID)
    return sh.worksheet("contracts")

