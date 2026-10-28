from flask import Blueprint, request, jsonify
import pandas as pd
import gspread
from modules.gsheet import client, SHEET_ID

sc_check_bp = Blueprint('sc_check', __name__)

def get_sc_check_sheet():
    sh = client.open_by_key(SHEET_ID)
    return sh.worksheet("硬碟檢測")

def fetch_and_sync_sc_check(xls):
    """
    從 Excel IM 分頁讀取資料，並與 Google Sheets '硬碟檢測' 分頁同步。
    篩選條件：
      1. 報修類別為 'HL-TM主機' 或 'HL-SC主機'
      2. 內容 包含 '更換' 與 '硬碟'
    取最新日期的前 5 筆資料回傳。
    """
    try:
        # 1. 讀取 Excel IM 分頁
        df_im = pd.read_excel(xls, sheet_name='IM')
        df_im.columns = df_im.columns.astype(str).str.replace('\n', '', regex=False).str.strip()
        
        print(f"📌 [DEBUG] Excel IM 總行數: {len(df_im)}")
        print(f"📌 [DEBUG] Excel 欄位名稱列表: {list(df_im.columns)}")
        
        # 尋找台芝工作案號欄位 (相容自動換行)
        case_no_col = None
        for col in df_im.columns:
            clean_col = col.replace('\n', '').replace(' ', '')
            if '台芝' in clean_col and '工作案號' in clean_col:
                case_no_col = col
                break
        if not case_no_col:
            case_no_col = '台芝工作案號'
        print(f"📌 [DEBUG] 決定使用的案號欄位: {case_no_col}")

        # 動態判定內容欄位（將您的「報修項目」或「工作內容」納入判斷）
        content_col = None
        for c in df_im.columns:
            if '工作內容' in c or '報修說明' in c or '報修項目' in c:
                content_col = c
                break
        if not content_col:
            content_col = df_im.columns[-1]
        print(f"📌 [DEBUG] 決定使用的內容/說明欄位: {content_col}")
        
        # 檢查報修類別欄位是否存在
        if '報修類別' not in df_im.columns:
            print("❌ [DEBUG 錯誤] Excel 中找不到 '報修類別' 欄位！")
            return []

        # 條件 1: 報修類別 == 'HL-TM主機' 或 'HL-SC主機'
        mask_category = df_im['報修類別'].astype(str).str.strip().isin(['HL-TM主機', 'HL-SC主機'])
        print(f"📌 [DEBUG] 符合類別 (HL-TM/HL-SC) 的筆數: {mask_category.sum()}")
        
        # 條件 2: 內容包含 '更換' 與 '硬碟'
        content_series = df_im[content_col].astype(str)
        mask_keyword_change = content_series.str.contains('更換', case=False, na=False)
        mask_keyword_hdd = content_series.str.contains('硬碟', case=False, na=False)
        mask_keywords = mask_keyword_change & mask_keyword_hdd
        
        print(f"📌 [DEBUG] 內容包含 '更換' 的筆數: {mask_keyword_change.sum()}")
        print(f"📌 [DEBUG] 內容包含 '硬碟' 的筆數: {mask_keyword_hdd.sum()}")
        print(f"📌 [DEBUG] 同時包含 '更換' 與 '硬碟' 的筆數: {mask_keywords.sum()}")
        
        # 綜合篩選
        filtered_im = df_im[mask_category & mask_keywords].copy()
        print(f"📌 [DEBUG] 雙重條件篩選後的總筆數: {len(filtered_im)}")
        
        if filtered_im.empty:
            print("⚠️ [DEBUG] 篩選後無任何符合資料，直接回傳空陣列 []")
            return []

        # 排序：最新日期在最前面
        time_col = None
        for t in ['離場時間', '報修時間']:
            if t in filtered_im.columns:
                time_col = t
                break
        if not time_col:
            time_col = filtered_im.columns[0]
            
        print(f"📌 [DEBUG] 決定使用的時間欄位: {time_col}")
        filtered_im['sort_time'] = pd.to_datetime(filtered_im[time_col], errors='coerce')
        
        # 檢查是否有時間解析失敗 (NaT)
        nat_count = filtered_im['sort_time'].isna().sum()
        if nat_count > 0:
            print(f"⚠️ [DEBUG] 注意：有 {nat_count} 筆資料的時間無法被正常解析 (變成 NaT)")

        filtered_im = filtered_im.sort_values(by='sort_time', ascending=False)

        # 2. 讀取 Google Sheet 硬碟檢測現有紀錄
        sheet = get_sc_check_sheet()
        gs_records = sheet.get_all_records()
        gs_df = pd.DataFrame(gs_records) if gs_records else pd.DataFrame()
        print(f"📌 [DEBUG] Google Sheet 現有記錄筆數: {len(gs_records)}")

        existing_keys = set()
        gs_data_map = {}
        if not gs_df.empty:
            gs_key_col = None
            for c in gs_df.columns:
                clean_c = c.replace('\n', '', '').replace(' ', '')
                if '台芝' in clean_c and '工作案號' in clean_c:
                    gs_key_col = c
                    break
            if not gs_key_col:
                gs_key_col = gs_df.columns[0]
            
            for row in gs_records:
                k = str(row.get(gs_key_col, '')).strip()
                if k:
                    existing_keys.add(k)
                    gs_data_map[k] = row

        print(f"📌 [DEBUG] Google Sheet 已存在的案號數量: {len(existing_keys)}")

        # 3. 比對並將缺少的紀錄 append 至 Google Sheet
        new_rows_to_append = []
        for _, im_row in filtered_im.iterrows():
            key_val = str(im_row.get(case_no_col, '')).strip()
            if not key_val or key_val in existing_keys:
                continue

            leave_time = str(im_row.get(time_col, '')).split('.')[0]
            store_id = str(im_row.get('門店編號', im_row.get('門市編號', ''))).replace('.0', '')
            store_name = str(im_row.get('門店名稱', im_row.get('門市名稱', '')))
            category = str(im_row.get('報修類別', ''))
            content = str(im_row.get(content_col, ''))

            new_row = [key_val, leave_time, store_id, store_name, category, content, "", "", "", ""]
            new_rows_to_append.append(new_row)
            
            gs_data_map[key_val] = {
                "離場時間": leave_time,
                "門店編號": store_id,
                "門店名稱": store_name,
                "報修類別": category,
                "工作內容": content,
                "SC(1)": "", "SC(2)": "", "TM(1)": "", "TM(2)": ""
            }

        if new_rows_to_append:
            print(f"📌 [DEBUG] 準備將以下新案號寫入 Google Sheet: {[r[0] for r in new_rows_to_append]}")
            sheet.append_rows(new_rows_to_append)
        else:
            print("📌 [DEBUG] 沒有需要新增至 Google Sheet 的新案號（可能都已存在）")

        # 4. 僅回傳最新前 5 筆
        result_list = []
        top_5 = filtered_im.head(5)
        print(f"📌 [DEBUG] 最終即將回傳的前 5 筆案號:")
        for _, im_row in top_5.iterrows():
            key_val = str(im_row.get(case_no_col, '')).strip()
            print(f"   - 案號: {key_val}, 內容: {im_row.get(content_col)}")
            item = gs_data_map.get(key_val, {})
            if item:
                result_list.append({
                    "case_no": key_val,
                    "leave_time": str(item.get('離場時間', im_row.get(time_col, ''))),
                    "store_id": str(item.get('門店編號', im_row.get('門店編號', ''))).replace('.0', ''),
                    "store_name": str(item.get('門店名稱', im_row.get('門市名稱', ''))),
                    "category": str(item.get('報修類別', im_row.get('報修類別', ''))),
                    "content": str(item.get('工作內容', im_row.get(content_col, ''))),
                    "sc1": str(item.get('SC(1)', '')),
                    "sc2": str(item.get('SC(2)', '')),
                    "tm1": str(item.get('TM(1)', '')),
                    "tm2": str(item.get('TM(2)', ''))
                })

        return result_list
    except Exception as e:
        print(f"❌ SC硬碟檢測同步失敗: {e}")
        import traceback
        traceback.print_exc()
        return []


@sc_check_bp.route('/sc_check/update', methods=['POST'])
def sc_check_update():
    """ 即時更新 SC(1), SC(2), TM(1), TM(2) 格子數值至 Google Sheet """
    data = request.json or {}
    case_no = data.get('case_no')
    field = data.get('field') # 'sc1', 'sc2', 'tm1', 'tm2'
    value = data.get('value', '')

    field_map = {'sc1': 'SC(1)', 'sc2': 'SC(2)', 'tm1': 'TM(1)', 'tm2': 'TM(2)'}
    col_name = field_map.get(field)
    
    if not case_no or not col_name:
        return jsonify({'status': 'error', 'message': '無效的參數'}), 400

    try:
        sheet = get_sc_check_sheet()
        all_rows = sheet.get_all_records()
        headers = sheet.row_values(1)
        
        target_row_idx = None
        target_col_idx = None
        key_col_idx = 1

        for idx, h in enumerate(headers, 1):
            clean_h = h.replace('\n', '').replace(' ', '')
            if '台芝' in clean_h and '工作案號' in clean_h:
                key_col_idx = idx
            if h.strip() == col_name:
                target_col_idx = idx

        if not target_col_idx:
            return jsonify({'status': 'error', 'message': f'找不到欄位 {col_name}'}), 400

        for i, row in enumerate(all_rows, start=2):
            k = str(row.get(headers[key_col_idx - 1], '')).strip()
            if k == str(case_no).strip():
                target_row_idx = i
                break

        if target_row_idx:
            sheet.update_cell(target_row_idx, target_col_idx, str(value))
            return jsonify({'status': 'success'})
        else:
            return jsonify({'status': 'error', 'message': '找不到對應案號資料'}), 404

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500


@sc_check_bp.route('/sc_check/delete', methods=['POST'])
def sc_check_delete():
    """ 同步刪除 Google Sheet 中的該筆資料 """
    data = request.json or {}
    case_no = data.get('case_no')

    if not case_no:
        return jsonify({'status': 'error', 'message': '缺少案號'}), 400

    try:
        sheet = get_sc_check_sheet()
        all_rows = sheet.get_all_records()
        headers = sheet.row_values(1)
        
        key_col_idx = 0
        for idx, h in enumerate(headers):
            clean_h = h.replace('\n', '').replace(' ', '')
            if '台芝' in clean_h and '工作案號' in clean_h:
                key_col_idx = idx
                break

        for i, row in enumerate(all_rows, start=2):
            k = str(row.get(headers[key_col_idx], '')).strip()
            if k == str(case_no).strip():
                sheet.delete_rows(i)
                return jsonify({'status': 'success'})

        return jsonify({'status': 'error', 'message': '找不到該筆資料可刪除'}), 404

    except Exception as e:
        return jsonify({'status': 'error', 'message': str(e)}), 500