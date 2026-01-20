import os
import json
import gspread
from google.oauth2.service_account import Credentials
# Import necessary libraries
import requests
from datetime import datetime
import time
# 引入 google.auth 模組
import google.auth
from google.auth.transport.requests import AuthorizedSession
# 由於您在 Section 3 嘗試使用 open/create 邏輯，需要引入 gspread.exceptions
import gspread.exceptions

# 1. 定義權限範圍
scopes = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# 2. 讀取環境變數 (GitHub Actions 會用到)
service_account_info = json.loads(os.environ["GCP_SERVICE_ACCOUNT_KEY"])
creds = Credentials.from_service_account_info(service_account_info, scopes=scopes)

# 3. 認證並打開試算表
client = gspread.authorize(creds)
spreadsheet = client.open("Test4Auto")
worksheet = spreadsheet.get_worksheet(0)

#######################################################################################################

# Define constants
ACCESS_TOKEN = os.environ["IG_ACCESS_TOKEN"]
MEDIA_ID = ''
API_URL = f'https://graph.instagram.com/v24.0{MEDIA_ID}/me?'
# Define parameters for the API call
params = {
    'fields': 'user_id,username,name,account_type,profile_picture_url,followers_count,follows_count,media_count',
    'access_token': ACCESS_TOKEN
}
# Make the API call
response = requests.get(API_URL, params=params)
user_id = ""
username = ""
id = ""
if response.status_code == 200:
    data = response.json()
    #print('Reel View Count:', data.get('view_count', 'N/A'))
    print(data)
    user_id = data.get('user_id', 'N/A')
    username = data.get('username', 'N/A')
    id = data.get('id', 'N/A')
    followers_count = data.get('followers_count', 'N/A')
    follows_count = data.get('follows_count', 'N/A')
    media_count = data.get('media_count', 'N/A')
else:
    print('Error:', response.status_code, response.text)

#######################################################################################################

def fetch_all_instagram_media_ids(user_id: str, access_token: str, api_version: str = 'v24.0') -> list:
    """
    使用 Instagram Graph API 遞迴抓取指定用戶的所有媒體素材 ID。

    Args:
        user_id (str): Instagram 用戶 ID (IG User ID)。
        access_token (str): 應用程式的存取權杖。
        api_version (str): Graph API 版本。

    Returns:
        list: 包含所有媒體素材 ID 的列表，如果發生錯誤則返回空列表。
    """

    # 設置 API 基礎 URL，使用 f-string 讓變數更清晰
    API_URL = f'https://graph.instagram.com/{api_version}/{user_id}/media'

    # 設置初始請求參數
    params = {
        'access_token': access_token,
        'limit': 100,  # 建議將單次請求上限設為 100，減少 API 呼叫次數
        # 這裡可以加入 fields 參數，例如 'fields': 'id,caption,media_type'
    }

    all_media_ids = []

    print(f"--- 開始抓取 IG User ID: {user_id} 的媒體素材 ---")

    while True:
        try:
            # 發出 API 請求
            response = requests.get(API_URL, params=params)

            # 檢查 HTTP 狀態碼
            if response.status_code != 200:
                print(f"Error: API 請求失敗，狀態碼 {response.status_code}")
                print(f"Error 訊息: {response.text}")
                break

            data = response.json()

            # 檢查 API 回應中是否有錯誤訊息 (例如權限不足)
            if 'error' in data:
                print(f"Error: API 回傳錯誤訊息: {data['error']}")
                break

            # 提取當前頁面的媒體 ID
            current_page_data = data.get('data', [])
            if not current_page_data:
                print("已抓取所有資料，或當前頁面資料為空。")
                break # 數據為空，結束迴圈

            for media in current_page_data:
                all_media_ids.append(media.get('id'))

            print(f"已抓取 {len(current_page_data)} 筆資料，總數: {len(all_media_ids)}")

            # 獲取下一頁的分頁指標 (cursors)
            cursors = data.get('paging', {}).get('cursors', {})
            after_cursor = cursors.get('after')

            # 如果沒有 'after' 游標，表示已經到達最後一頁
            if not after_cursor:
                print("到達最後一頁，抓取完成。")
                break

            # 更新參數，準備下一輪迴圈
            params['after'] = after_cursor

        except requests.exceptions.RequestException as e:
            print(f"Error: 發生網路請求錯誤: {e}")
            break
        except Exception as e:
            print(f"Error: 發生未知錯誤: {e}")
            break

    return all_media_ids

# --- 範例執行 ---
# 假設 id 變數已定義
# id = 'YOUR_INSTAGRAM_USER_ID'

# 請確保您的 ACCESS_TOKEN 和 ID 是正確且有效的
# 建議將您的 ID 放在程式碼頂部方便修改
INSTAGRAM_USER_ID = id
PASS_LIST = ['17995655195297716','18027606775595308','17993740604152764','17991807557282567','18292447336139034','17987986763500515','17983532429104920','17929002233721037','18042008242491057','18373890745022698','17994633335037107','17987612747281216','17976303950169941','17999997202918338','17849543381979617','17919266669654461','17978995856471488','18289872490111521','17986932202898407','17982428393156323','17986942325031875','18016956349587748','17956049732241832','17969350279910431','17953015732973936','18030493229751031','18064145252176676','18078394606966961','18397646335187210','17950944584882200','18087848413771569','17935945419025070','18064473098007781','17967244265824745','17984845673780540','17947898051905080','18025167386579575','18060341791695327','18041985898938665','18006684671531680','18035091601762434','18017482958031935','17914205837833869','18013174528920365','18182670862304964','18238957441172863','18236751793225579','18113562469320862','18060994348429330','18076860016383981','18026032189545217','18024837205538174','18007425304843416','18088928590347483','18004107316713356','18247055365167730','17987422222792512','18001620391621422','17893586645769378','18348560860000143','17937196775525731','17964515305994006','17971832623842386','18326388061007730','17998835098484054','17981198962680900','17975611672699611','17962215529993590','17947853309207898','17946600857106907','17895223502646239','17947640530963377','17936872322301359','17926154570157178','17943261994976302','17917754207398890','17879242364590359','18211145152186820','17924704319030587','18071439811295678','17939533846672417','17871212336611830','17870257163606324','17897183249250376','17934256144687890','17946340093862755','18186259429126268','18131977678215353','17900509904098512','17931425641644119','17947479892482957','17989460845368154','17866233446409746','17918792104578517','17858841401391380','17897554564661224','17856090242321476','18154084189048682','17998417594292466','17877160432657805','17872108594575599','17880658909505766','17858139883608171','18071082451181834','17897602831390344','18065102497187547','18048070408163201','17988218749259846','18074257537042067','18023313850166724','18045964561050479','17960796733243557','17857104145353616','17964240523030954','17917361305194908','17885653975236909','17895869335220615','17913943555158914','17865078949227710','17868823372160001','17852937691173947','17852181505112824','17851824448119260','17859207214041811','17849445586080879','17859102937019054','17849252173080673','17858749948028534','17855179777012011','17846421706097739','17846259700119538','17843686969021452','17843679622021452','17842182625021452','17842128241021452','17843613595021452','17843612479021452','17843610649021452','17843608468021452','17843607472021452','17852937691173947']

if __name__ == '__main__':
    if INSTAGRAM_USER_ID == 25544294341841753:
        print("請先將 INSTAGRAM_USER_ID 替換為實際的 IG 用戶 ID。")
    else:
        all_ids = fetch_all_instagram_media_ids(INSTAGRAM_USER_ID, ACCESS_TOKEN)

        print("\n--- 最終結果 ---")
        print("所有媒體 ID 列表:", all_ids)
        print("總素材數量:", len(all_ids))
        print("要跳過的數量:", len(PASS_LIST))
        print("剩餘數量:", len(all_ids) - len(PASS_LIST))

        set_a = set(all_ids)
        set_b = set(PASS_LIST)
        set_c = set_a - set_b
        detail_ids = list(set_c)
        print("明細Daily數量:", len(detail_ids))

#######################################################################################################

# Google 試算表檔案名稱
SHEET_TITLE = 'Instagram_EmilyNTU_v2_Head'

# 必須填入您資料夾的 ID 才能將文件移動到指定位置！
TARGET_FOLDER_ID = 'YOUR_FOLDER_ID_HERE' # 請替換為您的 Folder ID

# 試算表標題列 (Sheet 1 - 媒體資料)
HEADERS_MEDIA = ['media_id','write_date','comments_count','is_shared_to_feed','like_count','media_type',
          'media_url','permalink','shortcode','timestamp']

# 試算表標題列 (Sheet 2 - 用戶資料 Log)
WORKSHEET_TITLE_2 = 'User_Profile_Log'
HEADERS_PROFILE = ['write_date', 'user_id', 'username', 'name', 'account_type',
                   'profile_picture_url', 'followers_count', 'follows_count', 'media_count']


# 初始化變數
gc = None
worksheet_media = None # Sheet 1: 媒體資料工作表
worksheet_profile = None # Sheet 2: 用戶資料工作表
existing_data = []
media_id_to_row_index = {}


# --- 2. Google Colab 認證與 gspread 設定 ---
print("已完成 Google Colab 認證...")

# --- 3. 打開或創建 Google 試算表並設定工作表 ---
try:
    # 1. 嘗試透過名稱打開文件
    try:
        spreadsheet = client.open(SHEET_TITLE)
        print(f"已找到現有試算表: {SHEET_TITLE}")
    except gspread.exceptions.SpreadsheetNotFound:
        # 2. 如果文件不存在，則創建新文件
        spreadsheet = client.create(SHEET_TITLE)
        print(f"創建新試算表: {SHEET_TITLE}")

        # 3. 移動文件 (如果 TARGET_FOLDER_ID 已定義)
        if TARGET_FOLDER_ID and TARGET_FOLDER_ID != 'YOUR_FOLDER_ID_HERE':
            file_id = spreadsheet.id
            file = client.client.drive_api.get(fileId=file_id, fields='parents').execute()
            # 移動檔案
            client.client.drive_api.update(
                fileId=file_id,
                removeParents=','.join(file['parents']),
                addParents=TARGET_FOLDER_ID
            ).execute()
            print(f"試算表 '{SHEET_TITLE}' 已成功移動到指定資料夾。")
        else:
            print("未提供有效的 TARGET_FOLDER_ID，檔案保留在根目錄。")


    # --- 處理第一個工作表 (Sheet 1 - 媒體資料) ---
    # 選擇第一個工作表 (預設為 Sheet 1)
    worksheet_media = spreadsheet.sheet1
    # 建議給它一個名稱以保持清晰
    if worksheet_media.title == 'Sheet1':
          worksheet_media.update_title('Media_Data')

    # 檢查並設定標題列 (Sheet 1)
    if not worksheet_media.row_values(1):
        worksheet_media.append_row(HEADERS_MEDIA)
        print("已設定媒體資料試算表標題列。")
    else:
        print("媒體資料試算表標題列已存在，將讀取現有資料進行比對。")

    # 讀取 Sheet 1 所有資料，以便進行比對和更新
    existing_data = worksheet_media.get_all_records()
    media_id_to_row_index = {
        row.get('media_id'): index + 2
        for index, row in enumerate(existing_data) if row.get('media_id')
    }

    # --- 處理第二個工作表 (Sheet 2 - 用戶資料 Log) ---
    try:
        worksheet_profile = spreadsheet.worksheet(WORKSHEET_TITLE_2)
        print(f"已找到現有工作表: {WORKSHEET_TITLE_2}")
    except gspread.WorksheetNotFound:
        # 如果工作表不存在，則創建它
        worksheet_profile = spreadsheet.add_worksheet(title=WORKSHEET_TITLE_2, rows="100", cols="20")
        print(f"已創建新工作表: {WORKSHEET_TITLE_2}")

    # 檢查並設定標題列 (Sheet 2)
    if not worksheet_profile.row_values(1):
        worksheet_profile.append_row(HEADERS_PROFILE)
        print(f"已設定工作表 '{WORKSHEET_TITLE_2}' 標題列。")
    else:
        print(f"工作表 '{WORKSHEET_TITLE_2}' 標題列已存在，將直接新增資料。")


except Exception as e:
    print(f"處理試算表或移動檔案時發生錯誤: {e}")
    worksheet_media = None
    worksheet_profile = None
    existing_data = []
    media_id_to_row_index = {}


# --- 4. 迴圈處理 Instagram 媒體資料並寫入/更新試算表 (Sheet 1 - 更新/新增 邏輯) ---
print("\n開始查詢 Instagram 媒體資料並更新/新增寫入 Sheet 1...")

if worksheet_media is None:
    print("⚠️ 因 Sheet 1 無法連接，跳過媒體資料寫入步驟。")
else:
    STANDARD_INSIGHT_METRICS = 'reach,saved,shares,total_interactions,views' # 不含 Reels 的指標
    all_ids = all_ids[:10]

    for media_id in all_ids:
        # --- A. 呼叫 Media API (獲取基礎數據) ---
        MEDIA_URL = f'https://graph.instagram.com/v24.0/{media_id}/'
        media_params = {
            'fields': 'comments_count,is_shared_to_feed,like_count,media_type,media_url,permalink,shortcode,timestamp',
            'access_token': ACCESS_TOKEN
        }

        media_data = {}

        try:
            # 1. 取得 Media 資料
            time.sleep(0.5)
            response_media = requests.get(MEDIA_URL, params=media_params)
            response_media.raise_for_status()
            media_data = response_media.json()
            print(f"成功取得媒體 ID {media_id} 的基礎資料。")

            # --- 處理數據並準備寫入 ---
            write_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            # 準備要寫入試算表的一行資料 (必須與 HEADERS_MEDIA 順序一致)
            new_row_data = [
                media_id,
                write_date,
                media_data.get('comments_count', 0),
                media_data.get('is_shared_to_feed', False),
                media_data.get('like_count', 0),
                media_data.get('media_type', 'UNKNOWN'),
                media_data.get('media_url', ''),
                media_data.get('permalink', ''),
                media_data.get('shortcode', ''),
                media_data.get('timestamp', ''),
            ]

            # 嘗試查找現有行
            row_to_update = media_id_to_row_index.get(int(media_id))

            if row_to_update:
                # **更新邏輯**：檢查主要欄位 (comments_count 和 like_count)
                existing_row = existing_data[row_to_update - 2]

                current_comments = existing_row.get('comments_count', 0)
                current_likes = existing_row.get('like_count', 0)

                new_comments = new_row_data[HEADERS_MEDIA.index('comments_count')]
                new_likes = new_row_data[HEADERS_MEDIA.index('like_count')]

                # 轉換型別以確保正確比對
                try:
                    current_comments = int(current_comments)
                    current_likes = int(current_likes)
                except (ValueError, TypeError):
                    print(f"⚠️ 媒體 ID {media_id} 既有資料欄位非數字，強制更新。")
                    current_comments = -1 # 設定一個肯定會導致更新的值


                # 條件判斷：只有在留言或按讚數有變動時才更新 (避免頻繁寫入 API)
                if new_comments != current_comments or new_likes != current_likes:
                    # 執行更新：將整個新行數據寫入到既有的行數
                    worksheet_media.update(f'A{row_to_update}', [new_row_data])
                    print(f"媒體 ID {media_id} 資料已成功 **更新** (行: {row_to_update})。")
                # else:
                    # print(f"媒體 ID {media_id} 資料 **無變動**，跳過更新。")

            else:
                # **新增邏輯**：如果找不到 media_id，則新增一筆
                worksheet_media.append_row(new_row_data)
                print(f"媒體 ID {media_id} 合併資料已成功 **新增** 至 Sheet 1。")

        except requests.exceptions.HTTPError as http_err:
            error_message = f"HTTP 錯誤: {http_err} - {http_err.response.text}"
            print(f"處理媒體 ID {media_id} 發生錯誤: {error_message}")
        except requests.exceptions.RequestException as req_err:
            print(f"請求錯誤 (如連線問題): {req_err}")
        except Exception as e:
            print(f"發生未預期錯誤: {e}")

        # 在 API 調用之間加入延遲
        time.sleep(0.5)


# --- 5. 抓取 Instagram 用戶資料並新增至 Sheet 2 (User_Profile_Log) ---
print("\n開始查詢 Instagram 用戶資料並新增寫入 Sheet 2...")

if worksheet_profile is None:
    print("⚠️ 因 Sheet 2 無法連接，跳過用戶資料寫入步驟。")
else:
    # 修正 API URL，直接獲取當前用戶資料
    API_URL = 'https://graph.instagram.com/me'
    # Define parameters for the API call
    params = {
        'fields': 'user_id,username,name,account_type,profile_picture_url,followers_count,follows_count,media_count',
        'access_token': ACCESS_TOKEN
    }
    write_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    try:
        time.sleep(0.5)
        response = requests.get(API_URL, params=params)
        response.raise_for_status()
        data = response.json()
        print("成功取得用戶資料。")

        # 準備要寫入試算表的一行資料 (必須與 HEADERS_PROFILE 順序一致)
        new_row_data = [
            write_date,
            data.get('id', ''), # Instagram Graph API 返回 'id' 作為 User ID
            data.get('username', ''),
            data.get('name', ''),
            data.get('account_type', ''),
            data.get('profile_picture_url', ''),
            data.get('followers_count', 0),
            data.get('follows_count', 0),
            data.get('media_count', 0),
        ]

        # 寫入 (新增) 資料到 Sheet 2
        worksheet_profile.append_row(new_row_data)
        print(f"用戶資料已成功 **新增** 至工作表 '{WORKSHEET_TITLE_2}'。")

    except requests.exceptions.HTTPError as http_err:
        error_message = f"HTTP 錯誤: {http_err} - {http_err.response.text}"
        print(f"處理用戶資料發生錯誤: {error_message}")
    except requests.exceptions.RequestException as req_err:
        print(f"請求錯誤 (如連線問題): {req_err}")
    except Exception as e:
        print(f"發生未預期錯誤: {e}")

    time.sleep(0.5) # API 呼叫之間的延遲

print("\n所有資料處理完畢。請到您的 Google Drive 檢查試算表。")

#######################################################################################################

# --- 1. 設定常量 (Constants Setup) ---

# Google 試算表檔案名稱
SHEET_TITLE = 'Instagram_EmilyNTU_v2'
# WORKSHEET_TITLE_2 = 'Instagram_EmilyNTU_v1_Data'

# 必須填入您資料夾的 ID 才能將文件移動到指定位置！
TARGET_FOLDER_ID = 'YOUR_FOLDER_ID_HERE' # 請替換為您的 Folder ID

# 試算表標題列 (已新增所有 Insight API 指標)
HEADERS = ['media_id','write_date','comments_count','is_shared_to_feed','like_count','media_type',
           'media_url','permalink','shortcode','timestamp',
           # 新增的 Insight Metrics (保持欄位順序固定)
           'ig_reels_avg_watch_time','ig_reels_video_view_total_time',
           'reach','saved','shares','total_interactions','views']

# 初始化變數
gc = None
worksheet = None

# --- 2. Google Colab 認證與 gspread 設定 ---
print("正在執行 Google Colab 認證...")

# --- 3. 打開或創建 Google 試算表並移動到指定資料夾 (已修正) ---
try:
    # 1. 嘗試透過名稱打開文件
    try:
        spreadsheet = client.open(SHEET_TITLE)
        print(f"已找到現有試算表: {SHEET_TITLE}")
    except gspread.exceptions.SpreadsheetNotFound:
        # 2. 如果文件不存在，則創建新文件 (預設在根目錄)
        spreadsheet = client.create(SHEET_TITLE)
        print(f"創建新試算表: {SHEET_TITLE}")

        # 3. 如果 TARGET_FOLDER_ID 已定義，則移動文件
        if TARGET_FOLDER_ID and TARGET_FOLDER_ID != 'YOUR_FOLDER_ID_HERE':
            file_id = spreadsheet.id
            file = client.client.drive_api.get(fileId=file_id, fields='parents').execute()

            # 移動檔案
            client.client.drive_api.update(
                fileId=file_id,
                removeParents=','.join(file['parents']),
                addParents=TARGET_FOLDER_ID
            ).execute()

            print(f"試算表 '{SHEET_TITLE}' 已成功移動到指定資料夾。")
        else:
            print("未提供有效的 TARGET_FOLDER_ID，檔案保留在根目錄。")


    # 選擇第一個工作表 (Sheet 1)
    worksheet = spreadsheet.sheet1
    # --- 處理第二個工作表 (Sheet 2 - 用戶資料 Log) ---
    # try:
        # worksheet = spreadsheet.worksheet(WORKSHEET_TITLE_2)
        # print(f"已找到現有工作表: {WORKSHEET_TITLE_2}")
    # except gspread.WorksheetNotFound:
        # 如果工作表不存在，則創建它
        # worksheet = spreadsheet.add_worksheet(title=WORKSHEET_TITLE_2, rows="100", cols="20")
        # print(f"已創建新工作表: {WORKSHEET_TITLE_2}")

    # 檢查並設定標題列 (僅在工作表為空時寫入)
    if not worksheet.row_values(1):
        worksheet.append_row(HEADERS)
        print("已設定試算表標題列。")
    else:
        print("試算表標題列已存在，將直接新增資料。")

except Exception as e:
    print(f"處理試算表或移動檔案時發生錯誤: {e}")
    worksheet = None


# --- 4. 迴圈處理 Instagram API 請求並寫入試算表 (已更新) ---
print("\n開始查詢 Instagram 媒體資料並新增寫入試算表...")

# 用於處理 Insight 數據的輔助函式
# 調整：如果找不到值，返回空字串 ''，以便在試算表中留空
def get_insight_value(insights_data, metric_name):
    """從 Insight API 響應中提取特定指標的值，如果找不到則返回空字串"""
    for item in insights_data.get('data', []):
        if item.get('name') == metric_name and item.get('values'):
            return item['values'][0].get('value', 0)
    return '' # 返回空字串以在試算表中留空

if worksheet is None:
    print("⚠️ 因試算表無法連接，跳過資料寫入步驟。")
else:
    # 定義所有可能的 Insight 指標
    ALL_INSIGHT_METRICS = 'ig_reels_avg_watch_time,ig_reels_video_view_total_time,reach,saved,shares,total_interactions,views'
    STANDARD_INSIGHT_METRICS = 'reach,saved,shares,total_interactions,views' # 不含 Reels 的指標
    index = 1

    detail_ids = detail_ids[:10]

    for media_id in detail_ids:
        # --- A. 呼叫 Media API (獲取基礎數據) ---
        MEDIA_URL = f'https://graph.instagram.com/v24.0/{media_id}/'
        media_params = {
            'fields': 'comments_count,is_shared_to_feed,like_count,media_type,media_url,permalink,shortcode,timestamp',
            'access_token': ACCESS_TOKEN
        }

        INSIGHT_URL = f'https://graph.instagram.com/v24.0/{media_id}/insights'

        media_data = {}
        insight_data = {}
        insight_metrics = {}

        is_reels_metrics_failed = False
        current_insight_metrics_list = ALL_INSIGHT_METRICS # 預設嘗試所有指標

        try:
            # 1. 取得 Media 資料
            time.sleep(0.5)
            response_media = requests.get(MEDIA_URL, params=media_params)
            response_media.raise_for_status()
            media_data = response_media.json()
            # print(f"成功取得媒體 ID {media_id} 的基礎資料。")

            # 2. 嘗試取得 Insight 資料 (第一階段：包含 Reels 指標)
            try:
                insight_params = {
                    'metric': ALL_INSIGHT_METRICS,
                    'access_token': ACCESS_TOKEN
                }
                response_insight = requests.get(INSIGHT_URL, params=insight_params)
                response_insight.raise_for_status()
                insight_data = response_insight.json()
                # print(f"成功取得媒體 ID {media_id} 的所有 Insight 資料。")

            except requests.exceptions.HTTPError as insight_http_err:
                # 檢查是否為 'Unsupported feature' 或類似錯誤 (指標無效)
                if 'unsupported' in insight_http_err.response.text.lower() or 'metric' in insight_http_err.response.text.lower():
                    # print(f"媒體 ID {media_id} 的 Reels 指標無效，嘗試標準指標。")
                    is_reels_metrics_failed = True
                    current_insight_metrics_list = STANDARD_INSIGHT_METRICS # 設定為標準指標列表

                    # 2b. 嘗試取得 Insight 資料 (第二階段：不含 Reels 指標)
                    insight_params = {
                        'metric': STANDARD_INSIGHT_METRICS,
                        'access_token': ACCESS_TOKEN
                    }
                    response_insight = requests.get(INSIGHT_URL, params=insight_params)
                    response_insight.raise_for_status() # 這次如果失敗就直接拋出錯誤
                    insight_data = response_insight.json()
                    # print(f"成功取得媒體 ID {media_id} 的標準 Insight 資料。")

                else:
                    # 其他非指標相關的 HTTP 錯誤，例如 Token 過期
                    raise insight_http_err # 重新拋出錯誤

            # --- 處理數據並準備寫入 ---
            write_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

            # 提取 Insight 數據
            # 預設所有指標為空字串 (即試算表留空)
            insight_metrics = {
                'ig_reels_avg_watch_time': '',
                'ig_reels_video_view_total_time': '',
                'reach': '',
                'saved': '',
                'shares': '',
                'total_interactions': '',
                'views': '',
            }

            # 根據成功獲取的指標列表，更新實際值
            metrics_to_process = current_insight_metrics_list.split(',')
            for metric_name in metrics_to_process:
                value = get_insight_value(insight_data, metric_name.strip())
                # 如果成功取得值，則寫入，否則保持空字串 (已由 get_insight_value 處理)
                insight_metrics[metric_name.strip()] = value


            # 準備要寫入試算表的一行資料 (必須與 HEADERS 順序一致)
            row_data = [
                media_id,
                write_date,
                media_data.get('comments_count', 0),
                media_data.get('is_shared_to_feed', False),
                media_data.get('like_count', 0),
                media_data.get('media_type', 'UNKNOWN'),
                media_data.get('media_url', ''),
                media_data.get('permalink', ''),
                media_data.get('shortcode', ''),
                media_data.get('timestamp', ''),
                # Insight 數據部分 (按 HEADERS 順序固定填入)
                insight_metrics['ig_reels_avg_watch_time'],
                insight_metrics['ig_reels_video_view_total_time'],
                insight_metrics['reach'],
                insight_metrics['saved'],
                insight_metrics['shares'],
                insight_metrics['total_interactions'],
                insight_metrics['views'],
            ]

            # 寫入 (新增) 資料到試算表
            worksheet.append_row(row_data)
            print(f"{index}:媒體 ID {media_id} 合併資料已成功新增至試算表。")
            index += 1
            if(index==180):
              time.sleep(5)

        except requests.exceptions.HTTPError as http_err:
            # 處理 Media API 或第二次 Insight API 的錯誤
            error_message = f"HTTP 錯誤: {http_err} - {http_err.response.text}"
            print(f"處理媒體 ID {media_id} 發生錯誤: {error_message}")
        except requests.exceptions.RequestException as req_err:
            print(f"請求錯誤 (如連線問題): {req_err}")
        except Exception as e:
            print(f"發生未預期錯誤: {e}")

        # 在 API 調用之間加入延遲
        time.sleep(0.5)

    print("\n所有媒體資料處理完畢。請到您的 Google Drive 檢查試算表。")
