import os
import json
import gspread
from google.oauth2.service_account import Credentials

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

# 4. 您的 API 抓取與寫入邏輯
# worksheet.append_row([data1, data2])
worksheet.append_row(['data1', 'data2'])
