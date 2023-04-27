import os
import time
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build
import telegram
import asyncio
import datetime
import time

# Replace with the ID of the Google Sheet that you want to monitor
SHEET_ID = '1OThhPsNi2wgh1JhRYPp6-lYYS_W2z1iBHJYugAmc9mg'
# SHEET_ID = "1iM8t9OE_-eUiI0dCbS2Rv3d74L4hAyfjr-0uGTJ7z1U"

# Replace with the range of the sheet that you want to monitor
SHEET_RANGE = '5-Room!A1:ZZ'

#Tele API
load_dotenv()
BOT_TOKEN= os.getenv("TELE_BOT_API")

def authenticate():
    """Authenticate with the Google Sheets and Drive APIs using a service account."""
    SCOPES = ['https://www.googleapis.com/auth/drive.readonly', 'https://www.googleapis.com/auth/spreadsheets.readonly']
    SERVICE_ACCOUNT_FILE = 'rare-guide-384806-d91cbdf22e9f.json'
    credentials = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    sheets_service = build('sheets', 'v4', credentials=credentials)
    result = sheets_service.spreadsheets().get(
            spreadsheetId=SHEET_ID,
            includeGridData=True,
            ranges=[SHEET_RANGE]).execute()
    sheets = result.get('sheets', [])
    data = sheets[0].get('data', [])
    row_data = data[0].get('rowData', [])
    return result, row_data

def count_row_col(result):
    sheets = result.get('sheets', [])
    data = sheets[0].get('data', [])
    row_data = data[0].get('rowData', [])

    row_length = len(row_data)
    col_length = 0
    for row in row_data:
        temp = row.get('values', [])
        if(col_length < len(temp)):
            col_length = len(temp)
    return row_length, col_length

def init(row_data, row_length, col_length):
    empty_dict = {}
    array = [[empty_dict for i in range(col_length)] for j in range(row_length)]

    for i in range(len(row_data)):
        values = row_data[i].get('values', [])
        for j in range(len(values)):
            style = values[j].get('userEnteredFormat', {})
            array[i][j] = style
    return array

def bot_send_message(bot, channel_name, message):
    try:
        async def send_message():
            await bot.send_message(chat_id=channel_name, text=message)

        loop = asyncio.get_event_loop()
        loop.run_until_complete(send_message())
        return
    except:
        print("error")

def int_to_col_name(n):
    """
    Converts an integer to a Google Sheets column name (e.g. 1 -> 'A', 27 -> 'AA').
    """
    col_name = ""
    while n > 0:
        # find the remainder of n / 26, and convert it to a letter
        # A=0, B=1, ..., Z=25
        remainder = (n - 1) % 26
        letter = chr(65 + remainder)
        col_name = letter + col_name
        # subtract the remainder from n and divide by 26 to get the next digit
        n = (n - 1) // 26
    return col_name

def create_gsheet_url(row, col):
    cell_range = col + str(row)
    return f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit#gid=0&range={cell_range}"

def check_time():
    now = datetime.datetime.now()
    if now.hour == 0 and now.minute == 0:
        return True
    else:
        return False
def main():
    count = 0
    #Setting up telegram bot
    bot = telegram.Bot(token=BOT_TOKEN)
    channel_name = "@TengahGWF5"

    result, row_data = authenticate()
    row_length, col_length = count_row_col(result)
    previous_style = init(row_data, row_length, col_length)
    while True:
        result,row_data = authenticate()
        row_length, col_length = count_row_col(result)

        #Send message on how many selected today
        if check_time():
            message = "Total unit chosen today: " + str(count)
            bot_send_message(bot, channel_name, message)
            count = 0

        for i in range(len(row_data)):
            values = row_data[i].get('values', [])
            for j in range(len(values)):
                style = values[j].get('userEnteredFormat', {})
                if(style != previous_style[i][j]):
                    try:
                        if(style.get('textFormat').get('strikethrough') == True and ("textFormat" in previous_style[i][j]) == False):
                            previous_style[i][j] = style
                            col_name = int_to_col_name(j+1)
                            message = create_gsheet_url(str(i+1), col_name)
                            bot_send_message(bot, channel_name, message)
                            count += 1
                    except AttributeError:
                        continue


        time.sleep(300)



if __name__ == '__main__':
    main()
