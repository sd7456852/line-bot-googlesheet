import os
import requests
import threading
import pygsheets
from flask import Flask, abort, request
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import (MessageEvent, TextMessage, TextSendMessage)

gc = pygsheets.authorize(service_account_file='linebot-add-660150b76b49.json')

# Google Sheets 授權並打開 Google Sheet
gc = pygsheets.authorize(service_file='linebot-add-660150b76b49.json')
sh = gc.open_by_key('1_V0OCyZMK6ryTTf43IwGhdo1fIHPoeyTrBSCIROrBuY')  
wks = sh[0]  # 選擇第一個工作表

app = Flask(__name__)

# 用replit要改成下面的方法
token = os.environ['CHANNEL_ACCESS_TOKEN']
secret = os.environ['CHANNEL_SECRET']
line_bot_api = LineBotApi(token)
handler = WebhookHandler(secret)

# 创建锁对象
lock = threading.Lock()

# ////////////////////////////////////////////////////////////////
def input_data(user, name, date_input, amount_input):
    start_col, total_cell = (1, 'D2') if user == '1' else (5, 'H2')

    # 读取当前的总金额
    try:
        total_amount = float(wks.get_value(total_cell))
    except ValueError:
        total_amount = 0.0

    # 找到下一个可用的行
    next_row = len(wks.get_col(start_col, include_empty=False)) + 1

    # 将数据添加到 Google Sheet
    wks.update_row(next_row, [name, date_input, amount_input], col_offset=start_col - 1)

    # 更新总金额
    total_amount += float(amount_input)

    # 将总金额更新到指定单元格
    wks.update_value(total_cell, total_amount)

    next_row += 1

    return f"資料已添加: \n{name},{date_input},{amount_input}"



# ////////////////////////////////////////////////////////////////
# 定義一個函數來刪除數據
def delete_data(user):
    start_col, end_col = (1, 4) if user == '1' else (5, 8)
    # 找到所有數據的行數
    rows = len(wks.get_col(start_col, include_empty=False))
    if rows > 1:
        # 刪除指定範圍內的數據，從第二行開始
        for row in range(2, rows + 1):
            wks.update_row(row, [''] * (end_col - start_col + 1), col_offset=start_col - 1)
        return f"已刪除{user}的所有數據。"
    else:
        return "沒有可刪除的數據。"

# 定義一個函數來刪除 除了第一行之外的所有数据
def delete_all_data():

    header = wks.get_row(1)

    # 清空所有数据
    wks.clear()

    # 重新设置第一行的标题
    wks.update_row(1, header)

    return "已删除所有数据。"

# 定義一個函數來讀取數據
def read_data(user):
    start_col, end_col = (1, 4) if user == '1' else (5, 8)
    data = wks.get_values(start=(1, start_col), end=(wks.rows, end_col), include_tailing_empty_rows=False)
    result = "\n".join(["\t".join(row) for row in data])
    return result
# ---------------------------------------------

# 監聽所有來自 /callback 的 Post Request
@app.route("/callback", methods=['POST'])
def callback():
    # get X-Line-Signature header value
    signature = request.headers['X-Line-Signature']

    # get request body as text
    body = request.get_data(as_text=True)
    app.logger.info("Request body: " + body)

    # handle webhook body
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)

    return 'OK'


@handler.add(MessageEvent, message=TextMessage)
def handle_message(event):
    lock.acquire()
    text = event.message.text

    if text.startswith('輸入'):
        try:
            _, user, name, date_input, amount_input = text.split(' ')
            result = input_data(user.strip(), name.strip(), date_input.strip(), amount_input.strip())
        except ValueError:
            result = "輸入格式錯誤。請使用 '輸入 用戶 姓名 日期 金額' 的格式。"
            
    elif text.startswith('刪除'):
        try:
            _, user = text.split(' ')
            if user.strip() == '全部':
                result = delete_all_data()
            else:
                result = delete_data(user.strip())
        except ValueError:
            result = "刪除格式錯誤。請使用 '刪除,用戶' 或 '刪除,全部' 的格式。"
            
    elif text.startswith('讀取1'):
        
        result = read_data('1')
        
    elif text.startswith('讀取2'):
        
        result = read_data('2')
        
    elif text.startswith('全部讀取'):
        
        result = read_data('1') + '\n' + read_data('2')
        
    else:
        
        result = "無效的操作。請使用 '輸入 用戶 姓名 日期 金額'，'刪除 用戶'，'刪除 全部'，'讀取1'，'讀取2' 或 '全部讀取'。"
        
        
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=result))
    lock.release()
if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)

