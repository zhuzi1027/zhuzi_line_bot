# 使用xlsx檔作為儲存
import json
import datetime
import re
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage
from main.save_data import save
from main.load_data import load


app = Flask(__name__)
# 使用json檔除存重要資訊
with open('settings.json', 'r', encoding='UTF-8') as jfile:
    jdata = json.load(jfile)

Line_Bot_Api = LineBotApi(jdata['token'])
handler = WebhookHandler(jdata['acces'])

action = ""
day = datetime.datetime.today().strftime

@app.route("/callback", methods=['POST'])
def callback():
    signture = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signture)
    except InvalidSignatureError:
        abort(400)
    return 'OK'

@handler.add(MessageEvent, message=TextMessage)
def hamlle_message(event):
    global action
    msg = event.message.text

    regex = re.compile("\[(.+)\]")
    match = regex.match(msg)

    if match:
        action = match.group(1)
        reply_msg = getActionReplyMsg()

        Line_Bot_Api.reply_message(
            event.reply_token,
            TextMessage(text=reply_msg)
        )
        return
    else:
        if action == '收入':
            
            reply_msg = (
                day('%m%d') + '收入︰' + msg
                )
            Line_Bot_Api.reply_message(
                event.reply_token,
                TextMessage(text=reply_msg)
            )
            save.day_data(msg, action)
            action = ""

        elif action == '支出':
            reply_msg = (
                day('%m%d') + '支出︰' + msg
            )
            Line_Bot_Api.reply_message(
                event.reply_token,
                TextMessage(text=reply_msg)
            )
            save.day_data(msg, action)
            action = ""

        elif action == '月收入':
            if int(msg) not in [
                1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
                ]:
                reply_msg = (
                    '輸入錯誤'
                )
            else:
                if load.check_data(msg):
                    if load.load_data(msg, action) != False:
                        month_msg = day('%m') + '月收入︰' + load.load_data(msg, action)
                    else:
                        month_msg = f'本月尚未有收入紀錄'
                else:
                    month_msg = '查無該月資料'

                reply_msg = (
                    month_msg
                )
                    
            Line_Bot_Api.reply_message(
                event.reply_token,
                TextMessage(text=reply_msg)
            )
            action = ""

        elif action == '月支出':
            if int(msg) not in [
                1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12
                ]:
                reply_msg = (
                    '輸入錯誤'
                )
            else:
                if load.check_data(msg):
                    if load.load_data(msg, action) != False:
                        month_msg = day('%m') + '月支出︰' + load.load_data(msg, action)
                    else:
                        month_msg = f'本月尚未有支出紀錄'
                else:
                    month_msg = '查無該月資料'

                reply_msg = (
                    month_msg
                )
                    
            Line_Bot_Api.reply_message(
                event.reply_token,
                TextMessage(text=reply_msg)
            )
            action = ""
        
        elif action == '年收入':
            if len(msg) != 2:
                reply_msg = (
                    '輸入錯誤'
                )
            else:
                reply_msg = '年收入︰' + load.load_year_data(msg, action)[0] + '\n年收支︰' + load.load_year_data(msg, action)[1]
            
            Line_Bot_Api.reply_message(
                event.reply_token,
                TextMessage(text=reply_msg)
            )
            action = ""

        elif action == '年支出':
            if len(msg) != 2:
                reply_msg = (
                    '輸入錯誤'
                )
            else:
                reply_msg = '年支出︰' + load.load_year_data(msg, action)[0] + '\n年收支︰' + load.load_year_data(msg, action)[1]
            
            Line_Bot_Api.reply_message(
                event.reply_token,
                TextMessage(text=reply_msg)
            )
            action = ""
        
def getActionReplyMsg():
    global action
    if action == '收入':
        return day('%m%d') + '-收入︰'
    elif action == '支出':
        return day('%m%d') + '-支出︰'
    elif action == '月收入':
        return '請輸入要查詢的月份︰'
    elif action == '月支出':
        return '請輸入要查詢的月份︰'
    elif action == '年收入':
        return '請輸入要查詢的年份\n(2023請輸入23)︰'
    elif action == '年支出':
        return '請輸入要查詢的年份\n(2023請輸入23)︰'
    else:
        return '無效指令'

if __name__ == '__main__':
    app.run()