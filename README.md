# line_bot
建立一個line的記帳機器人
## 目的

學習使用line官方帳號製作一個可以記帳並寫入Excel表儲存於本地端
| 模組 | 使用介紹 |
|:---:|:---:|
| json | 用以存放需要使用的重要數值 |
| flask | python的輕架構網站模組 |
| line-bot-sdk | line機器人的運作 |
| re | 判斷字串內容，用以判別指令 |
| datetime | 用以載入系統時間 |
因為line的官方帳號需要連動https://開頭的網站
所以需要額外使用ngrok這個程式
這是一個可以將你所架設的網站位置http://127.0.0.1:port 轉換為https://的網站
## 內容

app.py
  - 主要運行程式內容

settings.json
  - 用以存放token, acces的數據

main
  - load_data.py
    1. check_data : 檢查檔案、活頁、當天日期的收支是否存在
    2. load_data : 用以載入月收入/支出的值，如果沒有則回傳False
    3. load_year_data : 用以載入年收入/支出的值與年收支的值
  - save_data.py
    1. day_data : 用以儲存輸入的收入/支出檔案，並建立*年分.xlsx* , *月份的活頁簿*, *當天日期的收支*

data
  - 23.xlsx
