# chip_analysis

現階段有掛自動化在TEJ公司內部電腦，交易日的18:30啟動程式後抓取資料，整理後發送信件

**檔案圖片**
![image](https://github.com/wuzheyu913/chip_analysis/assets/71300574/fe68f881-a307-4f8d-936a-e45c2cd5b5ef)

## 需要檔案說明
**籌碼分點統整通知_from_API.ipynb :**
使用TEJ API 來抓每日盤後的籌碼資料

**sotck-chip-analysis-4d71166d853b.json :** 
是用來抓取 google sheet 資料 (https://docs.google.com/spreadsheets/d/1OeCiGA9Bp6TdOjrBW7R8XAvneaOdb53InlmiXnnKrEA/edit#gid=0)

**mail_receive.xlsx :** 
收件名單 (在此新增/刪減收件人)

**公司所在縣市.xlsx :**
用來標出地緣



# 更新
## 2024/04/19 更新
1. 將買超金額沒有按順序的bug修掉
2. 新增"可能是地緣"的欄位 : 透過拿分點的地址&公司的地址去比對，如果同縣市，則標上Y


