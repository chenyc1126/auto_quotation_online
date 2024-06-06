---
title: '自動生成報價單程式'
disqus: hackmd
---

###### tags: `Loopick`

自動生成報價單程式
===
## 檔案結構
```bash
.
├── auto_quotation.py
├── message_to_json.py
├── output
├── quotation.json
├── requirement.txt
└── src
    ├── default.docx
    └── quotation.txt
```
## 程式介紹
本程式會從``` src/quotation.txt ```內讀取資料 (請按照提供之格式並注意不要有多餘的空白或換行，沒寫防呆QQ)，接著會自動生成 ```output/{報價單號}_{抬頭}``` 之 ```.doc ```檔案和 ```.json ```檔案
## quotation.txt 之固定格式
傳給消費者之內容格式，以訊息方式傳送
```
抬頭：（有報帳需求才需填寫）
統編：（有報帳需求才需填寫）
聯絡人：
電話：
信箱：
活動日期：
地址：
--------分隔線---------
餐廳名稱：
日期：
送達時間：
回收時間：
訂餐內容：
--------分隔線---------
```
拿到消費者回傳的訊息後，請手動新增前四行，完成後之 quotation.txt 如下，類別內容填（訂餐/租餐具）。
```
報價單號：B108
報價日期：8/18
運費：0
類別：訂餐 
抬頭：國際扶輪3522地區
統編：77017752
聯絡人：曾怡晨
電話：0928-018-986
信箱：rid3522.2223rac@gmail.com
活動日期：8/21
地址：北市中正區許昌街42號8樓（萊客 北車館）
--------分隔線---------
餐廳名稱：關山便當
日期：08/21
送達時間：11:45
回收時間：14:00
滷肉便當：*110 80
素食便當：*2 80
--------分隔線---------

```
## 用法
```bash
pip3 install -r requirement.txt
py auto_quotation.py #windows
python3 auto_quotation.py #linux or mac

```

