# Python樂天抓排行榜(依評價數，有分類表的)
# 使用軟體 https://www.jetbrains.com/pycharm/

import requests
import pandas as pd
import json
import datetime
from pandas import json_normalize
from tkinter import *

def rakuten_all():
    print("◉◉大分類◉◉")
    print("CD・DVD・楽器 (101240)")
    print("インテリア・寝具 (100804)")
    print("子供部屋インテリア・寝具・グッズ(203024)")
    print("おもちゃ・ホビー・ゲーム (101164)")
    print("キッチン・日用品雑貨・文具 (215783)")
    print("キッズ・ベビー・マタニティ (100533)")
    print("ジュエリー・腕時計 (216129)")
    print("スポーツ・アウトドア (101070)")
    print("ダイエット・健康・介護 (100938)")
    print("ドリンク・お酒 (100316)")
    print("パソコン・周辺機器 (100026)")
    print("バッグ・小物・ブランド雑貨 (216131)")
    print("花・園芸・エクステリア (100005)")
    print("ペット・ペットグッズ (101213)")
    print("家電・AV・カメラ (211742)")

# urlbase貼上想查的網站
urlbase = 'https://app.rakuten.co.jp/services/api/IchibaItem/Ranking/20170628?'
# 貼上申請好的API↓
appid = '1007613883261064353'  # 申請到的 アプリID
affid = '1fb05eb3.8afa909c.1fb05eb4.ac09c6ff'
# アフィリエイトID
rakuten_all()
genreid_input = input("請輸入編號")

# ↑編號ID(例：酒100316、居家100804）

genreid = genreid_input
period = 'realtime'

# ↑ランキング集計期間

url = urlbase + 'applicationId=' + appid + '&affiliateId=' + affid + '&genreId=' + genreid + '&period=' + period

# jsonデータの取得
r = requests.get(url)
jsondata = r.json()
# 出力(jsonデータの整形+日本語文字化け回避)
print(json.dumps(jsondata, indent=2, ensure_ascii=False))
# jsondata内のItems(商品情報)にアクセスした後に、データフレームに格納
df = json_normalize(jsondata['Items'])
# 上位5つだけ確認
df.head()
# 必要的情報抽出
df_pickup = df.loc[:, \
            ['Item.rank', \
             'Item.itemName', \
             # 'Item.catchcopy', \
             'Item.affiliateUrl', \
             'Item.reviewCount', \
             # 'Item.affiliateRate', \
             # 'Item.itemCaption', \
             'Item.itemPrice', \
             'Item.reviewAverage']]

# 項目の日本語辞書作成
rename_dic = \
    {'Item.rank': '排行榜名次', \
     'Item.itemName': '商品名', \
     # 'Item.catchcopy': 'キャッチコピー', \    reviewCount
     'Item.affiliateUrl': '網址URL', \
     'Item.reviewCount': '評價數', \
     # 'Item.affiliateRate': 'アフィリエイト利用率', \
     # 'Item.itemCaption': '商品説明', \
     'Item.itemPrice': '価格', \
     'Item.reviewAverage': '評價'}
# 項目名変更
df_pick_rename = df_pickup.rename(columns=rename_dic)

# データフレーム→CSVに出力 (輸出資料想存檔位置)

df_pick_rename.to_csv('/Users/hakukyokumorikawa/Desktop/rakuten.csv', encoding='utf_8_sig')

print(period)
import test20210425
test20210425.csv_to_xlsx_pd()

if __name__ == '__main__':
    test20210425.csv_to_xlsx_pd()

#將Excel的檔案排序
stexcel=pd.read_excel('/Users/hakukyokumorikawa/Desktop/rakuten.xlsx')
stexcel.sort_values(by='評價數',inplace=True,ascending=False)

print(stexcel)
#儲存排序後結果
stexcel.to_excel('/Users/hakukyokumorikawa/Desktop/'+str(genreid)+'.xlsx')

print('------------')

print("編號: "+genreid)
if int(genreid)== 101240:
    print("CD・DVD・楽器")
elif int(genreid)== 100804:
    print("インテリア・寝具 ")  
elif int(genreid)== 200166:
    print("収納家具")
elif int(genreid)==205515:
    print("ソファ")
elif int(genreid)==100823:
    print("ベッド関連")
elif int(genreid)==215566:
    print("寝具")
#-------------------------
elif int(genreid)==208252:
    print("介護寝具 ")
elif int(genreid)==111346:
    print("テーブル")
elif int(genreid)==111363:
    print("イス")
elif int(genreid)==210102:
    print("ダイニングセット")
elif int(genreid)==200167:
    print("デスク")
elif int(genreid)==100805:
    print("ライト・照明")
elif int(genreid)==111355:
    print("カーペット")
elif int(genreid)==205677:
    print("カーテン")
elif int(genreid)==205679:
    print("シェード・スクリーン")
elif int(genreid)==205680:
    print("ブラインド 百葉窗")
elif int(genreid)==205751:
    print("カーテンレール・用品")
elif int(genreid)==200119:
    print("インテリアファブリック")
elif int(genreid)==100863:
    print("裝飾小物")
elif int(genreid)==210143:
    print("鏡子類")
elif int(genreid)==203024:
    print("小孩房用品")
elif int(genreid)==200822:
    print("小孩房家具")
elif int(genreid)==200119:
    print("辦公家具")
elif int(genreid)==101863:
    print("洋式家具")
elif int(genreid)==101861:
    print("和式家具")
elif int(genreid)==101859:
    print("其它家具")
elif int(genreid)==101189:
    print("玩具")
elif int(genreid)==201695:
    print("腳踏車用品")
elif int(genreid)==215685:
    print("廚房收納用品")
elif int(genreid)==100094:
    print("USB小物")
elif int(genreid)==552925:
    print("零錢卡片包")
elif int(genreid)==201875:
    print("帳蓬類")
elif int(genreid)==201877:
    print("烤肉用具類")
elif int(genreid)==201887:
    print("露營睡袋")
elif int(genreid)==400782:
    print("傘架")
elif int(genreid)==216153:
    print("女用拖鞋")
elif int(genreid)==208214:
    print("圓座墊")
elif int(genreid)==565628:
    print("室內拖鞋")
elif int(genreid)==401449:
    print("隔熱墊")
elif int(genreid)==566068:
    print("廚房周邊小物")
elif int(genreid)==204964:
    print("高爾夫球")
elif int(genreid)==406401:
    print("厠紙架")
elif int(genreid)==215767:
    print("垃圾桶")
elif int(genreid)==565607:
    print("曬毛巾架")
elif int(genreid)==100645:
    print("洗衣周邊佩件")
elif int(genreid)==210168:
    print("廚房碗盤濾水架")
elif int(genreid)==501103:
    print("温水洗浄便座")
elif int(genreid)==568229:
    print("清潔刷，海綿")
elif int(genreid)==502792:
    print("吹風機")
elif int(genreid)==212585:
    print("熨斗")
elif int(genreid)==202245:
    print("美容儀器")
elif int(genreid)==212575:
    print("電剪脫毛器")