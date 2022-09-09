# AttendManageGAS
Create attendance management spreadsheets on Google SpreadSheet.  
Googleスプレッドシート上で出席率管理表を自動で生成します．

## How to Use
1. 拡張機能からApp Scriptを開く
2. main.jsをコピーして貼り付ける
3. setup()関数を実行
4. Configシートに情報を記入
5. createStatisticSheet()関数を実行

## Config
Configは，列のセルに分けて記入します．最大記入数は20です．（main.jsのmaxWidthを書き換えることで検出範囲を増やすことができます）
|項目|入力例（特筆のない場合，記入式）|
|---:|:---|
|実施回|1st|
|時間帯|月曜前半|
|場所|A教室|
|班数|3|
|集計分類|出席，未処理，欠席，見なし解約など|
|集計区分|出席，未処理，欠席，集計除外（プルダウン選択式）|
