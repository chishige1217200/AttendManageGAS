# AttendManageGAS
Create attendance management spreadsheets on Google SpreadSheet.  
Googleスプレッドシート上で出席率管理表を自動で生成します．

## How to Use
1. 拡張機能からApp Scriptを開く
2. main.jsをコピーして貼り付ける
3. setup()関数を実行
4. Configシートに情報を記入
5. createBase()関数を実行

## Config
Configは，列のセルに分けて記入します．最大記入数は20です．（main.jsのmax_widthを書き換えることで検出範囲を増やすことができます）
|項目|入力例|
|---:|:---|
|実施回|1st|
|時間帯|月曜前半|
|場所|A教室|
|班数|3|
|統計区別|出席，欠席，未処理など|
|出席（とみなす統計区別の）要素|出席，一時振替出席など|
|未処理（とみなす統計区別の）要素|未処理，一時振替未処理など|
