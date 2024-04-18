## 概要

毎日手動でエクセルファイルを複製して実績シートを作成していたが、地味に面倒なので可能な限り自動化。  

## 機能要件

* 毎朝7時になったら当日の日付が入った実績シートが作成される（厳密には直近のファイルを複製する）
* 月を跨ぐタイミングは格納フォルダも作成する（例：202404）
* 土・日はスキップする

**▼エクセルを格納するpath（例）**  
/path/to/dir/_日報実績シート/  
　- 202403, 202404, etc...  
　　- WFH Daily Report ‐ name - 20240417.xlsx, WFH Daily Report ‐ name - 20240418.xlsx, etc...  
    
## 実装環境

* MacOS - M2 Apple silicon
* Python 3.7.11（3系であれば問題ないと思われる）
* openpyxl インストール済（エクセルを読み書きするためのライブラリ）

## 実装手順

1. 自動化したいフォルダとexcel_auto_creat.py内のpathを合わせる
2. ローカル環境内で ```cron``` を設定

```
## 平日毎朝7時にexcel_auto_creat.pyを実行

crontab -e
0 7 * * 1-5 /usr/local/bin/python3 /path/to/dir/excel_auto_creat.py
```

done!
