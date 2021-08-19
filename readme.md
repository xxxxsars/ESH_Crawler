### AUO環安網頁爬蟲

操作步驟:

Step 1:
可以參考之前的proxy方式安裝，或者請it協助安裝
```commandline
$ pip install -r requirements.txt
```

Step 2:
修改setting_sample.ini檔名為setting.ini，並修改參數
```commandline
#重要參數說明

#user 區塊請輸入你的auo帳號密碼，須完整包含mail address
[User]
account = <Your AUO account> 
password = <Your AUO password>>

#爬蟲後存檔的位置，因為資安問題因此僅能存放在u槽，建議路徑為U:\ESH.xlsx
[Crawler]
...
xlsx_name = <Your xlsx path( notice: only allow saving it on "U" slot.) >

#Recipient(收件者)至少要有一個，可支援多個收件者，recipient1 recipient2... 做區隔即可
[Recipient]
..
#多個copy(cc副本)，同理用copy1 copy2區隔
[Copy]
..
```

Step 3:利用cmd-line tool執行不同的function，cmd-line tool說明如下，特別注意action參數僅支援crawler和alert兩個參數。
```commandline

$ python main.py --help

       USAGE: main.py [flags]
flags:

main.py:
  --action: running specific action [crawler|alert].
    (default: 'crawler')
  --config: the configuration path.
    (default: 'setting.ini')

```
