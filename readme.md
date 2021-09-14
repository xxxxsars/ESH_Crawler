### AUO環安網頁爬蟲

一般使用者操作步驟:

Step 1:
到[release]( http://tcaigitlab.corpnet.auo.com/mfg/l3d/01/esh_crawler/-/releases "release")下載，並修改setting.ini參數

```commandline
#重要參數說明
[System]
lunch_date = <The system lunch date> #該程式上線時間，在這之後的申請單才作警告動作 格式如下:2021/08/24
mail_suffix = @auo123.com #email地址，因怕測試時會干擾到收件者，因此先定義為錯誤的mail address實際上線改為正確address


#user 區塊請輸入你的auo帳號密碼，須完整包含mail address
[User]
account = <Your AUO account> 
password = <Your AUO password>>

#爬蟲後存檔的位置，因為資安問題因此僅能存放在u槽，建議路徑為U:\ESH.xlsx
[Crawler]
...
xlsx_name = <Your xlsx path( notice: only allow saving it on "U" slot.) >
..
```

Step 2:利用cmd-line tool執行不同的function，cmd-line tool說明如下，特別注意action參數僅支援crawler和mail兩個參數。
```commandline
$ cd dist
$ python main.exe --help

       USAGE: main.py [flags]
flags:

main.py:
  --action: running specific action [crawler|mail].
    (default: 'crawler')
  --config: the configuration path.
    (default: 'setting.ini')
  --[no]debug: Turn on/off the debug mode (crawler will show the browser).
    (default: 'false')
  --[no]log: Turn logging on/off.
    (default: 'false')

```

Step 3:建議開啟log mode如在開發測試功能可以開啟debug mode，以下為指令範例
```commandline
$ main.exe --log true --debug true --action crawler  #這個指令會紀錄log，並在執行時顯示瀏覽器，執行的動作為爬蟲
$ main.exe --log true --action mail    #這個指令會紀錄log，在執行時不顯示瀏覽器，執行的動作為逾期警告信 (沒有帶debug就會設置為False，log參數同)

#windows排成建議設置指令如下(建議都開啟log mode，debug mode則不需要開啟):
$ main.exe --log true --action crawler
$ main.exe --log true --action mail
```


開發者操作步驟:

Step 1:
可以參考之前的proxy方式安裝，或者請it協助安裝
```commandline
$ pip install -r requirements.txt
```

Step 2:修改程式碼後，重新編譯exe
**注意***:若有新增套件，在pyinstaller有可能會無法hook到，請自行抓取套件相關的dll檔案放到，package_lib目錄中
```commandline
$ pyinstaller -D main.spec
```

Step 3:會在dist目錄下看到重新編譯好的exe

Step 4:參照上面"一般使用者操作步驟"，將修改後的setting.ini放到dist目錄即可使用