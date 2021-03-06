﻿#Google Spreadsheet Update Log
﻿Version: 1.0.0.0 Beta
--------------------------
One day, when I am logged what I am doing, done, need to do and so on.
I create a log sheet to mark all the things and the status, that was good.
After a 10 record are recorded, it is difficult to read which tasks are not done.
I have given a related color to each status, pending is yellow, awaiting color, green is safe, a work done color.
Now I can easy to find which task are not done by color, that was good.
After that, I always change the background color by myself, that was so tired to be a robot.

That's why the app script come out, I would like to do those things by Google automatically.

Today date is 2013-12-06, a discouraged and exhilarating day.
I found that Google Spreadsheet provide 2 new function, when i right click in a cell

* Coordination formatting
* Data validation

The first one make me lose heart, my first mind: What the?!!!
That is definitely replacing what my app script written for if it change a row background
I'm totally lose, how can i fighting with google development team?
No, no, can not stop now, I still have something would like Google helping me to do so.

On the another hands, the second function is extremely

一天，在工作時因為要記錄做了什麼。一份記錄清單便開始了，起初沒有什麼特別事發生，
記錄開始多了，雖然有標明狀態，但還是看得很吃力。於是為每種狀態想一種顏色，
等待的工作是汽車在等黃燈的黃色，完成後的工作是安全的綠色。被判決為回報錯誤的或是已取消的工作是灰色，如此如此這般這般
當加了一項記錄，便根據狀態自己轉顏色，再加入記錄的日期和完成工作的日期。每次都是一項記錄才進行的動作還可以認付。
做多了就發覺很累人，所以想做一條大懶蟲的我就再花多一點時間寫一段script幫我做……

先是寫了一段幫整張Sheet(Excel叫工作表)根據不同狀況轉換不同顏色，但每次執行時也要自行按一次執行又太麻煩。
太煩了不想按執行，但是不按又看得很辛苦又很麻煩。
Google應該沒有那麼笨可以自動幫我做吧，到Google找了找又再寫多一段，漸漸就成了現在的幾段。

今天是2013-12-06，發現了Google Spreadsheet提供了二個新功能

* 條件式格式設定
* 驗證資料

第一個幾符完全取代了我寫下的app script，只差一點點就被google擺了一道。
差點全敗了，一個人的作業與Google開發團體的作業，輸贏早定了，不行，燃燒起來了。

話說回來，我愛死第二個功能了。可以令column成為 unique KEY，即係可以變成Primary key，
有了PK即是可以洐生出Foreign key，而且還可以建立範圍清單。哈哈哈，那代表提供了更多空間可以做到更多事情。

##Demo
--------------------------
<a href"http://keithbox.blogspot.com/2014/02/google-spreadsheet-update-log.html" target="_blank">Installation / Configuration</a>
<br>
<a href="https://docs.google.com/spreadsheet/ccc?key=0AlaVan9pZtAzdHQzeU5iRURUQ3JLNnAydHZaa3JMdUE&usp=drive_web#gid=0" target="_blank">Update Log - Template</a>

##Usage
--------------------------
###You need to change a few of value in the top of app script
```
var spreadsheetID = "0AlaVan9pZtAzdEF5Wm9HQzFiTlpNQVF4a3hmWDJxSGc";
var logSheetName = "Log";
var customizeStatusColorSheetName = "Status Color";
var statusChangeColumnName = "Status";
```
`spreadsheetID:` Replace with your spreadsheet ID<br>
`logSheetName:` This the name of your log sheet<br>
`customizeStatusColorSheetName:` If you would like overload the status color, 
put your own status color sheet name here, 
if you do not understand what this for just ignore it.<br>
`statusChangeColumnName:` This is your "Status" column header<br>

```
var backgroundColorPriority = [
  ["tailor make", "hardcode", "holding", "follow up", "misreporting", "cancelled", "pending", "release", "done"],
  ["#d9d2e9", "#f4cccc", "#f4cccc", "#c9daf8", "#efefef", "#efefef", "#fff2cc", "#d9ead3", "#d9ead3"]
];
```
Chagne the following color to your own, please notice that the status color are <br>


```
var addTodayWhenEdit = [
  ["Report By", "Report Date"],
  ["Completed By", "Completed Date"]
];
```

##Change Log
--------------------------
* (Future): Auto insert #.(issue) number
* (Future): Auto insert default status after typing issue description
* Can specify the status name, color and the priority in "Status Color" sheet without any coding
(low performance, not recommend to use if you have known how to change backgroundColorPriority)
> create a sheet call 'Status Color', the should be
> 
> | Status | Color in Hex/RGB | Priority |
> |:-----|:----------|:---------------|
> | Done | rgb(201,218,248) | 1 |
> | Pending | 255,242,204 | 2 |
> | testing | no color will set to white | no priority will set to the lowest

* Can specify the 'status change' column by your own

> change the variable statusChangeColumnName value

* Can set the status color in difference priority
* Can specify the status and color by your own
- (Auto) Insert current date after typing 'report by someone', you can specify the 'report by' column and 'report date' column
- (Auto) Color change after onEdit in 'status' column

## License
--------------------------
Please see the [LICENSE][license] file for further details.

[license]: https://github.com/keithbox/Google-Spreadsheet-Update-Log/blob/master/LICENSE


##Reference
--------------------------
| Session | Topic | URL | 
|:-----|:----------|:---------------|
| 0 | Understanding Events | <a href="https://developers.google.com/apps-script/understanding_events?hl=en" target="_blank">https://developers.google.com/apps-script/understanding_events?hl=en</a>
| 1 | Understanding Triggers | <a href="https://developers.google.com/apps-script/understanding_events?hl=en" target="_blank">https://developers.google.com/apps-script/understanding_events?hl=en</a>
| 2 | Class Spreadsheet | <a href="https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet" target="_blank">https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet</a>
| 3 | Class Sheet | <a href="https://developers.google.com/apps-script/reference/spreadsheet/sheet" target="_blank">https://developers.google.com/apps-script/reference/spreadsheet/sheet</a>
| 4 | Class Range | <a href="https://developers.google.com/apps-script/reference/spreadsheet/range" target="_blank">https://developers.google.com/apps-script/reference/spreadsheet/range</a>
| 5 | Spreadsheets size and complexity limits | <a href="https://support.google.com/drive/answer/2505921" target="_blank">https://support.google.com/drive/answer/2505921</a>

