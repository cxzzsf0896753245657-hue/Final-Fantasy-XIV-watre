# FF14 潛水艇計時器 — 檔案說明

```
index.html               網站本體（程式與版面）
data/                    網站數據（由數據總表轉出）
  sectors.js             海域、地點、航行距離、等級、部件
  loot.js                掉落物：層級、機率、數量
  items.js               物品名稱、高優先物品
  presets.js             常用配件預設
  craft.js               配件製作配方
  version.js             數據版本號（讓玩家的瀏覽器知道要重新下載數據）
tools/
  excel-to-data.html     數據總表（xlsx）→ data 檔案的轉換工具，用瀏覽器打開即可
  convert.js             轉換工具的程式
  vendor/xlsx.full.min.js  讀取 xlsx 用的函式庫（SheetJS，Apache-2.0）
```

## 部署
把整個資料夾（index.html、data、tools）放上網站即可。也可以直接雙擊 index.html 在本機使用。

## 版本更新數據
1. 在「FF14潛水艇數據表.xlsx」修改或新增資料。
2. 用瀏覽器打開 `tools/excel-to-data.html`，把 xlsx 拖進去，下載 6 個檔案。
3. 用下載的檔案覆蓋 `data/` 裡的同名檔案，就完成了。index.html 不用改。

新增地點時，在數據總表的「地點」填上座標，航行距離會自動計算。
