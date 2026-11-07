/* FF14 潛水艇計時器 — 常用配件預設
 * 由 tools/excel-to-data.html 從「FF14潛水艇數據表.xlsx」產生，也可以直接手動編輯。
 * 每個欄位的值都是標準 JSON；各欄位說明請見數據總表的「說明」分頁。 */
window.SUB_DATA = Object.assign(window.SUB_DATA || {}, {
  "GP_PRESETS": {"lv_1121":{"hull":"1","stern":"1","bow":"2","bridge":"1"},"lv_1122":{"hull":"1","stern":"1","bow":"2","bridge":"2"},"lv_3113":{"hull":"3","stern":"1","bow":"1","bridge":"3"},"gold_3124":{"hull":"3","stern":"1","bow":"2","bridge":"4"},"item_4254":{"hull":"4","stern":"2","bow":"5","bridge":"4"},"item_4254kai":{"hull":"4改","stern":"2改","bow":"5改","bridge":"4改"},"item_5124kai":{"hull":"5改","stern":"1改","bow":"2改","bridge":"4改"}},
  "GP_PRESET_HINTS": {"lv_1121":"20等後可換，懶人通用練級，可用到85等或更換配置","lv_1122":"20等後可換，無詳細資料","lv_3113":"25–50等適用，僅限溺沒海為最佳練級配置","gold_3124":"85等跑 OJ　90等跑 ROJZ　95等跑 MROJZ","item_4254":"可撈絕大部分潛艇產物","item_4254kai":"更注重探索和恩惠，撈最新海域高稀有物品佳，成本高","item_5124kai":"專撈工坊材料"}
});
