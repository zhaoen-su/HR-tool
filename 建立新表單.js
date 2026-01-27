function createTestSpreadsheet() {
  try {
    // 1. 建立新的試算表檔案
    const timestamp = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");
    const newSS = SpreadsheetApp.create("測試" + timestamp);
    const sheet = newSS.getSheets()[0]; // 取得第一個分頁
  } catch (e) {
    console.error("建立失敗: " + e.toString());
  }
}


// // 2. 定義標題列與測試資料
// const header = ["編號", "項目名稱", "測試狀態", "建立時間"];
// const testData = [
//   [1, "功能模組 A", "通過", timestamp],
//   [2, "功能模組 B", "待測", timestamp],
//   [3, "API 連線", "失敗", timestamp]
// ];

// // 3. 寫入資料到試算表
// // 寫入標題
// sheet.getRange(1, 1, 1, header.length).setValues([header]);

// // 寫入多列測試資料 (從第 2 列開始)
// sheet.getRange(2, 1, testData.length, testData[0].length).setValues(testData);

// // 4. 基本美化 (選用)
// sheet.getRange(1, 1, 1, header.length)
//      .setBackground("#4A86E8") // 藍色背景
//      .setFontColor("white")     // 白色字體
//      .setFontWeight("bold");    // 加粗

// // 5. 自動調整欄寬
// sheet.autoResizeColumns(1, header.length);

// // 輸出結果至 log
// console.log("測試試算表已成功建立！");
// console.log("檔案網址: " + newSS.getUrl());

// return newSS.getUrl(); // 回傳網址供後續使用
