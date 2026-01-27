function onOpen() {
  var ui = SpreadsheetApp.getUi(); // 取得 UI 環境
  
  ui.createMenu('✨功能表')      // 設定功能列顯示的名稱
    .addItem('幫新員工建表', 'createNewSpreadsheet') // (顯示名稱, 執行的函數名稱)
    .addToUi();                  // 必須呼叫這行才會渲染到畫面上
}

function createNewSpreadsheet() {
  try {
    // 1. 建立新的試算表檔案
    const timestamp = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");
    const newSS = SpreadsheetApp.create("測試" + timestamp);
    const sheet = newSS.getSheets()[0]; // 取得第一個分頁
  } catch (e) {
    console.error("建立失敗: " + e.toString());
  }
}