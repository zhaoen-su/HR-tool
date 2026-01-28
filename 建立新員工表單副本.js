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