function onOpen() {
  var ui = SpreadsheetApp.getUi(); // 取得 UI 環境
  
  ui.createMenu('✨功能表')      // 設定功能列顯示的名稱
    .addItem('建立日期', 'ADJUST_WORKDAY')
    .addSeparator()
    .addItem('建立新員工表單副本', 'createNewSpreadsheet') // (顯示名稱, 執行的函數名稱)
    .addToUi();                  // 必須呼叫這行才會渲染到畫面上
}