function onOpen() {
  var ui = SpreadsheetApp.getUi(); // 取得 UI 環境
  
  ui.createMenu('✨功能表')      // 設定功能列顯示的名稱
    .addItem('建立／更新日期', 'generateDate')
    .addSeparator()
    .addItem('建立新員工表單副本', 'createNewSpreadsheet')
    .addSeparator()
    .addItem('建立行事曆','createCalendar')
    .addItem('刪除行事曆','deleteCalendar')
    .addToUi();
}