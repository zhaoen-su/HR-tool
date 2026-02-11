function onOpen() {
  var ui = SpreadsheetApp.getUi(); // 取得 UI 環境
  
  ui.createMenu('✨功能表')      // 設定功能列顯示的名稱
    .addItem('建立／更新日期（迴避假日）', 'generateDate')
    .addSeparator()
    .addItem('建立行事曆','createCalendar')
    .addItem('刪除行事曆','deleteCalendar')
    .addSeparator()
    .addItem('建立新員工表單副本', 'createNewSpreadsheet')
    .addItem('匯出中文紀錄表', 'exportToChineseRecord')
    .addItem('匯出日文紀錄表', 'exportToEnglishRecord')
    .addItem('匯出韓文紀錄表', 'exportToKoreanRecord')
    .addToUi();
}