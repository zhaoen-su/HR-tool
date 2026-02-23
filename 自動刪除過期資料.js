function deleteExpiredRows() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("行事曆記錄");
  var rows = sheet.getDataRange().getValues();
  var today = new Date();
  
  // 從最後一行往回刪除，才不會因為行數變動導致錯位
  for (var i = rows.length - 1; i >= 1; i--) {
    var expiryDate = new Date(rows[i][2]);
    if (expiryDate < today) {
      sheet.deleteRow(i + 1);
    }
  }
}