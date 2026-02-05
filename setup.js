// 在你的測試檔頂部加入這段
global.SpreadsheetApp = {
  getActiveSpreadsheet: () => ({
    getSheetByName: () => ({
      getRange: () => ({
        getValue: () => "測試資料"
      })
    })
  })
};