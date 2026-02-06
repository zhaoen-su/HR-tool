function deleteCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("行事曆控制表");
  const activeRange = sheet.getActiveRange();
  const startRow = activeRange.getRow();
  const numRows = activeRange.getNumRows();
  const eventIds = activeRange.getValues();
  const calendar = CalendarApp.getDefaultCalendar();

  // 從選取範圍的「最後一列」開始往回跑
  // 這樣刪除列時，上方的列號才不會變動
  for (let i = numRows - 1; i >= 0; i--) {
    const eventId = eventIds[i][0]; // 假設 ID 在選取範圍的第一欄
    const currentRow = startRow + i;

    if (eventId) {
      try {
        const existingEvent = calendar.getEventById(eventId.toString());
        if (existingEvent) {
          existingEvent.deleteEvent();
          console.log(`已刪除行事曆事件: ${eventId}`);
        }
      } catch (e) {
        console.log(`行事曆事件不存在或 ID 無效 (列 ${currentRow}): ` + e.toString());
      }
    }

    // 執行刪除整列
    sheet.deleteRow(currentRow);
    console.log(`已刪除試算表第 ${currentRow} 列`);
  }
}