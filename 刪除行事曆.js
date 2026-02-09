function deleteCalendar() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("行事曆控制表");
  const activeRange = sheet.getActiveRange();
  const startRow = activeRange.getRow();
  const numRows = activeRange.getNumRows();
  const eventIds = activeRange.getValues();
  const calendar = CalendarApp.getDefaultCalendar();

  // --- 防呆提示框 ---
  const response = ui.alert(
    '系統提示',
    `您確定要刪除選取的 ${numRows} 列資料，並同時移除 Google 行事曆上的事件嗎？\n(此動作無法復原)`,
    ui.ButtonSet.YES_NO
  );

  // 如果使用者點選「否」或直接關閉視窗，則終止程式
  if (response !== ui.Button.YES) {
    console.log("使用者取消了刪除動作");
    return;
  }

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