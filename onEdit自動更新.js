function syncDateToCalendar(e) {
   const ui = SpreadsheetApp.getUi();
  const lock = LockService.getScriptLock();

  try {
    // 1. 取得鎖定，避免併發衝突
    lock.waitLock(5000);

    const range = e.range;
    const sheet = range.getSheet();

    // 2. 確認名稱一致
    if (sheet.getName() !== "行事曆控制表" || range.getColumn() !== 3 || range.getRow() < 2) {
      lock.releaseLock();
      return;
    }

    const row = range.getRow();
    const cellValue = range.getValue(); // 這裡拿到的通常已經是 Date 物件
    const eventId = sheet.getRange(row, 4).getValue();

    if (!eventId || !cellValue) {
      lock.releaseLock();
      return;
    }

    // 3. 確保轉換為有效的日期物件
    let targetDate;
    if (cellValue instanceof Date) {
      targetDate = cellValue;
    } else {
      targetDate = new Date(cellValue);
    }

    // 4. 防呆：檢查日期是否合法
    if (isNaN(targetDate.getTime())) {
      console.error("無效日期: " + cellValue);
      lock.releaseLock();
      return;
    }

    const calendar = CalendarApp.getDefaultCalendar();
    const event = calendar.getEventById(eventId.toString());

    if (event) {
      // 5. 執行更新
      event.setAllDayDate(targetDate);
      
      // 改用 toast，避免每次都要點彈窗
      SpreadsheetApp.getActiveSpreadsheet().toast(`成功同步日期：${Utilities.formatDate(targetDate, Session.getScriptTimeZone(), "yyyy/MM/dd")}`, "✅ 同步成功");
    }

    lock.releaseLock();

  } catch (err) {
    if (lock.hasLock()) lock.releaseLock();
    // 只有出錯才彈 alert
    ui.alert('⚠ 同步失敗', '錯誤訊息：' + err.message, ui.ButtonSet.OK);
  }
}