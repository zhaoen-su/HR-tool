
function deleteCalendar() {
    // 1. 取得基本資料
    const scheduelData = sheet.getRange("F3:F11").getValues();
    const calendar = CalendarApp.getDefaultCalendar();

    scheduelData.forEach((row) => {
        const eventId = row[0];

        if (eventId) {
            try {
                const existingEvent = calendar.getEventById(eventId);
                if (existingEvent) {
                    existingEvent.deleteEvent();
                }
            } catch (e) {
                console.log(`刪除舊事件失敗 (可能已不存在): ` + e.toString());
            }
        }
    });

    SpreadsheetApp.getActiveSpreadsheet().toast("行事曆刪除完成！");
}