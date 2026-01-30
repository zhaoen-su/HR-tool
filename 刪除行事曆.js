
function deleteCalendar() {
    // 1. 取得基本資料
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const scheduelData = sheet.getRange("F2:F11").getValues();
    const calendar = CalendarApp.getDefaultCalendar();

    scheduelData.forEach((row) => {
        const eventId = row[0];

        if (eventId) {
            try {
                const existingEvent = calendar.getEventById(eventId);
                if (existingEvent) {
                    existingEvent.deleteEvent();
                }

                sheet.getRange("F2:F11").clearContent();
            } catch (e) {
                console.log(`刪除舊事件失敗 (可能已不存在): ` + e.toString());
            }
        }
    });
}