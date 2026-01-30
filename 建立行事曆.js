
function createCalendar() {
    // 1. 取得基本資料
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const people = {
        employee: sheet.getRange("B4").getValue(),
        manager: sheet.getRange("B8").getValue(),
        mentor: sheet.getRange("B10").getValue()
    };

    const scheduelData = sheet.getRange("D3:F11").getValues();
    const calendar = CalendarApp.getDefaultCalendar();

    scheduelData.forEach((row, index) => {
        const title = row[0];
        const date = row[1];
        const eventId = row[2];

        // 3. 定義邏輯：根據標題關鍵字決定參與者
        let attendees = [people.employee]; // 員工必過
        if (title.includes("Relay")) {
            attendees.push(people.manager);
        }
        if (title.includes("回饋")) {
            attendees.push(people.mentor);
        }

        // 如果已經有 ID，先刪除舊的事件
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

        try {
            // 4. (重新)建立日曆事件    
            const eventTitle = `${people.employee} ／ ${title}`;

            // 設定時間為 15:00 ~ 15:30
            const startTime = new Date(date);
            startTime.setHours(15, 0, 0);

            const endTime = new Date(date);
            endTime.setHours(15, 30, 0);

            const event = calendar.createEvent(eventTitle, startTime, endTime, {
                guests: attendees.join(", "),
                sendInvites: true,
                description: "（自動生成）入職重要時程"
            });

            // 將新的 Event ID 寫回試算表 F 欄
            sheet.getRange(index + 2, 6).setValue(event.getId());

        } catch (e) {
            console.log(`建立事件失敗 (${title}): ` + e.toString());
        }
    });

    SpreadsheetApp.getUi().alert("行事曆同步完成！");
}