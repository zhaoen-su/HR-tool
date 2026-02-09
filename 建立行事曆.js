
function createCalendar() {
    // 1. 取得基本資料
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("資料輸入區");
    const people = {
        employeeName: sheet.getRange("B2").getValue(),
        employeeEmail: sheet.getRange("B4").getValue(),
        manager: sheet.getRange("B8").getValue(),
        mentor: sheet.getRange("B10").getValue()
    };

    const scheduelData = sheet.getRange("D3:F10").getValues();
    const calendar = CalendarApp.getDefaultCalendar();

    scheduelData.forEach((row, index) => {
        const title = row[0];
        const date = row[1];
        const eventId = row[2];

        // 3. 定義邏輯：根據標題關鍵字決定參與者
        let attendees = [people.employeeEmail]; // 員工必過
        // if (title.includes("Relay")) {
        //     attendees.push(people.manager);
        // }
        // if (title.includes("回饋")) {
        //     attendees.push(people.mentor);
        // }

        try {
            // 4. (重新)建立日曆事件    
            const eventTitle = `${people.employeeName} ／ ${title}`;

            const event = calendar.createAllDayEvent(eventTitle, date, {
                guests: attendees.join(", "),
                sendInvites: false,
                description: "（自動生成）入職重要時程"
            });

            // 將新的 Event ID 寫回試算表 F 欄
            // sheet.getRange(index + 2, 6).setValue(event.getId());
            ss.getSheetByName("行事曆控制表").appendRow([people.employeeName, title, date, event.getId()]);

        } catch (e) {
            console.log(`建立事件失敗 (${title}): ` + e.toString());
        }
    });
}

// 供本地測試環境使用
if (typeof module !== 'undefined') {
    module.exports = { createCalendar };
}