
function createCalendar() {
    // 1. 取得基本資料
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("各項時程");
    const people = {
        employeeName: sheet.getRange("B1").getValue(),
        employeeEmail: sheet.getRange("C1").getValue(),
        managerName: sheet.getRange("B15").getValue(),
        managerEmail: sheet.getRange("C15").getValue(),
        mentorName: sheet.getRange("B16").getValue(),
        mentorEmail: sheet.getRange("C16").getValue(),
        hrName: sheet.getRange("B17").getValue(),
        hrEmail: sheet.getRange("C17").getValue()
    };

    const scheduelData = sheet.getRange("A3:D12").getValues();
    const calendar = CalendarApp.getCalendarById("capsulecorporation.cc_b4fsbnukr1r08evseau268f7qo@group.calendar.google.com");

    scheduelData.forEach((row, index) => {
        const title = row[0];
        const date = row[2];
        const personInCharge = row[3];

        // 3. 定義邏輯：根據標題關鍵字決定參與者與對象姓名
        let attendees = [];
        let targetName = "";

        if (title.includes("主管")) {
            attendees.push(people.managerEmail);
            targetName = people.managerName;
        } else if (title.includes("Mentor")) {
            attendees.push(people.mentorEmail);
            targetName = people.mentorName;
        } else if (title.includes("HR")) {
            attendees.push(people.hrEmail);
            targetName = people.hrName;
        }

        try {
            // 4. (重新)建立日曆事件    
            // 如果有特定對象 (targetName)，就加入標題
            const eventTitle = targetName
                ? `${people.employeeName}／${targetName}　${title}`
                : `${people.employeeName}　${title}`;

            const event = calendar.createAllDayEvent(eventTitle, date, {
                guests: attendees.join(", "),
                sendInvites: false,
                description: "（自動生成）入職重要時程"
            });

            // 將新的 Event ID 寫回試算表 F 欄
            // sheet.getRange(index + 2, 6).setValue(event.getId());
            ss.getSheetByName("行事曆紀錄").appendRow([people.employeeName, title, date, event.getId()]);

        } catch (e) {
            console.log(`建立事件失敗 (${title}): ` + e.toString());
        }
    });
}

// 供本地測試環境使用
if (typeof module !== 'undefined') {
    module.exports = { createCalendar };
}