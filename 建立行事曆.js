
function createCalendar() {
    // 1. 取得基本資料
    const ss = SpreadsheetApp.getActiveSpreadsheet();   
    const sheet = ss.getActiveSheet();
    const employee = sheet.getRange("B4").getValue();
    const manager = sheet.getRange("B8").getValue();
    const mentor = sheet.getRange("B10").getValue();

    const scheduelData = sheet.getRange("D3:E11").getValues();
    scheduelData.forEach(row => {
        const title = row[0];
        const date = row[1];
    })
    
    try {
        // 2. 建立日曆
        const calendar = CalendarApp.getDefaultCalendar();

        const event = (title, date, attendees) => {
            calendar.createAllDayEvent(title, date, {
                guests: attendees.join(", "),
                sendInvites: true,
                description: "（自動生成）入職重要時程"
            });
        }

        // 3. 根據規則建立獨立事件
        event(`${employee} title`, baseDate, [employee, manager, mentor]);
    } catch (e) {
        SpreadsheetApp.getUi().alert("建立失敗: " + e.toString()); // 顯示錯誤給使用者
    }
}