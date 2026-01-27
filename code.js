const CONFIG = {
  TEMPLATE_ID: 'YOUR_TEMPLATE_SHEET_ID',
  FOLDER_ID: 'YOUR_TARGET_FOLDER_ID',
  MILESTONES: [
    { label: '到職日', offset: 0 },
    { label: '到職第一週', offset: 7 },
    { label: '到職第三週', offset: 21 },
    { label: '到職第一個月', offset: 30 },
    { label: '到職三個月', offset: 90 },
    { label: '到職半年', offset: 180 },
    { label: '到職一年', offset: 365 }
  ]
};

function generateEmployeeOnboarding() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getRange().getValues()[0]; // 取得選取列
  
  // 解構輸入資料
  const [name, dept, manager, mentor, startDate] = data;
  
  // 1. 複製專屬表單
  const newFile = DriveApp.getFileById(CONFIG.TEMPLATE_ID)
    .makeCopy(`HR_Onboarding_${name}_${dept}`, DriveApp.getFolderById(CONFIG.FOLDER_ID));
  
  const newSheet = SpreadsheetApp.open(newFile).getSheets()[0];
  
  // 2. 寫入里程碑日期
  CONFIG.MILESTONES.forEach((ms, index) => {
    const targetDate = calculateTargetDate(new Date(startDate), ms.offset);
    newSheet.getRange(index + 2, 1).setValue(ms.label);
    newSheet.getRange(index + 2, 2).setValue(targetDate);
    
    // 3. 同步至 Google Calendar
    createCalendarEvent(name, ms.label, targetDate, [manager, mentor]);
  });
  
  // 4. 個資權限設定 (僅限本人與相關人)
  // newFile.addEditor(employeeEmail); 
}

function createCalendarEvent(empName, title, date, guests) {
  const calendar = CalendarApp.getDefaultCalendar();
  calendar.createAllDayEvent(`[HR] ${empName} - ${title}`, date, {
    description: `部門相關人員：${guests.join(', ')}`,
    guests: guests.join(','), // 自動發送邀請給主管與 Mentor
    sendInvites: true
  });
}