// 設定常用變數
const ss = SpreadsheetApp.getActiveSpreadsheet();   
const sheet = ss.getActiveSheet();
const baseDate = sheet.getRange("B6").getValue();

// 算出兩年內的所有國定假日
const holidayCal = CalendarApp.getCalendarById("zh-tw.taiwan#holiday@group.v.calendar.google.com");
const startTime = new Date();
const endTime = new Date(); 
endTime.setFullYear(startTime.getFullYear() + 2);

const twHolidays = holidayCal.getEvents(startTime, endTime);

function calculateMeetingDate() {
  const baseDate = new Date(sheet.getRange("B6").getValue());
  const finalDate = getFinalMeetingDate(baseDate);
  
  // 假設要把結果寫在 B7
  sheet.getRange("B7").setValue(finalDate);
  Logger.log("最終開會日：" + finalDate);
}

/**
 * 核心邏輯：調整日期
 */
function getFinalMeetingDate(date) {
  let resultDate = new Date(date);
  let dayOfWeek = resultDate.getDay();

  // 1. 處理週末初步調整
  // 逢週六(6) -> 訂週五(-1)；逢週日(0) -> 訂週一(+1)
  if (dayOfWeek === 6) {
    resultDate.setDate(resultDate.getDate() - 1);
  } else if (dayOfWeek === 0) {
    resultDate.setDate(resultDate.getDate() + 1);
  }

  // 2. 處理國定假日（若調整後仍是假日，則一律提前）
  let i = 0;
  // 只要 checkDayType 回傳不是 "工作日"，就一直往前減一天
  while (checkDayType(resultDate) !== "工作日" && i < 10) {
    resultDate.setDate(resultDate.getDate() - 1);
    i++; // 防止無限迴圈的安全機制
  }

  return resultDate;
}

/**
 * 檢查日期類型
 */
function checkDayType(checkDate) {
  const dayOfWeek = checkDate.getDay(); 
  const isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);
  const targetDateStr = Utilities.formatDate(checkDate, "GMT+8", "yyyy-MM-dd");

  const isHoliday = twHolidays.some(event => {
    const hDateStr = Utilities.formatDate(event.getStartTime(), "GMT+8", "yyyy-MM-dd");
    return hDateStr === targetDateStr;
  });

  if (isHoliday) return "國定假日";
  if (isWeekend) return "普通週末";
  return "工作日";
}

/**
 * 主程式：根據 baseDate 自動產生開會日期並垂直寫入 E2:E9
 */
function generateDate() {
  // 確保 baseDate 是日期物件
  const baseDateObj = new Date(baseDate); 
  
  const intervals = [7, 30, 49, 56, 90, 150, 180, 365];
  
  // 3. 跑迴圈計算每一個日期，並直接包成 [[date1], [date2]] 的格式
  const results = intervals.map(days => {
    let targetDate = new Date(baseDateObj);
    targetDate.setDate(targetDate.getDate() + days);
    
    let finalDate = getFinalMeetingDate(targetDate);
    // 回傳 [日期] 這樣最終會變成二維陣列
    return [finalDate]; 
  });

  // 4. 將結果垂直寫入 E2 開始的儲存格
  // E 是第 5 欄，起始列為 2，列數為 results.length，欄數為 1
  sheet.getRange(2, 5, results.length, 1).setValues(results);
}