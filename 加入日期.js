// 設定常用變數
const ss = SpreadsheetApp.getActiveSpreadsheet();
const sheet = ss.getActiveSheet();
const baseDate = sheet.getRange("C3").getValue();

/**
 * 取得未來兩年內的國定假日 (回傳 Set 格式以提升查詢效能)
 */
function getTwNationalHolidays() {
  const calendar = CalendarApp.getCalendarById("zh-tw.taiwan#holiday@group.v.calendar.google.com");
  const startTime = new Date();
  const endTime = new Date();
  endTime.setFullYear(startTime.getFullYear() + 2);

  const events = calendar.getEvents(startTime, endTime);

  // 將日期轉為 "yyyy-MM-dd" 字串並存入 Set，查詢速度 O(1)
  const twNationalHolidays = new Set(
    events.map(event => Utilities.formatDate(event.getStartTime(), "GMT+8", "yyyy-MM-dd"))
  );

  return twNationalHolidays;
}

/**
 * 核心邏輯：計算日期 (避開週末與國定假日)
 */
function getFinalDate(date, holidaySet) {
  let resultDate = new Date(date);
  let weekday = resultDate.getDay();

  // 1. 週末初步調整：週六往前推1天，週日往後推1天
  if (weekday === 6) { // 週六
    resultDate.setDate(resultDate.getDate() - 1);
  } else if (weekday === 0) { // 週日
    resultDate.setDate(resultDate.getDate() + 1);
  }

  // 2. 檢查是否遇到國定假日 (若調整後是假日，則持續提前)
  let loopGuard = 0;
  // 只要是「非工作日」(即假日或週末)，就往前推一天
  while (!isWorkDay(resultDate, holidaySet) && loopGuard < 10) {
    resultDate.setDate(resultDate.getDate() - 1);
    loopGuard++;
  }

  return resultDate;
}

/**
 * 檢查是否為工作日 (Boolean)
 * true: 是工作日
 * false: 是假日或週末
 */
function isWorkDay(checkDate, holidaySet) {
  const weekday = checkDate.getDay();
  const isWeekend = (weekday === 0 || weekday === 6);
  const dateStr = Utilities.formatDate(checkDate, "GMT+8", "yyyy-MM-dd");

  // 如果是國定假日 OR 是週末，就代表「不是工作日」
  if (holidaySet.has(dateStr) || isWeekend) {
    return false;
  }
  return true;
}

/**
 * 主程式：根據 baseDate 自動產生開會日期並垂直寫入 E2
 */
function generateDate() {
  // 1. 先抓取一次假日資料 (轉為 Set)
  const holidaySet = getTwNationalHolidays();

  const intervals = [14, 35, 42, 56, 90, 90, 150, 180, 365];

  // 3. 跑迴圈計算每一個日期，並直接包成 [[date1], [date2]] 的格式
  const results = intervals.map(days => {
    let targetDate = new Date(baseDate);
    targetDate.setDate(targetDate.getDate() + days);

    // 將 holidaySet 傳入以供查詢
    let finalDate = getFinalDate(targetDate, holidaySet);
    return [finalDate];
  });

  // 3. 輸出結果到 E 欄 (從第 2 列開始)
  sheet.getRange(4, 3, results.length, 1).setValues(results);
}