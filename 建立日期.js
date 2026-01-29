/**
 * 檢查日期並自動調整：週六改週五、週日改週一、避開國定假日。
 * * @param {Date} inputDate 要檢查的日期
 * @return 調整後的工作日
 * @customfunction
 */
function ADJUST_WORKDAY(inputDate) {
  if (!(inputDate instanceof Date)) {
    // 處理若是從儲存格傳進來的日期字串
    inputDate = new Date(inputDate);
  }
  
  if (isNaN(inputDate.getTime())) return "格式錯誤";

  let tempDate = new Date(inputDate);
  let isWorkDay = false;
  
  // 取得台灣節慶日曆
  const holidayCal = CalendarApp.getCalendarById("zh-tw.taiwan#holiday@group.v.calendar.google.com");

  // 避免無限迴圈，設定最大嘗試次數（例如 10 天）
  let safetyCounter = 0;

  while (!isWorkDay && safetyCounter < 10) {
    let day = tempDate.getDay(); // 0:日, 6:六
    let holidays = holidayCal.getEventsForDay(tempDate);

    if (day === 6) { 
      // 遇週六 -> 改週五
      tempDate.setDate(tempDate.getDate() - 1);
    } else if (day === 0) { 
      // 遇週日 -> 改週一
      tempDate.setDate(tempDate.getDate() + 1);
    } else if (holidays.length > 0) {
      // 遇國定假日 -> 往後延
      tempDate.setDate(tempDate.getDate() + 1);
    } else {
      isWorkDay = true;
    }
    safetyCounter++;
  }
  
  return tempDate;
}