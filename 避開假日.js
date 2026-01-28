/**
 * 檢查日期是否為週末或國定假日，並自動調整：
 * 遇週六 -> 改週五
 * 遇週日 -> 改週一
 * 遇國定假日 -> 往後延一天
 */
function avoidHoliday(date) {
  let tempDate = new Date(date);
  let isWorkDay = false;
  
  // 取得台灣節慶日曆 (如果你在其他地區，請更換日曆 ID)
  const holidayCal = CalendarApp.getCalendarById("zh-tw.taiwan#holiday@group.v.calendar.google.com");

  while (!isWorkDay) {
    let day = tempDate.getDay(); // 0:日, 6:六
    let holidays = holidayCal.getEventsForDay(tempDate);

    if (day === 6) { 
      // 遇週六 -> 改週五
      tempDate.setDate(tempDate.getDate() - 1);
    } else if (day === 0) { 
      // 遇週日 -> 改週一
      tempDate.setDate(tempDate.getDate() + 1);
    } else if (holidays.length > 0) {
      // 遇國定假日 -> 往後延一天 (可依需求改為往前)
      tempDate.setDate(tempDate.getDate() + 1);
    } else {
      // 既不是週末也不是假日，確定為工作日
      isWorkDay = true;
    }
  }
  return tempDate;
}