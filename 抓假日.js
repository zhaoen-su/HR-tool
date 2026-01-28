/**
 * 同步台灣 Google 公開假期 + 合併表內既有手動假日，並自動去重
 * 表頭固定為：Date | Name | Source
 * - Date：日期（去除時間）
 * - Name：節日名稱（多個以 ／ 合併）
 * - Source：資料來源（Google Holiday ICS／Manual 等，以 ／ 合併）
 */
function syncTaiwanHolidays() {
  const ICS_URL = "https://www.google.com/calendar/ical/taiwan__zh_tw%40holiday.calendar.google.com/public/basic.ics";
  const SHEET = "假日清單";
  const SRC_ICS = "Google Holiday ICS";
  const YEAR_FROM = new Date().getFullYear() ; // 可調整，-1就是前一年（現在是從當年度開始抓）
  const YEAR_TO   = new Date().getFullYear() + 1; // 可調整，現在是抓到明年

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  const tz = Session.getScriptTimeZone();

  // 確保表頭
  sh.getRange(1,1,1,3).setValues([["Date","Name","Source"]]);

  // 讀取現有資料，先建 map（保留手動資料）
  const last = sh.getLastRow();
  const map = new Map(); // key=yyyy-MM-dd -> {date, names:Set, sources:Set}
  if (last >= 2) {
    const existing = sh.getRange(2,1,last-1,3).getValues();
    for (const [d, name, src] of existing) {
      if (!(d instanceof Date)) continue;
      const key = Utilities.formatDate(d, tz, "yyyy-MM-dd");
      if (!map.has(key)) map.set(key, {
        date: new Date(d.getFullYear(), d.getMonth(), d.getDate()),
        names: new Set(),
        sources: new Set()
      });
      if (name) map.get(key).names.add(String(name).trim());
      if (src)  map.get(key).sources.add(String(src).trim());
      if (!src) map.get(key).sources.add("Manual"); // 沒來源視為手動
    }
  }

  // 抓 ICS 並解析
  const txt = UrlFetchApp.fetch(ICS_URL, {muteHttpExceptions: true}).getContentText();
  const lines = txt.split(/\r?\n/);
  let inEvent = false, dt = null, name = null;

  for (const raw of lines) {
    const line = raw.trim();
    if (line === "BEGIN:VEVENT") { inEvent = true; dt = null; name = null; continue; }
    if (line === "END:VEVENT") {
      if (dt) {
        const y = Number(dt.slice(0,4));
        if (y >= YEAR_FROM && y <= YEAR_TO) {
          const dateObj = new Date(y, Number(dt.slice(4,6)) - 1, Number(dt.slice(6,8)));
          const key = Utilities.formatDate(dateObj, tz, "yyyy-MM-dd");
          if (!map.has(key)) map.set(key, { date: dateObj, names: new Set(), sources: new Set() });
          if (name) map.get(key).names.add(name.trim());
          map.get(key).sources.add(SRC_ICS);
        }
      }
      inEvent = false; continue;
    }
    if (!inEvent) continue;
    if (line.startsWith("DTSTART")) {
      const m = line.match(/:(\d{8})/);
      if (m) dt = m[1];
    }
    if (line.startsWith("SUMMARY:")) {
      name = line.replace(/^SUMMARY:/, "");
    }
  }

  // 轉陣列、排序
  const out = [...map.values()]
    .sort((a, b) => a.date - b.date)
    .map(o => [
      o.date,
      [...o.names].join("／"),
      [...o.sources].join("／")
    ]);

  // 清除舊資料（保留表頭）並寫回
  if (last >= 2) sh.getRange(2,1,last-1,3).clearContent();
  if (out.length) {
    sh.getRange(2,1,out.length,3).setValues(out);
    sh.getRange(2,1,out.length,1).setNumberFormat("yyyy/MM/dd");
  }

  // 命名範圍 HOLIDAYS 指向 A 欄
  const named = ss.getNamedRanges().find(n => n.getName() === "HOLIDAYS");
  if (named) named.remove();
  ss.setNamedRange("HOLIDAYS", ss.getRange(`${SHEET}!A:A`));
}


/**
 * 建立一個「每年一次」的觸發器
 * 例如：每年 1 月 1 日 上午 7 點執行 syncTaiwanHolidays()
 */
function createYearlyTrigger() {
  // 先刪掉舊的同名觸發器（避免重複）
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "syncTaiwanHolidays")
    .forEach(t => ScriptApp.deleteTrigger(t));

  // 建立新的時間驅動觸發器
  ScriptApp.newTrigger("syncTaiwanHolidays")
    .timeBased()
    .onMonthDay(1)      // 每月第 1 天（這裡會在1月1號觸發）
    .inMonth(1)         // 第 1 月 = 1 月
    .atHour(7)          // 上午 7 點執行
    .create();
}

