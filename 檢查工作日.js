/**
 * 計算 offset 並避開週末
 * @param {Date} baseDate 起始日
 * @param {number} days 增加天數
 */
function calculateTargetDate(baseDate, days) {
  const result = new Date(baseDate);
  result.setDate(result.getDate() + days);
  
  const day = result.getDay();
  if (day === 6) result.setDate(result.getDate() + 2); // Sat -> Mon
  if (day === 0) result.setDate(result.getDate() + 1); // Sun -> Mon
  
  return result;
}