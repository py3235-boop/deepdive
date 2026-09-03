/**
 * 공통 날짜 유틸. year/month는 전부 1-based 월(1~12)을 쓴다.
 */

function dateKey_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function daysInMonth_(year, month) {
  return new Date(year, month, 0).getDate(); // month(1-based)의 마지막 날
}

function formatDayLabel_(day) {
  return day + '일';
}

/**
 * 그 달에서 특정 요일(0=일 ... 6=토)에 해당하는 날짜를 전부 반환한다.
 */
function getWeekdayDatesInMonth_(year, month, weekday) {
  const total = daysInMonth_(year, month);
  const dates = [];
  for (let d = 1; d <= total; d++) {
    const date = new Date(year, month - 1, d);
    if (date.getDay() === weekday) dates.push(date);
  }
  return dates;
}
