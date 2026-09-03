/**
 * 실행이력 — generatePlan() 등을 실행할 때마다 [실행일시, 실행자, 함수, 결과, 메시지]를 한 줄씩
 * "실행이력" 탭(이 스프레드시트 안)에 누적 기록한다. 메시지 칸에는 검증 경고가 있으면 건수만이
 * 아니라 "어떤 품목이 왜 경고인지" 실제 내용까지 그대로 적어서, Apps Script 실행 로그(보기 →
 * 로그, 브라우저 닫으면 사라짐)를 안 봐도 시트에서 바로 확인할 수 있게 한다.
 */
function appendExecutionLog_(functionName, result, message) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = '실행이력';
  let sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['실행일시', '실행자', '함수', '결과', '메시지']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold');
    sheet.setColumnWidth(1, 140);
    sheet.setColumnWidth(2, 200);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 60);
    sheet.setColumnWidth(5, 600);
    sheet.setFrozenRows(1);
  }

  let executor;
  try {
    executor = Session.getActiveUser().getEmail() || '(이메일 비공개)';
  } catch (e) {
    executor = '(확인 불가)';
  }

  sheet.appendRow([
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    executor,
    functionName,
    result,
    message || '',
  ]);

  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 5).setWrap(true).setVerticalAlignment('top');
}
