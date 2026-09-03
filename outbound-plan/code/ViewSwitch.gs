/**
 * "출고계획 보기" 메뉴 — 재계산 없이 마지막에 생성한 계획(LAST_PLAN)을 다른 배치로 다시 그린다.
 * generatePlan()이 매번 문서 속성(DocumentProperties)에 저장해둔 값을 읽어서 쓰므로,
 * 발주서를 다시 읽거나 트럭버킷 계산을 다시 돌리지 않는다(가이드 §5 "_rebuildPlanLayout" 방식).
 */

function showByVendor() {
  _rebuildPlanLayout_('vendor');
}

function showByItem() {
  _rebuildPlanLayout_('item');
}

function _rebuildPlanLayout_(viewMode) {
  const raw = PropertiesService.getDocumentProperties().getProperty('LAST_PLAN');
  if (!raw) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '저장된 계획이 없습니다. 먼저 "② 계획 생성"을 한 번 실행하세요.', '안내', 6
    );
    return;
  }

  const saved = JSON.parse(raw);
  const holidaySet = new Set(saved.holidayDates || []);

  writePlanSheet(saved.year, saved.month, saved.rows, {
    holidaySet: holidaySet,
    issueCount: saved.issueCount || 0,
    viewMode: viewMode,
    actualCutoffDateKey: saved.actualCutoffDateKey || null,
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(
    (viewMode === 'vendor' ? '회사별 보기로 전환' : '품목별 보기로 전환') + ' 완료', '완료', 5
  );
}
