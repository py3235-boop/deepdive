/**
 * 배송/근무일 계산에 쓰는 휴일 집합.
 *
 * 공휴일.xlsx(공개 공휴일, '월'/'일' 반복 표기 — 연도 없음) + 기준정보.xlsx의 '휴무' 탭(회사 자체
 * 휴무, 정확한 날짜 있음)을 합쳐서 하나의 휴일 Set으로 쓴다. 원래는 성격이 다른 값이지만, 이번
 * 버전에서는 "이 날은 배송/생산 근무일이 아니다"라는 같은 목적으로만 쓰기 때문에 하나로 합쳤다.
 * REFERENCE_FOLDER_ID가 없으면(참고파일 미연동) 빈 Set을 반환해서 공휴일 없이 그냥 진행한다.
 */
function loadHolidays_(year) {
  const set = new Set();
  if (!CONFIG.REFERENCE_FOLDER_ID) return set;

  try {
    const ss = getReferenceSpreadsheet_('공휴일');
    const sheet = ss.getSheets()[0];
    const values = sheet.getDataRange().getValues();
    const header = values[0].map(h => String(h).trim());
    const mIdx = header.indexOf('월');
    const dIdx = header.indexOf('일');
    if (mIdx !== -1 && dIdx !== -1) {
      values.slice(1).forEach(r => {
        if (r[mIdx] === '' || r[mIdx] == null) return;
        const date = new Date(year, Number(r[mIdx]) - 1, Number(r[dIdx]));
        set.add(dateKey_(date));
      });
    }
  } catch (e) {
    Logger.log('공휴일.xlsx 로드 실패(무시하고 진행): ' + e.message);
  }

  try {
    const ss = getReferenceSpreadsheet_('기준정보');
    const sheet = ss.getSheetByName('휴무');
    if (sheet) {
      const values = sheet.getDataRange().getValues();
      const header = values[0].map(h => String(h).trim());
      const dateIdx = header.indexOf('날짜');
      if (dateIdx !== -1) {
        values.slice(1).forEach(r => {
          const v = r[dateIdx];
          if (!v) return;
          const date = v instanceof Date ? v : new Date(v);
          if (date.getFullYear() === year) set.add(dateKey_(date));
        });
      }
    }
  } catch (e) {
    Logger.log("기준정보.xlsx '휴무' 탭 로드 실패(무시하고 진행): " + e.message);
  }

  return set;
}
