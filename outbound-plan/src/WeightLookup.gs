/**
 * '제품 중량' 참고파일에서 품목코드/규격별 단위중량(보빈 1개 기준 등)을 읽어온다.
 * 코드로 못 찾으면 규격 문자열로 한 번 더 시도한다(`lookupUnitWeight_`).
 */
function loadUnitWeights_() {
  const weights = { byCode: {}, bySpec: {} };
  if (!CONFIG.REFERENCE_FOLDER_ID) return weights;

  try {
    const ss = getReferenceSpreadsheet_('제품 중량');
    const sheet = ss.getSheets()[0];
    const values = sheet.getDataRange().getValues();
    const header = values[0].map(h => String(h).trim());

    const codeIdx = findColumnIndex_(header, ['품목코드', '제품코드', 'sc코드']);
    const specIdx = findColumnIndex_(header, ['규격', '사양']);
    const weightIdx = findColumnIndex_(header, ['평균중량', '단위중량', '중량']);
    if (weightIdx === -1) return weights;

    values.slice(1).forEach(r => {
      const w = Number(r[weightIdx]);
      if (!w) return;
      if (codeIdx !== -1 && r[codeIdx]) {
        const code = String(r[codeIdx]).trim();
        weights.byCode[code] = w;
        const stripped = code.replace(/^[A-Za-z]+/, '');
        if (stripped && stripped !== code) weights.byCode[stripped] = w;
      }
      if (specIdx !== -1 && r[specIdx]) {
        weights.bySpec[String(r[specIdx]).trim()] = w;
      }
    });
  } catch (e) {
    Logger.log('제품 중량 참고파일 로드 실패(단위중량 보정 없이 진행): ' + e.message);
  }

  return weights;
}

function lookupUnitWeight_(weights, code, spec) {
  if (weights.byCode[code] != null) return weights.byCode[code];
  const stripped = String(code).replace(/^[A-Za-z]+/, '');
  if (weights.byCode[stripped] != null) return weights.byCode[stripped];
  if (spec && weights.bySpec[spec] != null) return weights.bySpec[spec];
  return null;
}
