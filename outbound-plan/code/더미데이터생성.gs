/**
 * 발주서 형태를 흉내낸 테스트 파일("YYYY-MM 발주서(테스트)")을 별도 구글시트로 만들어
 * 드라이브의 전용 테스트 폴더에 저장하고, 그 폴더 ID를 스크립트 속성 ORDER_FOLDER_ID에 자동
 * 등록한다 — 발주서가 출고계획 결과 파일과 분리된 별도 파일이어야 하기 때문.
 * 등록까지 자동으로 되므로 이 함수 실행 후 바로 generatePlan()을 돌리면 이 테스트 파일을 읽는다.
 *
 * 시트 1행에 '딥다이브' 마커, 2행에 헤더(품목코드/규격/중량/납기일), 3행부터 데이터.
 * 시트명은 '고객사 A' 형태로 지어야 vendor 인식(/고객사\s*([A-Za-z])/)이 된다.
 *
 * 품목코드/규격은 임의값이 아니라 기준정보.xlsx(CAPA/최대설비가동수/기초재고/적정재고)와
 * 제품 중량 (1).xlsx에 공통으로 들어있는 실제 품목코드 중에서 골랐다 — REFERENCE_FOLDER_ID를
 * 연동했을 때 단위중량 조회(ReferenceFiles.gs)와 생산capa 검증(ProductionCapa.gs)이 실제로
 * 매칭돼서 동작하는지까지 같이 확인하기 위함이다.
 */
function createTestOrderStatus() {
  const today = new Date();
  const y = today.getFullYear();
  const m = today.getMonth(); // 0-based
  const fileName = Utilities.formatString('%d-%02d 발주서(테스트)', y, m + 1);

  // 지금 VENDOR_TYPE_MAP은 고객사A=monday_bucket, 고객사B=wednesday_bucket(둘 다 트럭버킷),
  // 고객사C=friday_even(단순 균등분배)이다. 어떤 타입도 납기일을 안 쓰므로(date_as_is를 쓰는 업체가
  // 없음) 납기일 칸은 전부 비워둔다.
  //
  // 규격을 여러 개 섞고, 업체끼리 같은 품목코드를 겹쳐서도 넣었다(1900190을 A/B 둘 다 씀) —
  // 같은 품목코드를 여러 업체가 같이 쓸 때 codeDateLoad로 생산capa를 전사 공통으로 나눠 쓰는
  // 로직까지 같이 확인하기 위함이다.
  //
  // 수량은 전부 "보빈 개수 × 제품 중량 (1).xlsx의 단위중량(중량(KG))" 정확한 배수로 맞췄다
  // (단위중량: 7000260=417, 7000320=422, 1600190=429, 1900160=425, 1900190=395, 2400190=427,
  //  3000173=394, 3700260=425).
  //
  // A/B사(트럭버킷, 생산capa 검증 대상) 품목은 기준정보.xlsx(CAPA×최대설비가동수로 일일capa,
  // 기초재고, 적정재고)로 계산한 "한 달 최대 생산 가능량" 대비 30~70% 수준으로 잡아서(1900190은
  // A+B 합산 기준) capa 부족 경고 없이 정상적인 발주량을 흉내냈다. C사는 friday_even이라 capa/
  // 단위중량 검증 대상이 아니라서 그냥 임의로 채웠다.
  const sheetsData = {
    // A사 = monday_bucket : 트럭버킷 대상, 3개 규격 (합계 46,275)
    '고객사 A': [
      ['7000260', '7/0.260', 35 * 417, ''], // 14,595 (capa 상한 약 47,900 중 약 30%)
      ['1900190', '19/0.190', 39 * 395, ''], // 15,405 (A+B 합산 기준 capa 상한 약 37,300 중 A 몫 41%)
      ['3000173', '30/0.173', 26 * 394, ''], // 10,244 (capa 상한 약 31,300 중 약 33%)
    ],
    // B사 = wednesday_bucket : 트럭버킷 대상, A사와 1900190 겹침 (합계 29,631)
    '고객사 B': [
      ['7000320', '7/0.320', 22 * 422, ''], // 9,284 (capa 상한 약 35,600 중 약 26%)
      ['2400190', '24/0.190', 31 * 427, ''], // 13,237 (capa 상한 약 30,600 중 약 43%)
      ['1900190', '19/0.190', 18 * 395, ''], // 7,110 (A의 15,405와 합쳐 22,515 — capa 상한 약 37,300 중 약 60%)
    ],
    // C사 = friday_even : capa/단위중량 검증 없음, 그 달 금요일에 단순 균등분배, 3개 규격 (합계 28,154)
    '고객사 C': [
      ['1600190', '16/0.190', 26 * 429, ''], // 11,154
      ['1900160', '19/0.160', 22 * 425, ''], // 9,350
      ['3700260', '37/0.260', 18 * 425, ''], // 7,650
    ],
  };
  // 전체 총중량 = 40,244(A) + 29,631(B) + 28,154(C) = 98,029kg(약 98톤, 오차 29kg=0.03%)

  const folder = _getReferenceFolder_();

  // 같은 이름의 기존 테스트 파일이 있으면 지우고 새로 만든다(재실행 시 계속 누적되지 않게).
  const existing = folder.getFilesByName(fileName);
  while (existing.hasNext()) existing.next().setTrashed(true);

  const ss = SpreadsheetApp.create(fileName);
  const file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file); // create()가 기본으로 넣어준 내 드라이브 루트에서는 뺌

  Object.keys(sheetsData).forEach(sheetName => {
    const sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1).setValue(CONFIG.DEEP_DIVE_MARKER); // 1행: 딥다이브 마커
    sheet.getRange(2, 1, 1, 4).setValues([['품목코드', '규격', '중량', '납기일']]).setFontWeight('bold');

    const rows = sheetsData[sheetName];
    sheet.getRange(3, 1, rows.length, 4).setValues(rows);
    sheet.getRange(3, 4, rows.length, 1).setNumberFormat('yyyy-MM-dd');
    sheet.autoResizeColumns(1, 4);
  });

  // create()가 기본으로 만들어준 빈 "시트1"을 지운다(다른 탭들을 먼저 만든 뒤라야 마지막 탭 삭제 제약에 안 걸림).
  const defaultSheet = ss.getSheets()[0];
  if (Object.keys(sheetsData).indexOf(defaultSheet.getName()) === -1) {
    ss.deleteSheet(defaultSheet);
  }

  // 방금 만든 폴더를 ORDER_FOLDER_ID로 자동 등록 — 코드에는 ID가 안 남고 스크립트 속성에만 저장됨.
  PropertiesService.getScriptProperties().setProperty('ORDER_FOLDER_ID', folder.getId());

  Logger.log('테스트 발주서 파일: ' + ss.getUrl());
  Logger.log('ORDER_FOLDER_ID로 자동 등록된 폴더: ' + folder.getId());
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '테스트 발주서 파일 생성 및 ORDER_FOLDER_ID 등록 완료: ' + fileName,
    '완료',
    10
  );
}

/** GWP 공유문서함(REFERENCE_FOLDER_ID) 폴더 자체를 반환한다 — 하위 폴더 안 만들고 그 안에 파일만 둠. */
function _getReferenceFolder_() {
  if (!CONFIG.REFERENCE_FOLDER_ID) {
    throw new Error(
      'REFERENCE_FOLDER_ID가 아직 없어서 공유문서함 위치를 몰라요. ' +
      '공휴일/기준정보 파일이 있는 그 폴더 ID를 먼저 스크립트 속성 REFERENCE_FOLDER_ID에 등록해주세요.'
    );
  }
  return DriveApp.getFolderById(CONFIG.REFERENCE_FOLDER_ID);
}
