/**
 * 30_Extract.gs — 형식별 추출 계층
 *
 * 단일 진입점 추출(입력) 이 형식만 보고 분기한다. 고객사는 여기서 신경쓰지 않는다.
 *
 * 이 계층은 **해석하지 않는다.** 표가 있으면 표로, 없으면 텍스트 줄로 "있는 그대로" 내보내고
 * 고객사별 해석은 40_Parse.gs 가 한다. 이미지 OCR 은 표 구조를 잃고 줄 텍스트로만 나오는
 * 경우가 많으므로 표와 줄을 **둘 다** 반환한다.
 *
 * 교체 지점: 정확도가 부족한 형식이 나오면 아래 _추출_XXX 함수 하나만 갈아끼우면 된다
 * (예: 이미지 경로를 LLM 비전으로). 상위 계층은 손대지 않는다.
 */

/** 지원 확장자 → 내부 형식 이름 */
const 형식표 = {
  docx: 'docx', doc: 'docx',
  xlsx: 'xlsx', xls: 'xlsx',
  pdf: 'pdf',
  png: '이미지', jpg: '이미지', jpeg: '이미지', gif: '이미지', bmp: '이미지', webp: '이미지',
  csv: 'csv', txt: 'csv',
};

/** 한글 파일은 Drive 가 변환하지 못한다. 명시적으로 거른다. */
const 미지원표 = { hwp: '한글(.hwp)', hwpx: '한글(.hwpx)' };

/** 추출 결과의 빈 껍데기 */
function _추출결과(형식, 메타) {
  return {
    성공: false,
    형식: 형식 || '?',
    표들: [],      // string[][][]  — 표 단위. 진짜 표가 있을 때만 채워진다
    줄들: [],      // string[]      — 텍스트 줄
    원문: '',
    메타:메타 || {},
  };
}

/**
 * 단일 진입점.
 *
 * @param {Object} 입력
 *   - blob      {Blob}   내용. 파일 경로면 필수
 *   - 파일명    {string} 확장자 판정에 쓴다
 *   - html      {string} 메일 본문 HTML (blob 대신 이걸 주면 html 경로로 간다)
 *   - 원본ID    {string} 드라이브 파일 ID (있으면 로그에 남긴다)
 * @return {Object} _추출결과 형태
 */
function 추출(입력) {
  입력 = 입력 || {};
  const 파일명 = String(입력.파일명 || '');
  const 메타 = { 파일명: 파일명, 원본ID: 입력.원본ID || '', ocr: false };

  // 메일 본문 HTML 경로
  if (입력.html != null && 입력.html !== '') {
    return _추출_html(입력.html, 메타);
  }

  if (!입력.blob) {
    const r = _추출결과('?', 메타);
    r.메타.실패사유 = '내용(blob)도 html 도 없습니다';
    return r;
  }

  const 확장자 = 파일명.indexOf('.') >= 0
    ? 파일명.split('.').pop().toLowerCase()
    : '';

  if (미지원표[확장자]) {
    const r = _추출결과(확장자, 메타);
    r.메타.실패사유 = 미지원표[확장자] + ' 은 지원하지 않습니다 (Drive 가 변환하지 못함)';
    return r;
  }

  const 형식 = 형식표[확장자];
  if (!형식) {
    const r = _추출결과(확장자 || '?', 메타);
    r.메타.실패사유 = '지원하지 않는 확장자: ' + (확장자 || '(없음)');
    return r;
  }

  try {
    if (형식 === 'docx') return _추출_문서변환(입력.blob, 메타, false);
    if (형식 === 'pdf') return _추출_pdf(입력.blob, 메타);
    if (형식 === '이미지') return _추출_문서변환(입력.blob, 메타, true);
    if (형식 === 'xlsx') return _추출_스프레드시트(입력.blob, 메타);
    if (형식 === 'csv') return _추출_csv(입력.blob, 메타);
  } catch (e) {
    const r = _추출결과(형식, 메타);
    r.메타.실패사유 = String(e && e.message ? e.message : e);
    return r;
  }

  const r = _추출결과(형식, 메타);
  r.메타.실패사유 = '분기 누락: ' + 형식;
  return r;
}

// ─────────────────────────────────────────────────────────────
// Drive 변환 공통
// ─────────────────────────────────────────────────────────────

/**
 * blob 을 구글 문서로 변환한다. 변환 결과 파일 ID 를 돌려준다.
 * 호출한 쪽이 반드시 _임시파일삭제 로 지워야 한다.
 *
 * ⚠️ OCR 경로에서는 목표 mimeType 을 지정하면 Drive 가 거부한다:
 *    "OCR is not supported for files of type application/vnd.google-apps.document"
 *    OCR 을 켤 때는 title 만 주고 형식은 Drive 가 정하게 둔다.
 */
function _구글문서로변환(blob, ocr) {
  const 자원 = { title: '_임시변환_' + new Date().getTime() };

  if (ocr) {
    return Drive.Files.insert(자원, blob, { ocr: true, ocrLanguage: 'ko' }).id;
  }

  자원.mimeType = MimeType.GOOGLE_DOCS;
  return Drive.Files.insert(자원, blob, { convert: true, supportsAllDrives: true }).id;
}

/** 변환 산출물을 지운다. 실패해도 조용히 넘어간다 (드라이브 오염만 남을 뿐이므로). */
function _임시파일삭제(파일ID) {
  if (!파일ID) return;
  try {
    DriveApp.getFileById(파일ID).setTrashed(true);
  } catch (e) {
    Logger.log('임시파일 삭제 실패(무시): ' + 파일ID + ' — ' + e);
  }
}

/** 구글 문서에서 표와 텍스트를 뽑는다. */
function _문서읽기(문서ID) {
  const body = DocumentApp.openById(문서ID).getBody();

  const 표들 = body.getTables().map(function (t) {
    const 격자 = [];
    for (let r = 0; r < t.getNumRows(); r++) {
      const row = t.getRow(r);
      const 행 = [];
      for (let c = 0; c < row.getNumCells(); c++) {
        행.push(row.getCell(c).getText().replace(/\s+/g, ' ').trim());
      }
      격자.push(행);
    }
    return 격자;
  });

  const 원문 = body.getText();
  const 줄들 = 원문.split(/\r?\n/).map(function (s) { return s.trim(); })
    .filter(function (s) { return s !== ''; });

  return { 표들: 표들, 줄들: 줄들, 원문: 원문 };
}

/** docx / 이미지 → 구글 문서 변환 → 표·텍스트 */
function _추출_문서변환(blob, 메타, ocr) {
  const 형식 = ocr ? '이미지' : 'docx';
  const 결과 = _추출결과(형식, 메타);
  결과.메타.ocr = !!ocr;

  let 문서ID = null;
  try {
    문서ID = _구글문서로변환(blob, ocr);
    const 읽음 = _문서읽기(문서ID);
    결과.표들 = 읽음.표들;
    결과.줄들 = 읽음.줄들;
    결과.원문 = 읽음.원문;
    결과.성공 = 읽음.표들.length > 0 || 읽음.줄들.length > 0;
    if (!결과.성공) 결과.메타.실패사유 = '변환은 됐지만 표도 텍스트도 없습니다';
  } finally {
    _임시파일삭제(문서ID);
  }
  return 결과;
}

/**
 * PDF → 구글 문서.
 *
 * 글자 정확도는 텍스트 레이어(OCR 없이)가 낫지만, **표 구조는 OCR 쪽이 살려주는 경우가 있다.**
 * 실측: 고객사C PDF 는 텍스트 레이어로 232자를 정확히 건졌지만 표는 0개였고,
 * 값이 여러 줄에 흩어진 채 `이 하 여 백` 같은 노이즈와 붙어 나왔다.
 *
 * 그래서 표가 나오는 쪽을 우선한다:
 *   1) OCR 없이 → 표가 있으면 그걸 쓴다 (정확도·구조 둘 다 최선)
 *   2) 표가 없으면 OCR 로 다시 → 표가 있으면 그걸 쓴다 (구조를 얻는다)
 *   3) 둘 다 표가 없으면 글자를 더 많이 건진 쪽 (스캔 PDF 는 OCR 만 건진다)
 */
function _추출_pdf(blob, 메타) {
  const 없이 = _추출_문서변환(blob, Object.assign({}, 메타), false);
  없이.형식 = 'pdf';
  없이.메타.ocr = false;
  const 없이글자 = 없이.원문.replace(/\s+/g, '').length;

  if (없이.표들.length > 0) {
    없이.메타.판단 = '텍스트 레이어에서 표 ' + 없이.표들.length + '개 확보 (문자 ' + 없이글자 + '자)';
    return 없이;
  }

  let OCR = null;
  try {
    OCR = _추출_문서변환(blob, Object.assign({}, 메타), true);
    OCR.형식 = 'pdf';
    OCR.메타.ocr = true;
  } catch (e) {
    없이.메타.판단 = '텍스트 레이어만 사용 (표 없음, 문자 ' + 없이글자 + '자). OCR 시도 실패: ' + e;
    return 없이;
  }

  const OCR글자 = OCR.원문.replace(/\s+/g, '').length;

  if (OCR.표들.length > 0) {
    OCR.메타.판단 = 'OCR 에서 표 ' + OCR.표들.length + '개 확보 (텍스트 레이어는 표 0개)';
    return OCR;
  }

  if (없이글자 >= 20 && 없이글자 >= OCR글자) {
    없이.메타.판단 = '둘 다 표 없음 → 텍스트 레이어 채택 (레이어 ' + 없이글자 + '자 vs OCR ' + OCR글자 + '자)';
    return 없이;
  }

  OCR.메타.판단 = '둘 다 표 없음 → OCR 채택 (레이어 ' + 없이글자 + '자 vs OCR ' + OCR글자 + '자)';
  return OCR;
}

/** xlsx → 구글 시트 변환 → 각 탭을 표로 */
function _추출_스프레드시트(blob, 메타) {
  const 결과 = _추출결과('xlsx', 메타);

  let 시트ID = null;
  try {
    const 파일 = Drive.Files.insert(
      { title: '_임시변환_' + new Date().getTime(), mimeType: MimeType.GOOGLE_SHEETS },
      blob,
      { convert: true, supportsAllDrives: true }
    );
    시트ID = 파일.id;

    const ss = SpreadsheetApp.openById(시트ID);
    ss.getSheets().forEach(function (s) {
      if (s.getLastRow() === 0) return;
      결과.표들.push(s.getDataRange().getValues().map(function (행) {
        return 행.map(function (v) { return String(v == null ? '' : v).trim(); });
      }));
    });

    결과.줄들 = 결과.표들.reduce(function (acc, 표) {
      표.forEach(function (행) {
        const s = 행.join('\t').trim();
        if (s) acc.push(s);
      });
      return acc;
    }, []);
    결과.원문 = 결과.줄들.join('\n');
    결과.성공 = 결과.표들.length > 0;
    if (!결과.성공) 결과.메타.실패사유 = '변환은 됐지만 데이터가 있는 탭이 없습니다';
  } finally {
    _임시파일삭제(시트ID);
  }
  return 결과;
}

/** csv / txt → 직접 파싱 */
function _추출_csv(blob, 메타) {
  const 결과 = _추출결과('csv', 메타);
  const 문자열 = blob.getDataAsString('UTF-8');
  결과.원문 = 문자열;
  결과.줄들 = 문자열.split(/\r?\n/).map(function (s) { return s.trim(); })
    .filter(function (s) { return s !== ''; });
  try {
    결과.표들 = [Utilities.parseCsv(문자열).map(function (행) {
      return 행.map(function (v) { return String(v == null ? '' : v).trim(); });
    })];
  } catch (e) {
    결과.메타.비고 = 'CSV 파싱 실패, 줄 텍스트만 제공: ' + e;
  }
  결과.성공 = 결과.줄들.length > 0;
  return 결과;
}

// ─────────────────────────────────────────────────────────────
// 메일 본문 HTML
// ─────────────────────────────────────────────────────────────

/**
 * 메일 본문 HTML 에서 <table> 을 뽑는다. **고객사 A 의 주 경로다.**
 *
 * Apps Script 에는 DOM 파서가 없고 메일 HTML 은 잘 깨져 있어서 XmlService 를 못 쓴다.
 * 그래서 정규식으로 행·셀을 훑는다. 중첩 표는 안쪽까지 각각 하나의 표로 잡힌다
 * (레이아웃용 표를 걸러내는 일은 40_Parse.gs 가 컬럼을 보고 판단한다).
 */
function _추출_html(html, 메타) {
  const 결과 = _추출결과('html', 메타);
  const 원본 = String(html);

  const 표정규식 = /<table[^>]*>([\s\S]*?)<\/table>/gi;
  let m;
  while ((m = 표정규식.exec(원본)) !== null) {
    const 표 = _html표파싱(m[1]);
    if (표.length) 결과.표들.push(표);
  }

  결과.원문 = _html에서텍스트(원본);
  결과.줄들 = 결과.원문.split(/\r?\n/).map(function (s) { return s.trim(); })
    .filter(function (s) { return s !== ''; });

  결과.성공 = 결과.표들.length > 0 || 결과.줄들.length > 0;
  if (!결과.성공) 결과.메타.실패사유 = '본문에서 표도 텍스트도 찾지 못했습니다';
  결과.메타.표개수 = 결과.표들.length;
  return 결과;
}

/** <tr>/<td> 를 훑어 2차원 배열로. 중첩 <table> 안쪽은 셀 텍스트로 눌러 담는다. */
function _html표파싱(내부) {
  const 행들 = [];
  const 행정규식 = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  let mr;
  while ((mr = 행정규식.exec(내부)) !== null) {
    const 셀들 = [];
    const 셀정규식 = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
    let mc;
    while ((mc = 셀정규식.exec(mr[1])) !== null) {
      셀들.push(_html에서텍스트(mc[1]).replace(/\s+/g, ' ').trim());
    }
    if (셀들.length) 행들.push(셀들);
  }
  return 행들;
}

/** 태그를 벗기고 엔티티를 풀어 텍스트로. <br>/</tr>/</p> 는 줄바꿈으로 바꾼다. */
function _html에서텍스트(html) {
  return String(html)
    .replace(/<(script|style)[^>]*>[\s\S]*?<\/\1>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|tr|h[1-6]|li)>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/&#(\d+);/g, function (_, n) { return String.fromCharCode(Number(n)); })
    .replace(/ /g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n{3,}/g, '\n\n');
}

// ─────────────────────────────────────────────────────────────
// 드라이브 파일 편의 함수
// ─────────────────────────────────────────────────────────────

/** 드라이브 파일 ID 로 추출한다. */
function 추출_드라이브파일(파일ID) {
  const f = DriveApp.getFileById(파일ID);
  return 추출({ blob: f.getBlob(), 파일명: f.getName(), 원본ID: 파일ID });
}

/** 폴더 안의 파일들을 전부 추출한다. */
function 추출_폴더(폴더ID) {
  const it = DriveApp.getFolderById(폴더ID).getFiles();
  const 목록 = [];
  while (it.hasNext()) {
    const f = it.next();
    목록.push({ 파일: f, 결과: 추출({ blob: f.getBlob(), 파일명: f.getName(), 원본ID: f.getId() }) });
  }
  return 목록;
}

// ─────────────────────────────────────────────────────────────
// 검증
// ─────────────────────────────────────────────────────────────

/**
 * Phase 2 검증용. 고객사 첨부파일 폴더의 파일을 전부 추출해 원시 결과를 로그로 뿜는다.
 * 여기서 이미지 OCR 정확도를 눈으로 실측한다.
 */
function 추출_테스트() {
  const 줄 = [];
  function p(s) { 줄.push(s); Logger.log(s); }

  고객사목록.forEach(function (고객사) {
    const 폴더ID = 설정전체()['폴더.' + 고객사];
    p('');
    p('##################  ' + 고객사 + '  ##################');
    if (!폴더ID) { p('  폴더 ID 설정 없음'); return; }

    const 목록 = 추출_폴더(폴더ID);
    if (!목록.length) { p('  (폴더가 비어 있습니다)'); return; }

    목록.forEach(function (항목) {
      const r = 항목.결과;
      p('');
      p('── ' + 항목.파일.getName());
      p('   형식=' + r.형식 + ' / 성공=' + r.성공 + ' / OCR=' + r.메타.ocr +
        (r.메타.판단 ? ' / ' + r.메타.판단 : '') +
        (r.메타.실패사유 ? ' / 실패=' + r.메타.실패사유 : ''));
      p('   표 ' + r.표들.length + '개, 줄 ' + r.줄들.length + '개, 원문 ' +
        r.원문.replace(/\s+/g, '').length + '자');

      r.표들.forEach(function (표, ti) {
        p('   [표 ' + (ti + 1) + '] ' + 표.length + '행 × ' +
          (표[0] ? 표[0].length : 0) + '열');
        표.slice(0, 15).forEach(function (행, ri) {
          p('      ' + (ri + 1) + ': ' + 행.join(' | '));
        });
        if (표.length > 15) p('      … 이하 ' + (표.length - 15) + '행 생략');
      });

      if (!r.표들.length) {
        p('   [표 없음 — 줄 텍스트만]');
        r.줄들.slice(0, 30).forEach(function (s, i) { p('      ' + (i + 1) + ': ' + s); });
        if (r.줄들.length > 30) p('      … 이하 ' + (r.줄들.length - 30) + '줄 생략');
      }
    });
  });

  return 줄.join('\n');
}

/** 시트 메뉴용 래퍼 */
function 추출_테스트_메뉴() {
  const 결과 = 추출_테스트();
  const ui = SpreadsheetApp.getUi();
  ui.alert('추출 테스트', 결과.slice(0, 4000), ui.ButtonSet.OK);
}
