/**
 * 세인트 다운로드 파일(.xls, 실제로는 HTML 테이블) 파서 (스펙 2장)
 * - UTF-8 텍스트로 읽힌 HTML 문자열에서 <tr>/<td> 를 추출해 행 객체 배열로 변환
 * - 헤더 행은 '참조전표'와 '개별항목' 컬럼이 함께 존재하는 행으로 식별
 */

/**
 * @param {string} content HTML 원문
 * @return {Object[]} SRC_COLS 를 키로 갖는 행 객체 배열 (금액/일자는 정규화됨)
 */
function parseSaintHtml_(content) {
  if (!content || typeof content !== 'string') {
    throw new Error('파일 내용이 비어 있습니다.');
  }
  var trs = content.match(/<tr[\s\S]*?<\/tr>/gi) || [];
  if (!trs.length) {
    throw new Error('HTML 테이블(<tr>)을 찾을 수 없습니다. 세인트에서 받은 .xls 파일이 맞는지 확인해 주세요.');
  }

  var headerMap = null; // SRC_COLS 각 컬럼명 → 셀 인덱스
  var rows = [];

  for (var i = 0; i < trs.length; i++) {
    var cells = extractCells_(trs[i]);
    if (!cells.length) continue;

    if (isHeaderRow_(cells)) {
      // 페이지마다 헤더가 반복될 수 있으므로 매번 갱신하고 데이터로는 취급하지 않음
      headerMap = buildHeaderMap_(cells);
      continue;
    }
    if (!headerMap) continue; // 헤더 이전의 제목/요약 행 무시

    var src = {};
    var missing = [];
    for (var c = 0; c < SRC_COLS.length; c++) {
      var col = SRC_COLS[c];
      var idx = headerMap[col];
      if (idx === undefined || idx >= cells.length) {
        missing.push(col);
        src[col] = '';
      } else {
        src[col] = cells[idx];
      }
    }
    // 키가 없는 행(합계/빈 행)은 스킵
    if (!normStr_(src['참조전표']) && !normStr_(src['개별항목'])) continue;

    // 정규화
    src['사용금액'] = parseAmount_(src['사용금액']);
    src['전표일자'] = parseDateStr_(src['전표일자']);
    src['지급일자'] = parseDateStr_(src['지급일자']);
    ['수입/지출', '자금', '자금명', '약정항목', '약정항목 명', '전표상태', '전표유형명', '참조전표', '개별항목', '텍스트']
      .forEach(function (k) { src[k] = normStr_(src[k]); });

    rows.push(src);
  }

  if (!headerMap) {
    throw new Error('헤더 행(참조전표/개별항목 컬럼 포함)을 찾을 수 없습니다. 파일 형식을 확인해 주세요.');
  }
  return rows;
}

function isHeaderRow_(cells) {
  var norm = cells.map(normColName_);
  return norm.indexOf('참조전표') >= 0 && norm.indexOf('개별항목') >= 0;
}

/** 헤더 셀 배열 → SRC_COLS 각 컬럼의 셀 인덱스 매핑 */
function buildHeaderMap_(cells) {
  var norm = cells.map(normColName_);
  var map = {};
  var missing = [];
  for (var i = 0; i < SRC_COLS.length; i++) {
    var want = normColName_(SRC_COLS[i]);
    var idx = norm.indexOf(want);
    if (idx < 0) missing.push(SRC_COLS[i]);
    else map[SRC_COLS[i]] = idx;
  }
  if (missing.length) {
    throw new Error('필수 컬럼을 찾을 수 없습니다: ' + missing.join(', ') +
      '\n(파일의 실제 헤더: ' + cells.join(' | ') + ')');
  }
  return map;
}

/** 컬럼명 비교용 정규화: 모든 공백 제거 ('약정항목 명' ↔ '약정항목명' 허용) */
function normColName_(s) {
  return normStr_(s).replace(/\s+/g, '');
}

/** <tr> 문자열 → 텍스트 셀 배열 */
function extractCells_(tr) {
  var matches = tr.match(/<t[dh][^>]*>[\s\S]*?<\/t[dh]>/gi) || [];
  return matches.map(function (cell) {
    var inner = cell.replace(/^<t[dh][^>]*>/i, '').replace(/<\/t[dh]>$/i, '');
    inner = inner.replace(/<[^>]+>/g, ' ');
    return decodeHtmlEntities_(inner).replace(/\s+/g, ' ').trim();
  });
}

function decodeHtmlEntities_(s) {
  return s
    .replace(/&nbsp;/gi, ' ')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&#(\d+);/g, function (_, d) { return String.fromCharCode(Number(d)); })
    .replace(/&amp;/gi, '&');
}
