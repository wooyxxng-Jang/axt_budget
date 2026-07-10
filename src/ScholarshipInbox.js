/**
 * 장학금 분류 대기함 (스펙 5장, 4장 A)
 * - 전표유형=등록금 건 중 신규만 노출
 * - 사람이 귀속(아텍/타학과) + 해당 시 상세분류(세부항목) 태깅
 * - 메뉴 [장학금 태깅 → 반영]으로 원장/감사로그에 반영, 행은 반영완료로 남겨 영구 저장
 */

/** @param {Array<{src: Object}>} items */
function appendScholarshipRows_(items) {
  if (!items.length) return;
  var sh = ensureSheet_(SHEET.SCHOLARSHIP, SCHOLARSHIP_HEADERS);
  var colDetail = SCHOLARSHIP_HEADERS.indexOf('상세분류') + 1; // 1-based

  var rows = items.map(function (it) {
    return srcToRowPrefix_(it.src).concat(['', '', STATUS.PENDING, '']);
  });
  var startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rows.length, SCHOLARSHIP_HEADERS.length).setValues(rows);

  // 행별 상세분류 드롭다운 (해당 자금+약정항목 그룹의 마스터 후보)
  items.forEach(function (it, i) {
    var candidates = getDetailCandidates_(it.src['자금'], it.src['약정항목']);
    if (candidates.length > 1) {
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(candidates, true)
        .setAllowInvalid(false)
        .build();
      sh.getRange(startRow + i, colDetail).setDataValidation(rule);
    }
  });
}

/** 메뉴: 귀속이 태깅된 대기 건을 원장/감사로그에 반영 */
function applyScholarshipTags() {
  var ui = SpreadsheetApp.getUi();
  var sh = ensureSheet_(SHEET.SCHOLARSHIP, SCHOLARSHIP_HEADERS);
  var last = sh.getLastRow();
  if (last < 2) { ui.alert('장학금 분류 대기함에 처리할 건이 없습니다.'); return; }

  var width = SCHOLARSHIP_HEADERS.length;
  var colTag = SCHOLARSHIP_HEADERS.indexOf('귀속');
  var colDetail = SCHOLARSHIP_HEADERS.indexOf('상세분류');
  var colStatus = SCHOLARSHIP_HEADERS.indexOf('상태');
  var colTime = SCHOLARSHIP_HEADERS.indexOf('처리일시');

  var data = sh.getRange(2, 1, last - 1, width).getValues();
  var now = nowStr_();

  var ledgerItems = [], auditItems = [];
  var included = 0, excluded = 0, needDetail = 0;

  data.forEach(function (row) {
    if (normStr_(row[colStatus]) !== STATUS.PENDING) return;
    var tag = normStr_(row[colTag]);
    if (TAG_VALUES.indexOf(tag) < 0) return;

    var src = rowPrefixToSrc_(row);
    if (tag === '아텍') {
      var candidates = getDetailCandidates_(src['자금'], src['약정항목']);
      var detail = normStr_(row[colDetail]);
      if (candidates.length > 1 && !detail) {
        needDetail++; // 상세분류 미선택 → 대기 유지
        return;
      }
      if (candidates.length === 1 && !detail) detail = candidates[0];
      ledgerItems.push({ src: src, detail: detail, reason: '장학금태깅(아텍)', path: PATH.SCHOLARSHIP });
      included++;
    } else {
      auditItems.push({ src: src, reason: '장학금태깅(타학과)' });
      excluded++;
    }
    row[colStatus] = STATUS.DONE;
    row[colTime] = now;
  });

  if (!included && !excluded) {
    var msg = '귀속(아텍/타학과)이 태깅된 대기 건이 없습니다.';
    if (needDetail) msg += '\n\n상세분류 미선택으로 보류된 건: ' + needDetail + '건 (상세분류를 선택해 주세요)';
    ui.alert(msg);
    return;
  }

  sh.getRange(2, 1, data.length, width).setValues(data);
  upsertLedgerRows_(ledgerItems);
  appendAuditRows_(auditItems, null);
  rebuildMonthlySummary();

  var doneMsg = '장학금 태깅 반영 완료\n\n아텍(원장): ' + included + '건\n타학과(감사로그): ' + excluded + '건';
  if (needDetail) doneMsg += '\n상세분류 미선택 보류: ' + needDetail + '건';
  ui.alert(doneMsg);
}
