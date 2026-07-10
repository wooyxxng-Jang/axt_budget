/**
 * 검토대기함 (스펙 3-2장)
 * - 교수성함/타학과 키워드가 둘 다 없거나 충돌하는 건
 * - 사람이 '판정' 컬럼에 포함/제외 선택 → 메뉴 [검토대기함 → 확정 반영]
 * - 처리된 행은 상태=반영완료로 남겨 영구 저장 (재업로드 시 다시 묻지 않음)
 */

/** @param {Array<{src: Object, reason: string}>} items */
function appendReviewRows_(items) {
  if (!items.length) return;
  var sh = ensureSheet_(SHEET.REVIEW, REVIEW_HEADERS);
  var rows = items.map(function (it) {
    return srcToRowPrefix_(it.src).concat([it.reason || '', '', STATUS.PENDING, '']);
  });
  var startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rows.length, REVIEW_HEADERS.length).setValues(rows);
}

/** 메뉴: 판정(포함/제외)이 입력된 대기 건을 원장/감사로그에 반영 */
function applyReviewDecisions() {
  var ui = SpreadsheetApp.getUi();
  var sh = ensureSheet_(SHEET.REVIEW, REVIEW_HEADERS);
  var last = sh.getLastRow();
  if (last < 2) { ui.alert('검토대기함에 처리할 건이 없습니다.'); return; }

  var width = REVIEW_HEADERS.length;
  var colVerdict = REVIEW_HEADERS.indexOf('판정');
  var colStatus = REVIEW_HEADERS.indexOf('상태');
  var colTime = REVIEW_HEADERS.indexOf('처리일시');

  var data = sh.getRange(2, 1, last - 1, width).getValues();
  var now = nowStr_();

  var ledgerItems = [], auditItems = [], detailItems = [];
  var included = 0, excluded = 0, toDetail = 0;

  data.forEach(function (row) {
    if (normStr_(row[colStatus]) !== STATUS.PENDING) return;
    var verdict = normStr_(row[colVerdict]);
    if (VERDICT_VALUES.indexOf(verdict) < 0) return;

    var src = rowPrefixToSrc_(row);
    if (verdict === '포함') {
      var route = routeDetail_(src);
      if (route.resolved) {
        ledgerItems.push({ src: src, detail: route.detail, reason: '검토확정(포함)', path: PATH.REVIEW });
        row[colStatus] = STATUS.DONE;
        included++;
      } else {
        detailItems.push({ src: src, reason: '검토확정(포함)', candidates: route.candidates });
        row[colStatus] = STATUS.TO_DETAIL;
        toDetail++;
      }
    } else {
      auditItems.push({ src: src, reason: '검토확정(제외)' });
      row[colStatus] = STATUS.DONE;
      excluded++;
    }
    row[colTime] = now;
  });

  if (!included && !excluded && !toDetail) {
    ui.alert('판정(포함/제외)이 입력된 대기 건이 없습니다.\n판정 컬럼을 선택한 뒤 다시 실행해 주세요.');
    return;
  }

  sh.getRange(2, 1, data.length, width).setValues(data);
  upsertLedgerRows_(ledgerItems);
  appendAuditRows_(auditItems, null);
  appendDetailQueueRows_(detailItems);
  rebuildMonthlySummary();

  ui.alert('검토대기함 반영 완료\n\n포함(원장): ' + included + '건\n제외(감사로그): ' + excluded +
    '건\n상세분류 확인 필요: ' + toDetail + '건');
}

/** 시트 행(업서트키 + SRC_COLS 순) → 소스 객체 복원 */
function rowPrefixToSrc_(row) {
  var src = {};
  for (var i = 0; i < SRC_COLS.length; i++) {
    src[SRC_COLS[i]] = row[i + 1];
  }
  src['사용금액'] = parseAmount_(src['사용금액']);
  src['전표일자'] = parseDateStr_(src['전표일자']);
  src['지급일자'] = parseDateStr_(src['지급일자']);
  return src;
}
