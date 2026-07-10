/**
 * 상세분류 확인 대기열 (스펙 4장 B, 6장)
 * - 아텍 귀속은 확정됐으나 상세분류(세부라인) 자동 라우팅에 실패한 건
 * - 후보 상세분류를 드롭다운으로 제시 → 사람이 원클릭 선택
 * - 메뉴 [상세분류 선택 → 반영]으로 원장 반영
 */

/** @param {Array<{src: Object, reason: string, candidates: string[]}>} items */
function appendDetailQueueRows_(items) {
  if (!items.length) return;
  var sh = ensureSheet_(SHEET.DETAIL_QUEUE, DETAILQ_HEADERS);
  var colSelect = DETAILQ_HEADERS.indexOf('선택상세분류') + 1; // 1-based

  var rows = items.map(function (it) {
    return srcToRowPrefix_(it.src).concat([
      it.reason || '',
      (it.candidates || []).join(', '),
      '',
      STATUS.PENDING,
      ''
    ]);
  });
  var startRow = sh.getLastRow() + 1;
  sh.getRange(startRow, 1, rows.length, DETAILQ_HEADERS.length).setValues(rows);

  items.forEach(function (it, i) {
    var candidates = it.candidates || [];
    if (candidates.length) {
      var rule = SpreadsheetApp.newDataValidation()
        .requireValueInList(candidates, true)
        .setAllowInvalid(false)
        .build();
      sh.getRange(startRow + i, colSelect).setDataValidation(rule);
    }
  });
}

/** 메뉴: 선택상세분류가 입력된 대기 건을 원장에 반영 */
function applyDetailSelections() {
  var ui = SpreadsheetApp.getUi();
  var sh = ensureSheet_(SHEET.DETAIL_QUEUE, DETAILQ_HEADERS);
  var last = sh.getLastRow();
  if (last < 2) { ui.alert('상세분류 확인 대기열에 처리할 건이 없습니다.'); return; }

  var width = DETAILQ_HEADERS.length;
  var colReason = DETAILQ_HEADERS.indexOf('판정근거');
  var colSelect = DETAILQ_HEADERS.indexOf('선택상세분류');
  var colStatus = DETAILQ_HEADERS.indexOf('상태');
  var colTime = DETAILQ_HEADERS.indexOf('처리일시');

  var data = sh.getRange(2, 1, last - 1, width).getValues();
  var now = nowStr_();

  var ledgerItems = [];
  var applied = 0;

  data.forEach(function (row) {
    if (normStr_(row[colStatus]) !== STATUS.PENDING) return;
    var detail = normStr_(row[colSelect]);
    if (!detail) return;

    var src = rowPrefixToSrc_(row);
    var reason = normStr_(row[colReason]);
    ledgerItems.push({
      src: src,
      detail: detail,
      reason: (reason ? reason + ' / ' : '') + '상세분류 수동선택',
      path: PATH.DETAIL
    });
    row[colStatus] = STATUS.DONE;
    row[colTime] = now;
    applied++;
  });

  if (!applied) {
    ui.alert('선택상세분류가 입력된 대기 건이 없습니다.\n선택상세분류 컬럼에서 후보를 고른 뒤 다시 실행해 주세요.');
    return;
  }

  sh.getRange(2, 1, data.length, width).setValues(data);
  upsertLedgerRows_(ledgerItems);
  rebuildMonthlySummary();
  ui.alert('상세분류 반영 완료: ' + applied + '건이 원장에 반영되었습니다.');
}
