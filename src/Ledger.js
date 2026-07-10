/**
 * 원장(Ledger) — 아텍 귀속 확정 전건 누적, upsert 키 = 참조전표+개별항목 (스펙 6장)
 *
 * upsert 규칙:
 *  - 신규 키 → 행 추가 (상세분류/판정근거/반영경로/반영일시 포함)
 *  - 기존 키 → 소스 컬럼(전표상태/사용금액 등)만 갱신.
 *    상세분류·판정근거·반영경로는 유지 — 사람이 확정한 결과가 재업로드로 깨지지 않게.
 *    단, 수동 확정 경로(검토확정/장학금태깅/상세분류확정)로 다시 반영하는 경우는 덮어씀.
 */

/**
 * @param {Array<{src: Object, detail: string, reason: string, path: string, updateOnly: boolean}>} items
 * @return {{inserted: number, updated: number}}
 */
function upsertLedgerRows_(items) {
  if (!items.length) return { inserted: 0, updated: 0 };

  var sh = ensureSheet_(SHEET.LEDGER, LEDGER_HEADERS);
  var width = LEDGER_HEADERS.length;
  var colDetail = LEDGER_HEADERS.indexOf('상세분류');
  var colReason = LEDGER_HEADERS.indexOf('판정근거');
  var colPath = LEDGER_HEADERS.indexOf('반영경로');
  var colTime = LEDGER_HEADERS.indexOf('반영일시');

  var last = sh.getLastRow();
  var data = last > 1 ? sh.getRange(2, 1, last - 1, width).getValues() : [];
  var idx = {};
  for (var r = 0; r < data.length; r++) {
    var k = normStr_(data[r][0]);
    if (k) idx[k] = r;
  }

  var now = nowStr_();
  var appends = [];
  var dirty = false;
  var inserted = 0, updated = 0;

  items.forEach(function (it) {
    var key = makeKey_(it.src['참조전표'], it.src['개별항목']);
    var prefix = srcToRowPrefix_(it.src); // [업서트키, ...SRC_COLS]

    if (idx[key] !== undefined) {
      var row = data[idx[key]];
      for (var c = 0; c < prefix.length; c++) row[c] = prefix[c];
      var manualReapply = it.path && it.path !== PATH.AUTO;
      if (manualReapply) {
        row[colDetail] = it.detail || '';
        row[colReason] = it.reason || '';
        row[colPath] = it.path;
      }
      row[colTime] = now;
      dirty = true;
      updated++;
    } else if (!it.updateOnly) {
      var newRow = prefix.concat([it.detail || '', it.reason || '', it.path || PATH.AUTO, now]);
      appends.push(newRow);
      inserted++;
    } else {
      // updateOnly 인데 원장에 없음(비정상) — 안전하게 추가
      appends.push(prefix.concat(['', '재업로드 복원', PATH.AUTO, now]));
      inserted++;
    }
  });

  if (dirty && data.length) {
    sh.getRange(2, 1, data.length, width).setValues(data);
  }
  if (appends.length) {
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, width).setValues(appends);
  }
  return { inserted: inserted, updated: updated };
}

/** 원장 전체 데이터(2행부터) — 월별집계/드릴다운용 */
function readLedgerData_() {
  var sh = ensureSheet_(SHEET.LEDGER, LEDGER_HEADERS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, LEDGER_HEADERS.length).getValues();
}
