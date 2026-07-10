/**
 * 월별집계 (스펙 6장)
 * - 배정예산 마스터와 동일한 레이아웃(자금/비목/상세분류) × 월별 컬럼
 * - 원장 데이터를 전표일자 기준으로 자동 합산 + 배정액 대비 잔액 계산
 * - 매번 원장 전체를 기준으로 재계산 (재업로드에도 안전)
 */

var SUMMARY_FIXED_HEADERS = ['자금', '자금명', '약정항목', '약정항목 명', '상세분류', '배정액'];

function rebuildMonthlySummary() {
  var master = getMasterIndex_();
  var ledger = readLedgerData_();

  var cFund = LEDGER_HEADERS.indexOf('자금');
  var cFundName = LEDGER_HEADERS.indexOf('자금명');
  var cItem = LEDGER_HEADERS.indexOf('약정항목');
  var cItemName = LEDGER_HEADERS.indexOf('약정항목 명');
  var cDate = LEDGER_HEADERS.indexOf('전표일자');
  var cAmt = LEDGER_HEADERS.indexOf('사용금액');
  var cDetail = LEDGER_HEADERS.indexOf('상세분류');

  // 원장 → 라인(자금|약정항목|상세분류)별 × 월별 합산
  var usage = {};        // lineKey → { monthKey → amount }
  var monthsSet = {};
  var ledgerLines = {};  // lineKey → 라인 메타 (마스터에 없는 라인 대비)

  ledger.forEach(function (row) {
    var fund = normStr_(row[cFund]);
    var itemCode = normStr_(row[cItem]);
    var detail = effectiveDetail_(row[cDetail], fund, itemCode);
    var lk = fund + '|' + itemCode + '|' + detail;
    var mk = monthKey_(row[cDate]) || '(일자없음)';
    monthsSet[mk] = true;

    if (!usage[lk]) usage[lk] = {};
    usage[lk][mk] = (usage[lk][mk] || 0) + (Number(row[cAmt]) || 0);

    if (!ledgerLines[lk]) {
      ledgerLines[lk] = {
        fund: fund,
        fundName: normStr_(row[cFundName]),
        itemCode: itemCode,
        itemName: normStr_(row[cItemName]),
        detail: detail,
        budget: ''
      };
    }
  });

  var months = Object.keys(monthsSet).sort();

  // 라인 목록: 마스터 라인 전체(실적 0이어도 표시) + 마스터에 없는 원장 라인
  var lines = [];
  var lineKeySeen = {};
  master.lines.forEach(function (l) {
    var lk = l.fund + '|' + l.itemCode + '|' + l.detail;
    if (lineKeySeen[lk]) return;
    lineKeySeen[lk] = true;
    lines.push(l);
  });
  Object.keys(ledgerLines).forEach(function (lk) {
    if (!lineKeySeen[lk]) {
      lineKeySeen[lk] = true;
      lines.push(ledgerLines[lk]);
    }
  });
  lines.sort(function (a, b) {
    if (a.fund !== b.fund) return a.fund < b.fund ? -1 : 1;
    if (a.itemCode !== b.itemCode) return a.itemCode < b.itemCode ? -1 : 1;
    return a.detail < b.detail ? -1 : a.detail > b.detail ? 1 : 0;
  });

  // 출력 테이블 구성
  var headers = SUMMARY_FIXED_HEADERS.concat(months, ['사용합계', '잔액']);
  var out = lines.map(function (l) {
    var lk = l.fund + '|' + l.itemCode + '|' + l.detail;
    var byMonth = usage[lk] || {};
    var total = 0;
    var monthCells = months.map(function (mk) {
      var v = byMonth[mk] || 0;
      total += v;
      return v === 0 ? '' : v;
    });
    var budget = (l.budget === '' || l.budget === null || l.budget === undefined) ? '' : Number(l.budget);
    var remain = budget === '' ? '' : budget - total;
    return [l.fund, l.fundName, l.itemCode, l.itemName, l.detail, budget]
      .concat(monthCells, [total, remain]);
  });

  // 시트 기록
  var sh = ensureSheet_(SHEET.SUMMARY, null);
  sh.clearContents();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  if (out.length) {
    sh.getRange(2, 1, out.length, headers.length).setValues(out);
    // 금액 컬럼 포맷 (배정액부터 끝까지)
    sh.getRange(2, SUMMARY_FIXED_HEADERS.length, out.length, headers.length - SUMMARY_FIXED_HEADERS.length + 1)
      .setNumberFormat('#,##0');
  }
  sh.setFrozenRows(1);
  sh.setFrozenColumns(5);
  return { lines: out.length, months: months.length };
}
