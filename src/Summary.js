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

  var months = fiscalMonths_(monthsSet);

  // 라인 순서 = 배정예산 마스터 시트의 행 순서를 그대로 따른다.
  //  1) 마스터 라인 전체(실적 0이어도 표시)를 마스터 원본 순서대로
  //  2) 마스터에 없는 원장 라인(배정예산 미편성 집행)은 맨 아래에 별도로 모아 표시
  //     — 자금/약정항목 순으로 정렬해, 계획에 없던 집행을 한눈에 확인 가능
  var lines = [];
  var extraLines = [];
  var lineKeySeen = {};
  master.lines.forEach(function (l) {
    var lk = l.fund + '|' + l.itemCode + '|' + l.detail;
    if (lineKeySeen[lk]) return;
    lineKeySeen[lk] = true;
    lines.push(l); // 마스터 원본 순서 유지 (재정렬하지 않음)
  });
  Object.keys(ledgerLines).forEach(function (lk) {
    if (!lineKeySeen[lk]) {
      lineKeySeen[lk] = true;
      extraLines.push(ledgerLines[lk]);
    }
  });
  extraLines.sort(function (a, b) {
    if (a.fund !== b.fund) return a.fund < b.fund ? -1 : 1;
    if (a.itemCode !== b.itemCode) return a.itemCode < b.itemCode ? -1 : 1;
    return a.detail < b.detail ? -1 : a.detail > b.detail ? 1 : 0;
  });
  lines = lines.concat(extraLines);

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

/**
 * 월 컬럼 목록 — 배정예산 마스터와 동일하게 회계연도(3월~익년 2월) 12개월 고정.
 * 회계연도는 원장 데이터의 가장 이른 월로 판단하고, 범위 밖 월과 '(일자없음)'은 뒤에 덧붙인다.
 */
function fiscalMonths_(monthsSet) {
  var dataMonths = Object.keys(monthsSet).filter(function (m) { return /^\d{4}-\d{2}$/.test(m); }).sort();
  if (!dataMonths.length) {
    return monthsSet['(일자없음)'] ? ['(일자없음)'] : [];
  }

  var y = Number(dataMonths[0].substring(0, 4));
  var m = Number(dataMonths[0].substring(5, 7));
  var fyStartYear = m >= 3 ? y : y - 1;

  var months = [];
  for (var i = 0; i < 12; i++) {
    var mm = 3 + i;
    var yy = fyStartYear;
    if (mm > 12) { mm -= 12; yy += 1; }
    months.push(yy + '-' + ('0' + mm).slice(-2));
  }
  // 회계연도 범위 밖 데이터 월 보존
  dataMonths.forEach(function (dm) {
    if (months.indexOf(dm) < 0) months.push(dm);
  });
  if (monthsSet['(일자없음)']) months.push('(일자없음)');
  return months;
}
