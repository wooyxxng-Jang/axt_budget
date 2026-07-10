/**
 * 업로드 파이프라인 (스펙 7장 처리 흐름)
 *
 * 1. 세인트 .xls(HTML) 파싱
 * 2. 이미 원장에 있는 키 → 소스 데이터만 갱신 (수동 확정 결과 보존)
 * 3. 전표유형=등록금 → 장학금 분류 대기함 (신규만)
 * 4. 나머지 → 키워드 우선순위 판정
 *    - 포함 → 상세분류 라우팅 → 원장 or 상세분류 확인 대기열
 *    - 제외 → 감사로그
 *    - 판정불가 → 검토대기함
 * 5. 월별집계 재계산
 */

function showUploadDialog() {
  var html = HtmlService.createHtmlOutputFromFile('UploadDialog')
    .setWidth(480)
    .setHeight(420);
  SpreadsheetApp.getUi().showModalDialog(html, '세인트 파일 업로드');
}

/**
 * 업로드 다이얼로그에서 호출. HTML 원문을 받아 전체 파이프라인 실행.
 * @param {string} content 파일 텍스트(UTF-8)
 * @return {Object} 처리 결과 통계
 */
function processUpload(content) {
  var rows = parseSaintHtml_(content);
  if (!rows.length) throw new Error('파싱된 데이터 행이 없습니다.');

  var kw = getKeywordLists_();
  if (!kw.strong.length && !kw.professor.length && !kw.otherDept.length) {
    throw new Error('설정 시트에 키워드가 없습니다. 메뉴 [AXT 예산 > 시트 초기화/점검]을 먼저 실행해 주세요.');
  }
  getMasterIndex_(); // 마스터 설정 오류 조기 감지 + 캐시 준비

  var ledgerKeys = buildKeySet_(ensureSheet_(SHEET.LEDGER, LEDGER_HEADERS));
  var auditKeys = buildKeySet_(ensureSheet_(SHEET.AUDIT, AUDIT_HEADERS));
  var reviewKeys = buildKeySet_(ensureSheet_(SHEET.REVIEW, REVIEW_HEADERS));
  var scholKeys = buildKeySet_(ensureSheet_(SHEET.SCHOLARSHIP, SCHOLARSHIP_HEADERS));
  var detailKeys = buildKeySet_(ensureSheet_(SHEET.DETAIL_QUEUE, DETAILQ_HEADERS));

  var stats = {
    parsed: rows.length,
    dupInFile: 0,
    ledgerUpdated: 0,
    autoIncluded: 0,
    autoExcluded: 0,
    reviewQueued: 0,
    detailQueued: 0,
    scholarshipNew: 0,
    alreadyKnown: 0
  };

  var ledgerItems = [], auditItems = [], reviewItems = [], scholItems = [], detailItems = [];
  var seenInFile = {};

  rows.forEach(function (src) {
    var key = makeKey_(src['참조전표'], src['개별항목']);
    if (seenInFile[key]) { stats.dupInFile++; return; }
    seenInFile[key] = true;

    // 이미 원장에 확정된 건 → 소스 데이터만 갱신 (판정/상세분류 유지)
    if (ledgerKeys[key]) {
      ledgerItems.push({ src: src, updateOnly: true });
      stats.ledgerUpdated++;
      return;
    }

    // 장학금(등록금 유형) — 자동분류 대상 아님, 신규만 대기함으로 (스펙 5장)
    if (src['전표유형명'].indexOf(SCHOLARSHIP_DOC_TYPE) >= 0) {
      if (scholKeys[key] || auditKeys[key]) { stats.alreadyKnown++; return; }
      scholItems.push({ src: src });
      scholKeys[key] = true;
      stats.scholarshipNew++;
      return;
    }

    var cls = classifyText_(src['텍스트'], kw);

    if (cls.verdict === CLS.INCLUDE) {
      var route = routeDetail_(src);
      if (route.resolved) {
        var reason = cls.reason + (route.routeReason ? ' / ' + route.routeReason : '');
        ledgerItems.push({ src: src, detail: route.detail, reason: reason, path: PATH.AUTO });
        stats.autoIncluded++;
      } else {
        if (detailKeys[key]) { stats.alreadyKnown++; return; }
        detailItems.push({ src: src, reason: cls.reason, candidates: route.candidates });
        detailKeys[key] = true;
        stats.detailQueued++;
      }
    } else if (cls.verdict === CLS.EXCLUDE) {
      if (auditKeys[key]) { stats.alreadyKnown++; return; }
      auditItems.push({ src: src, reason: cls.reason });
      stats.autoExcluded++;
    } else { // REVIEW
      if (reviewKeys[key] || auditKeys[key]) { stats.alreadyKnown++; return; }
      reviewItems.push({ src: src, reason: cls.reason });
      reviewKeys[key] = true;
      stats.reviewQueued++;
    }
  });

  var up = upsertLedgerRows_(ledgerItems);
  stats.ledgerInserted = up.inserted;
  appendAuditRows_(auditItems, auditKeys);
  appendReviewRows_(reviewItems);
  appendScholarshipRows_(scholItems);
  appendDetailQueueRows_(detailItems);

  rebuildMonthlySummary();
  return stats;
}
