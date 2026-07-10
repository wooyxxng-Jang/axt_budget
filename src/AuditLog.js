/**
 * 감사로그 — 자동/수동 제외된 건 전체 기록, 스팟체크용 (스펙 6장)
 * 키가 이미 기록돼 있으면 다시 추가하지 않는다 (재업로드 중복 방지).
 */

/**
 * @param {Array<{src: Object, reason: string}>} items
 * @param {Object} existingKeys buildKeySet_ 결과 (없으면 내부에서 생성)
 * @return {number} 추가된 건수
 */
function appendAuditRows_(items, existingKeys) {
  if (!items.length) return 0;
  var sh = ensureSheet_(SHEET.AUDIT, AUDIT_HEADERS);
  var keys = existingKeys || buildKeySet_(sh);
  var now = nowStr_();

  var rows = [];
  items.forEach(function (it) {
    var key = makeKey_(it.src['참조전표'], it.src['개별항목']);
    if (keys[key]) return;
    keys[key] = true;
    rows.push(srcToRowPrefix_(it.src).concat([it.reason || '', now]));
  });

  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, AUDIT_HEADERS.length).setValues(rows);
  }
  return rows.length;
}
