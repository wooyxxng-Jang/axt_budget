/**
 * 상세분류(세부 라인) 자동 라우팅 (스펙 4장)
 * - 자금+약정항목 그룹의 상세분류가 0~1개면 즉시 확정 (26/32 그룹 → 100% 자동)
 * - 2개 이상이면 설정 시트의 라우팅 규칙(키워드/prefix/기본값) 적용
 * - 규칙 매칭 실패 시 '상세분류 확인 대기열'로 보냄
 */

/**
 * @param {Object} src 소스 행 객체
 * @return {{resolved: boolean, detail: string, candidates: string[], routeReason: string}}
 */
function routeDetail_(src) {
  var fund = src['자금'], itemCode = src['약정항목'];
  var candidates = getDetailCandidates_(fund, itemCode);

  if (candidates.length === 0) {
    // 마스터 미연결 또는 상세분류 없는 그룹 → 비목 레벨로 확정
    return { resolved: true, detail: '', candidates: [], routeReason: '' };
  }
  if (candidates.length === 1) {
    return { resolved: true, detail: candidates[0], candidates: candidates, routeReason: '단일 상세분류' };
  }

  var text = normStr_(src['텍스트']);
  var textUpper = text.toUpperCase();
  var rules = getRoutingRules_().filter(function (r) {
    return groupKey_(r.fund, r.itemCode) === groupKey_(fund, itemCode);
  });

  for (var i = 0; i < rules.length; i++) {
    var rule = rules[i];
    var matched = false;
    if (rule.mode === '기본값') {
      matched = true;
    } else if (rule.mode === 'prefix') {
      matched = textUpper.indexOf(rule.keyword.toUpperCase()) === 0;
    } else { // '포함'
      matched = textUpper.indexOf(rule.keyword.toUpperCase()) >= 0;
    }
    if (!matched) continue;

    // 규칙의 상세분류가 마스터 후보에 실제 존재할 때만 확정 (오타/플레이스홀더 방지)
    if (candidates.indexOf(rule.detail) >= 0) {
      var why = rule.mode === '기본값' ? '기본값 규칙' : '규칙 매칭(' + rule.mode + ': ' + rule.keyword + ')';
      return { resolved: true, detail: rule.detail, candidates: candidates, routeReason: why };
    }
  }

  return { resolved: false, detail: '', candidates: candidates, routeReason: '자동 라우팅 실패' };
}
