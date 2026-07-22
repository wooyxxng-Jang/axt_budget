/**
 * 상세분류(세부 라인) 자동 라우팅 (스펙 4장)
 * - 자금+약정항목 그룹의 상세분류가 1개면 즉시 확정 (대부분 그룹 → 100% 자동)
 * - 2개 이상이면 설정 시트의 라우팅 규칙(키워드/prefix/기본값) 적용
 * - 그래도 못 정하면, 후보 라벨 앞부분의 영문 접두어(TFI/AWI/ATC 등)가 텍스트에 그대로
 *   있는지로 자동 매칭 (설정 불필요 — 마스터에서 라벨을 이 규칙대로 짓기만 하면 바로 동작)
 * - 약정항목 코드가 마스터에 없으면(세인트 코드≠마스터 편성 코드) 자금 단위 후보로 대기열 이동
 * - 그래도 매칭 실패 시 '상세분류 확인 대기열'로 보내 사람이 직접 선택
 */

/**
 * @param {Object} src 소스 행 객체
 * @return {{resolved: boolean, detail: string, candidates: string[], routeReason: string}}
 */
function routeDetail_(src) {
  var fund = src['자금'], itemCode = src['약정항목'];
  var groupCand = getDetailCandidates_(fund, itemCode);

  // 그룹(자금+약정항목) 상세분류가 1개뿐 → 즉시 확정 (고신뢰 자동)
  if (groupCand.length === 1) {
    return { resolved: true, detail: groupCand[0], candidates: groupCand, routeReason: '단일 상세분류' };
  }

  // 후보 결정: 그룹 후보가 2개 이상이면 그룹 후보, 그룹에 없으면 자금 단위 후보로 확장
  var fundLevel = groupCand.length === 0;
  var candidates = fundLevel ? getFundDetailCandidates_(fund) : groupCand;

  if (candidates.length === 0) {
    // 마스터 미연결이거나 해당 자금 자체가 마스터에 없음 → 비목 레벨로 확정(미편성 집행)
    return { resolved: true, detail: '', candidates: [], routeReason: '' };
  }

  // 라우팅 규칙 적용 (후보에 실제 존재하는 상세분류로만 확정)
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
    if (candidates.indexOf(rule.detail) >= 0) {
      var why = rule.mode === '기본값' ? '기본값 규칙' : '규칙 매칭(' + rule.mode + ': ' + rule.keyword + ')';
      return { resolved: true, detail: rule.detail, candidates: candidates, routeReason: why };
    }
  }

  // 라벨 접두어 키워드 자동 매칭 (설정 불필요) — 후보 라벨 앞부분의 영문 접두어
  // (TFI/AWI/ATC 등)가 텍스트에 그대로 들어있으면 그 라벨로 확정.
  // 새 접두어 프로그램이 생겨도 마스터 라벨만 이 규칙대로 지으면 설정 추가 없이 바로 동작한다.
  var prefixMatches = candidates.filter(function (c) {
    var prefix = detailLabelPrefix_(c);
    return prefix && textUpper.indexOf(prefix) >= 0;
  });
  if (prefixMatches.length === 1) {
    return {
      resolved: true, detail: prefixMatches[0], candidates: candidates,
      routeReason: '라벨 접두어 매칭(' + detailLabelPrefix_(prefixMatches[0]) + ')'
    };
  }
  if (prefixMatches.length > 1) {
    candidates = prefixMatches; // 접두어로 좁혀진 후보만 대기열에 제시
  }

  // 자동 라우팅 실패 → 상세분류 확인 대기열 (사람이 후보 중 직접 선택)
  var reason = fundLevel
    ? '약정항목 코드(' + itemCode + ')가 마스터에 없어 자금 단위 상세분류 선택 필요'
    : '자동 라우팅 실패';
  return { resolved: false, detail: '', candidates: candidates, routeReason: reason };
}

/** 상세분류 라벨 앞부분의 영문 접두어(2자 이상, 예: TFI/AWI/ATC) 추출. 없으면 ''. */
function detailLabelPrefix_(label) {
  var m = normStr_(label).match(/^[A-Za-z]{2,}/);
  return m ? m[0].toUpperCase() : '';
}
