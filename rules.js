"use strict";
/* 규칙 기반 예상 질의 자동 생성 (AI 없이 동작)
   발표자료의 슬라이드 제목·본문·수치·RFP 번호·주장 표현을 읽어 평가위원 관점의 질문을 만든다.
   답변은 만들지 않는다. 대신 화면에서 해당 슬라이드 내용을 답변 준비 근거로 보여 준다. */
(function(root){
  const RFP_RE = /\b([A-Z]{2,5}-\d{2,4}(?:-\d{1,3})?)\b/g;
  const NUM_RE = /\d[\d,.]*\s*(%|퍼센트|배|건|시간|분|초|억|만|천|TB|GB|ms|회)(?![가-힣A-Za-z])/;
  const YEAR_ONLY = /^\s*(19|20)\d{2}\s*년?\s*\.?\s*(\d{1,2}\s*월?)?\s*$/;
  const SKIP_TITLE = /^(목차|순서|contents?|agenda|index|감사합니다|thank|q\s*&\s*a|질의\s*응답|end|끝)(?![가-힣A-Za-z])/i;
  const VAGUE = ['최초','최고','유일','독보적','혁신적','획기적','완벽','최적','무중단','100%','완전 자동','원스톱','세계적','선도'];
  const TECH_WORDS = ['온톨로지','에이전트','생성형','머신러닝','딥러닝','클라우드','데이터 레이크','데이터레이크','파이프라인','디지털 트윈','디지털트윈','지식그래프','지식 그래프','벡터','임베딩','파인튜닝','추론','멀티모달','마이크로서비스','컨테이너','쿠버네티스','오픈소스'];
  const ACR_STOP = new Set(['PM','PL','PMO','AS','IS','TO','BE','QA','HW','SW','OK','NO','VS','KR','EN','CEO','CTO','PPT','PDF','URL','FAQ','ETC','TBD','DB']);
  const KW = {
    security: /보안|개인정보|망\s*분리|권한|암호화|접근\s*통제|취약점|인증/,
    integrate: /연계|연동|기존\s*시스템|현행|이관|마이그레이션|전환|레거시|인터페이스/,
    schedule: /일정|착수|개월|단계별|마일스톤|WBS|투입\s*인력|인력\s*구성|M\/M|조직\s*구성|수행\s*조직/,
    diff: /차별|경쟁|강점|독자|특허|자체\s*개발|원천\s*기술|레퍼런스|구축\s*사례|수행\s*실적/,
    effect: /기대\s*효과|효과|절감|향상|개선|단축|증대|ROI|생산성/,
    maint: /유지\s*보수|운영\s*지원|기술\s*이전|교육|하자\s*보수|안정화|운영\s*이관/,
    data: /데이터|품질|정합성|표준화|메타데이터|수집|적재/
  };
  const CAT = {rfp:'수행 방안', num:'기대효과·정량 목표', vague:'경쟁사 대비 차별화', tech:'기술·아키텍처', security:'데이터·보안', integrate:'기술·아키텍처', schedule:'일정·인력', diff:'경쟁사 대비 차별화', effect:'기대효과·정량 목표', maint:'수행 방안', data:'데이터·보안', title:'사업 이해', global:'수행 방안'};

  const clip = (s, n=38) => { s = String(s||'').replace(/\s+/g,' ').trim(); return s.length > n ? s.slice(0, n - 1).trim() + '…' : s; };
  const norm = s => String(s||'').replace(/[\s\p{P}\p{S}]/gu,'').toLowerCase();
  const pick = (arr, seed) => arr[Math.abs(seed) % arr.length];

  function slideLines(s){ return String((s.title||'') + '\n' + (s.body||'')).split('\n').map(x => x.trim()).filter(Boolean); }
  function techTerm(text){
    text = text.split('\n').filter(l => !/현행|기존|As-?Is/i.test(l)).join('\n');
    const acr = (text.match(/\b[A-Z][A-Za-z0-9]*[A-Z0-9][A-Za-z0-9]*(?:[-/][A-Za-z0-9]+)?\b/g) || []).filter(w => w.length >= 2 && !ACR_STOP.has(w.toUpperCase()) && !/^[A-Z]{2,5}-\d/.test(w));
    const ko = TECH_WORDS.filter(w => text.includes(w));
    const all = [...acr.filter(w => w.length >= 3), ...ko];
    return all[0] || '';
  }
  function titleNoun(t){ return clip(String(t||'').replace(/^\s*[\dⅠ-Ⅻ.\-\)]+\s*/, '').replace(/\(.*?\)/g,''), 24); }

  function forSlide(s, total, self){
    const out = [];
    const lines = slideLines(s); const text = lines.join('\n');
    const no = s.no; const t = titleNoun(s.title);
    const add = (prio, kind, q, extra={}) => out.push({prio, kind, cat: CAT[kind], q, page: String(no), rfp: extra.rfp || '', kw: extra.kw || ''});
    // 1. RFP 요구사항 번호
    const ids = [...new Set((text.match(RFP_RE) || []))].slice(0, 2);
    ids.forEach((id, i) => add(100 - i, 'rfp', pick([
      `${id} 요구사항은 구체적으로 어떤 방식으로 충족하며, 검수 단계에서 무엇으로 확인할 수 있습니까?`,
      `${id} 요구사항 중 이행이 가장 까다로운 부분은 무엇이고, ${self}는 어떻게 해결할 계획입니까?`
    ], no + i), {rfp: id, kw: [id, t].filter(Boolean).join(', ')}));
    // 2. 정량 수치
    const numLine = KW.schedule.test(text) ? '' : lines.slice(1).find(l => NUM_RE.test(l) && !YEAR_ONLY.test(l) && /[가-힣]/.test(l)) || (NUM_RE.test(lines[0]||'') && /[가-힣]/.test(lines[0]) && !YEAR_ONLY.test(lines[0]) ? lines[0] : '');
    if(numLine){
      const isEffect = KW.effect.test(numLine) || KW.effect.test(s.title||'');
      const early0 = no <= Math.max(2, Math.round(total * 0.2));
      add(90, isEffect ? 'effect' : 'num', pick(isEffect ? [
        `'${clip(numLine)}'라고 제시하셨는데, 사업 종료 시점에 어떤 지표와 방법으로 이 효과를 확인할 수 있습니까?`,
        `'${clip(numLine)}' 목표를 달성하지 못하면 어떻게 대응합니까? 측정 기준은 무엇입니까?`
      ] : [
        `'${clip(numLine)}'라는 수치는 어떤 기준과 데이터로 산정한 것입니까?`,
        `'${clip(numLine)}'라고 하셨는데, 이 수치를 검증할 수 있는 근거나 실적이 있습니까?`
      ], no), {kw: t});
      if(!isEffect) out[out.length - 1].cat = early0 ? '사업 이해' : '수행 방안';
    }
    // 3. 근거가 필요한 주장 표현
    const vw = VAGUE.find(w => text.includes(w));
    if(vw){
      const vl = lines.find(l => l.includes(vw)) || vw;
      add(80, 'vague', pick([
        `'${clip(vl)}'라고 표현하셨는데, 이를 뒷받침하는 실적이나 근거를 구체적으로 말씀해 주십시오.`,
        `'${clip(vl)}'라는 부분이 다른 제안사의 방식과 실제로 무엇이 다른지 설명해 주시겠습니까?`
      ], no), {kw: [vw, t].filter(Boolean).join(', ')});
    }
    // 4. 주제별 질문
    const themed = [
      ['security', 70, [`${no}쪽 보안 관련 내용과 관련해, 내부 데이터와 개인정보가 외부로 나가지 않는다는 것을 어떻게 보장합니까?`, `${no}쪽에서 말씀하신 보안 대책은 발주기관 보안 지침과 어떻게 맞춰 적용합니까?`]],
      ['integrate', 68, [`기존 시스템과 연계·이관할 때 데이터 정합성은 어떻게 확인하고, 서비스 중단 없이 전환하는 방안은 무엇입니까?`, `'${t}'에서 현행 시스템과 연계되는 범위는 어디까지이며, 연계 대상 기관이나 시스템의 협조가 필요한 부분은 무엇입니까?`]],
      ['schedule', 66, [`제시하신 일정과 투입 인력으로 범위를 모두 소화할 수 있습니까? 지연되면 어떤 단계부터 조정합니까?`, `핵심 투입 인력이 중간에 바뀌면 어떻게 대응합니까? 핵심 인력의 유사 사업 수행 경험도 말씀해 주십시오.`]],
      ['diff', 64, [`'${t}' 부분에서 다른 제안사와 비교한 ${self}만의 차별점은 무엇이며, 확인할 수 있는 근거는 무엇입니까?`, `말씀하신 수행 실적은 이번 사업과 규모나 범위가 어느 정도 비슷합니까?`]],
      ['maint', 60, [`사업 종료 후 운영·유지보수와 기술 이전은 어떤 방식과 기간으로 진행합니까?`, `발주기관 담당자가 직접 운영할 수 있도록 교육과 기술 이전을 어떻게 계획하고 있습니까?`]],
      ['data', 58, [`'${t}'에서 데이터 품질은 어떤 기준으로 관리하고, 품질 문제가 생기면 어떻게 조치합니까?`, `'${t}'에 필요한 데이터를 확보하지 못하거나 품질이 낮으면 어떻게 대응합니까?`]]
    ];
    for(const [kind, prio, tpls] of themed){
      if(KW[kind].test(text)){ add(prio, kind, pick(tpls, no), {kw: t}); break; }
    }
    // 5. 기술 용어
    const tech = techTerm(text);
    if(tech) add(62, 'tech', pick([
      `${tech} 적용과 관련해, 기술적으로 가장 위험한 지점은 무엇이고 어떻게 검증하실 계획입니까?`,
      `${tech} 관련 유사 사업 적용 사례가 있습니까? 있다면 어떤 성과가 있었습니까?`,
      `${tech} 구성에서 오류나 장애가 생기면 어떻게 감지하고 복구합니까?`
    ], no), {kw: [tech, t].filter(Boolean).join(', ')});
    // 6. 슬라이드 핵심 질문 (항상 하나)
    const early = no <= Math.max(2, Math.round(total * 0.2));
    if(t && (out.length === 0 || (early && out.length < 3))){
      add(early ? 55 : 40, 'title', pick([
        `'${t}'에서 가장 강조하신 내용을 한 문장으로 정리하면 무엇이고, 발주기관이 얻는 효과는 무엇입니까?`,
        `'${t}' 부분은 현행 방식과 비교해 무엇이 달라지는지 구체적으로 설명해 주시겠습니까?`,
        `'${t}' 내용을 실제로 구현할 때 가장 어려운 부분은 무엇이고, 어떻게 해결할 계획입니까?`
      ], no), {kw: t});
      if(!early) out[out.length - 1].cat = KW.data.test(text) ? '데이터·보안' : '수행 방안';
    }
    return out.sort((a,b) => b.prio - a.prio);
  }

  function proposalRfp(text, deckIds){
    const out = []; const seen = new Set(deckIds);
    const parts = String(text||'').split(/(?=\[제안서 p\.\d+\])/);
    for(const part of parts){
      const m = part.match(/^\[제안서 p\.(\d+)\]/); const page = m ? '제안서 p.' + m[1] : '제안서';
      for(const id of (part.match(RFP_RE) || [])){
        if(seen.has(id)) continue; seen.add(id);
        out.push({prio: 50, kind: 'rfp', cat: CAT.rfp, q: `${id} 요구사항은 발표에서 따로 다루지 않으셨는데, 어떻게 충족할 계획인지 간단히 설명해 주시겠습니까?`, page, rfp: id, kw: id});
      }
    }
    return out;
  }

  function generate({slides, proposalText, existing, self, max, perSlide, includeProposal}){
    self = self || '제안사'; max = max || 40; perSlide = perSlide || 3;
    const usable = (slides || []).filter(s => {
      const lines = slideLines(s);
      if(!lines.length) return false;
      if(SKIP_TITLE.test(String(s.title||'').trim())) return false;
      if(s.no === 1 && lines.length <= 3) return false; // 표지
      return true;
    });
    const per = usable.map(s => forSlide(s, slides.length, self).slice(0, perSlide));
    const seen = new Set((existing || []).map(norm));
    const out = [];
    const push = c => { const k = norm(c.q); if(seen.has(k)) return; seen.add(k); out.push({cat: c.cat, q: c.q, a: '', page: c.page, rfp: c.rfp, kw: c.kw}); };
    // 슬라이드마다 우선순위 높은 것부터 돌아가며 뽑아 고르게 분포
    for(let round = 0; round < perSlide && out.length < max; round++){
      for(const list of per){ if(out.length >= max) break; if(list[round]) push(list[round]); }
    }
    if(usable.length){
      [[`이 사업에서 가장 큰 위험 요소 세 가지를 꼽는다면 무엇이고, 각각 어떻게 대응합니까?`, '수행 방안'],
       [`발표 내용 중 발주기관이 꼭 기억해야 할 ${self}의 강점 한 가지를 꼽는다면 무엇입니까?`, '경쟁사 대비 차별화']]
        .forEach(([q, cat]) => { if(out.length < max) push({q, cat, page: '', rfp: '', kw: '종합'}); });
    }
    if(includeProposal && out.length < max){
      const deckIds = usable.flatMap(s => slideLines(s).join(' ').match(RFP_RE) || []);
      for(const c of proposalRfp(proposalText, deckIds)){ if(out.length >= max) break; push(c); }
    }
    return out;
  }

  root.RuleQA = {generate, _forSlide: forSlide};
})(typeof window !== 'undefined' ? window : globalThis);
