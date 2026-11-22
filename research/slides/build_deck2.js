// 발표 자료 v2 (장수 제한 없음). 실행: node build_deck2.js
const pptxgen = require('pptxgenjs');
const SK = '/root/.claude/skills/synced/b09188c3-4889-4bd8-b551-d0a19e63b56f_facbcb1f-1d3a-4bc4-bdfb-98f8c08ef25b/pptx';
const { applyTheme } = require(SK + '/scripts/apply_theme.js');
const FIG = '/home/user/pango-gmp/research/figures/';
const THEME = { name: 'DC Study', headFontFace: 'Malgun Gothic', bodyFontFace: 'Malgun Gothic',
  colors: { dk1: '1B2733', lt1: 'FFFFFF', dk2: '0F766E', lt2: 'F1F5F9', accent1: '0F766E', accent2: 'F2A93B', accent3: 'B23A48', accent4: '64748B', accent5: '2B3B4C', accent6: 'CBD5E1', hlink: '0F766E', folHlink: '64748B' } };
const pres = new pptxgen(); pres.layout = 'LAYOUT_WIDE';
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = '건축 초기 계획설계의 지식 확장을 위한 Micro-App 분석'; pres.author = '임상철';
const C = pres.SchemeColor;
pres.defineSlideMaster({ title: 'DARK', background: { color: '1B2733' }, objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.7, y: 0.9, w: 11.5, h: 1.8, fontSize: 40, bold: true, color: C.background1, valign: 'top', align: 'left', margin: 0 }, text: '' } }] });
pres.defineSlideMaster({ title: 'CONTENT', background: { color: 'FFFFFF' }, objects: [
  { placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.35, w: 12.1, h: 0.85, fontSize: 30, bold: true, color: C.text1, valign: 'middle', align: 'left', margin: 0 }, text: '' } },
  { text: { text: '한양대학교 공학대학원 | 건축계획연구', options: { x: 0.6, y: 7.05, w: 6, h: 0.3, fontSize: 10, color: C.accent4, margin: 0 } } }],
  slideNumber: { x: 12.2, y: 7.05, w: 0.5, h: 0.3, fontSize: 10, color: C.accent4, align: 'right' } });
const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0 }, o));
const content = (title, notes) => { const s = pres.addSlide({ masterName: 'CONTENT' }); s.addText(title, { placeholder: 'title' }); if (notes) s.addNotes(notes); return s; };
const card = (s, x, y, w, h, head, body, o = {}) => {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: o.fill || C.background2 }, rectRadius: 0.1 });
  T(s, head, { x: x + 0.2, y: y + 0.12, w: w - 0.4, h: o.hh || 0.5, fontSize: o.hs || 18, bold: true, color: o.hc || C.text1, valign: 'middle' });
  if (body) T(s, body, { x: x + 0.2, y: y + (o.hh || 0.5) + 0.15, w: w - 0.4, h: h - (o.hh || 0.5) - 0.25, fontSize: o.bs || 15, color: o.bc || C.text1, valign: 'top' });
};
const img = (s, f, x, y, w, h) => s.addImage({ path: FIG + f, x, y, w, h, sizing: { type: 'contain', w, h } });
const src = (s, t, y = 6.7) => T(s, t, { x: 0.6, y, w: 12.1, h: 0.28, fontSize: 11, color: C.accent4 });
const bullets = (s, items, x, y, w, h, fs = 16) => s.addText(items.map((t, i) => ({ text: t, options: { bullet: { indent: 18 }, breakLine: i < items.length - 1, paraSpaceAfter: 8 } })), { x, y, w, h, fontSize: fs, color: C.text1, valign: 'top', margin: 0 });

// 1 표지
{ const s = pres.addSlide({ masterName: 'DARK' });
  s.addText('건축 초기 계획설계의 지식 확장을 위한\nMicro-App 분석', { placeholder: 'title' });
  T(s, '바이브 코딩으로 개발한 데이터센터 검토 앱 3개 사례를 중심으로', { x: 0.7, y: 3.2, w: 11.5, h: 0.5, fontSize: 20, color: C.accent6 });
  T(s, '공학대학원 건축학전공  임상철 (2026160492)', { x: 0.7, y: 5.3, w: 11, h: 0.4, fontSize: 18, color: C.background1 });
  T(s, '담당교수 김영주  |  2026년 11월 23일', { x: 0.7, y: 5.75, w: 11, h: 0.4, fontSize: 16, color: C.accent6 });
  s.addNotes('안녕하십니까. 건축 초기 계획설계의 지식 확장을 위한 Micro-App 분석을 발표하겠습니다. 바이브 코딩으로 개발한 데이터센터 검토 앱 3개를 사례로 분석했습니다.'); }

// 2 문제
{ const s = content('초기 계획설계: 건축가가 혼자 판단해야 한다', '초기 계획설계에서는 구조, 기계, 전기 분야의 협업과 지원을 받기 어려워 건축가가 경험과 사례로 기술 조건을 판단합니다. 이 판단이 나중의 전문 검토와 크게 다르면 건축 규모와 계획이 바뀌는 설계 변경 리스크가 생깁니다. 그래서 설계자의 지식 확장이 필요합니다.');
  const st = [['건축 외 분야의\n협업·지원이 어려움', C.accent5], ['건축가가 경험과 사례로\n기술 조건을 판단', C.accent5], ['이후 전문 검토와\n크게 다름', C.accent4], ['건축 규모·계획 변경\n= 설계 변경 리스크', C.accent3]];
  st.forEach((b, i) => { const x = 0.6 + i * 3.1;
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.6, w: 2.7, h: 1.7, fill: { color: b[1] }, rectRadius: 0.1 });
    T(s, b[0], { x: x + 0.15, y: 1.6, w: 2.4, h: 1.7, fontSize: 17, bold: true, color: C.background1, align: 'center', valign: 'middle' });
    if (i < 3) s.addShape(pres.ShapeType.rightArrow, { x: x + 2.72, y: 2.25, w: 0.34, h: 0.4, fill: { color: C.accent4 } }); });
  T(s, '그래서  설계자의 지식 확장이 필요하다', { x: 0.6, y: 3.7, w: 12.1, h: 0.55, fontSize: 24, bold: true, color: C.accent1, align: 'center' });
  const k = [['① 식별', '무엇을 검토해야 하는지 안다'], ['② 이해', '그 값이 왜 나오는지 근거를 이해한다'], ['③ 설명', '조건이 바뀌면 다른 분야와 공간에\n어떤 영향이 생기는지 설명한다']];
  T(s, '본 연구의 지식 확장 = 설계자가 건축 외 분야(전기·설비 등)에 대해', { x: 0.6, y: 4.45, w: 12.1, h: 0.4, fontSize: 16, color: C.accent4, align: 'center' });
  k.forEach((c, i) => card(s, 0.6 + i * 4.1, 4.95, 3.9, 1.6, c[0], c[1], { hs: 20, bs: 16, hc: C.accent1 })); }

// 3 일반 vs 특수
{ const s = content('일반 시설과 특수 시설은 다르다', '사람이 상주하는 공간이 중심인 일반 시설은 설계 각론과 법규, 사례로 어느 정도 정확하게 공간을 계획할 수 있습니다. 하지만 설비와 전기 공간의 비율이 높은 특수 시설은 그것만으로 한계가 있습니다. 데이터센터는 그 대표 사례이고, 국내 사례 연구에서 서버실 층과 기계·전기실 면적이 약 1대1로 요구됩니다.');
  card(s, 0.6, 1.5, 5.9, 2.45, '일반 시설', '사람이 상주하는 공간 중심\n발주처 요구사항, 용도별 설계 각론·법규·사례로\n공간 계획이 가능', { fill: C.background2, hs: 22, bs: 17 });
  card(s, 6.8, 1.5, 5.9, 2.45, '특수 시설', '산업시설, 방송통신시설의 데이터센터 등\n설비·전기 공간의 점유 비율이 높음\n경험과 사례만으로는 필요 공간 검토에 한계', { fill: 'E1F3EF', hs: 20, bs: 17, hc: C.accent1 });
  s.addShape(pres.ShapeType.rightArrow, { x: 6.52, y: 2.55, w: 0.25, h: 0.35, fill: { color: C.accent4 } });
  T(s, '왜 데이터센터인가', { x: 0.6, y: 4.2, w: 12, h: 0.45, fontSize: 20, bold: true, color: C.accent1 });
  const w = [['서버실 층 : 기계·전기실 면적\n≈ 1 : 1', '김재성·이원재(2018)'], ['전력·냉각·비상 전원이\n공간 규모를 직접 결정\n(조건이 수치로 표현됨)', '검증과 명시화에 적합'], ['전력 수요\n2030년까지 2배 이상', 'IEA(2025)']];
  w.forEach((c, i) => { const x = 0.6 + i * 4.1;
    s.addShape(pres.ShapeType.roundRect, { x, y: 4.75, w: 3.9, h: 1.85, fill: { color: C.accent5 }, rectRadius: 0.1 });
    T(s, c[0], { x: x + 0.2, y: 4.85, w: 3.5, h: 1.2, fontSize: 17, bold: true, color: C.background1, valign: 'middle', align: 'center' });
    T(s, c[1], { x: x + 0.2, y: 6.1, w: 3.5, h: 0.35, fontSize: 12, color: C.accent6, align: 'center' }); }); }

// 4 공간구성
{ const s = content('데이터센터의 공간 구성', '데이터센터는 서버룸에서 시작해 기준층, 층고, 서포팅 인프라, 운영동으로 확장됩니다. 제가 만든 앱 A, B, C는 이 중 층고와 설비 공간, 즉 서포팅 인프라 검토에 해당합니다.');
  img(s, 'fig_2_1_space.png', 0.6, 1.5, 12.1, 3.57);
  T(s, [{ text: '검토가 어려운 부분  ', options: { bold: true, color: C.accent1 } }, { text: '층고(볼륨)와 설비 공간(서포팅 인프라): 설비 용량에 따라 건축 규모가 바뀜', options: {} }], { x: 0.6, y: 5.4, w: 12.1, h: 0.5, fontSize: 18, color: C.text1 });
  T(s, [{ text: '본 연구의 사례  ', options: { bold: true, color: C.accent1 } }, { text: 'A 전력 이중화 · B 데이터홀 단면 · C 발전기실 급·배기', options: {} }], { x: 0.6, y: 5.95, w: 12.1, h: 0.5, fontSize: 18, color: C.text1 });
  src(s, '자료: 저자 작성 (이명진, 2021, 그림 1을 참고하여 재구성; 비율은 김재성·이원재, 2018)'); }

// 5 주요 실
{ const s = content('실마다 건축 계획 요소가 뚜렷하다', '김재성·이원재의 국내 사례 조사에서 정리한 주요 실입니다. 서버실과 항온항습실은 층고 4.5미터 이상과 이중마루가 필요하고, 수변전실과 발전기실의 적정 층고는 7미터이며 발전기실은 급·배기구 면적을 확보해야 합니다. 이처럼 용도에 맞는 스페이스 프로그램이 먼저 갖춰져야 초기 계획이 가능합니다.');
  img(s, 'fig_2_2_rooms.png', 0.6, 1.45, 8.2, 3.84);
  const f = [['서버실·항온항습실', '층고 4.5m 이상, 이중마루'], ['수변전실·발전기실', '적정 층고 7m'], ['발전기실', '급·배기구 면적 확보, 방화댐퍼']];
  f.forEach((c, i) => card(s, 9.1, 1.45 + i * 1.3, 3.6, 1.15, c[0], c[1], { hs: 16, bs: 15, hh: 0.4, hc: C.accent1 }));
  T(s, '구체적인 면적 비율보다, 용도에 맞는 스페이스 프로그램을 먼저 갖추는 것이 초기 계획의 출발점', { x: 0.6, y: 5.6, w: 12.1, h: 0.6, fontSize: 18, bold: true, color: C.text1 });
  src(s, '자료: 김재성·이원재(2018)의 표 4~6에서 인용'); }

// 6 선행연구
{ const s = content('선행연구와 연구 공백', '선행연구는 설계 지식의 구조화와 AI 설계지원, LLM과 온톨로지 기반 검토, 데이터센터의 공간구성과 냉각 전략 방향으로 이루어졌습니다. 다만 설비 용량을 산정해 공간 영향을 검토하는 도구, 그리고 실무자가 생성형 AI로 직접 만든 검토 도구의 지식을 분석한 연구는 제가 검토한 범위에서 확인되지 않았습니다.');
  const H = { bold: true, color: 'FFFFFF', fill: { color: '1B2733' }, fontSize: 15, valign: 'middle' };
  const cell = (t, o) => ({ text: t, options: Object.assign({ fontSize: 15, valign: 'middle', color: '1B2733' }, o || {}) });
  s.addTable([
    [{ text: '분야', options: H }, { text: '선행연구', options: H }, { text: '주된 연구 방향', options: H }],
    [cell('설계 지식 구조화·AI 설계지원', { bold: true }), cell('이병수 외(2016), 김혜정 외(2019), 이용주(2020)'), cell('지식의 도구화, 디자인 대안 생성, 전문가 시스템')],
    [cell('LLM·온톨로지·에이전트', { bold: true }), cell('Chen 외(2024), Dong 외(2025)'), cell('규정 검토 자동화, 설계 조정 지원')],
    [cell('데이터센터 공간구성', { bold: true }), cell('김재성·이원재(2018), 이명진(2021)'), cell('실 프로그램과 공간구성 정리')],
    [cell('데이터센터 냉각', { bold: true }), cell('Cai & Gou(2026)'), cell('냉각 전략(패시브·액티브) 종합')],
  ], { x: 0.6, y: 1.45, w: 12.1, colW: [3.6, 4.7, 3.8], rowH: 0.62, border: { type: 'solid', pt: 0.75, color: 'CBD5E1' }, margin: [0, 0.12, 0, 0.12] });
  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 4.85, w: 12.1, h: 1.75, fill: { color: 'E1F3EF' }, rectRadius: 0.1 });
  T(s, '연구 공백', { x: 0.85, y: 4.95, w: 4, h: 0.45, fontSize: 20, bold: true, color: C.accent1 });
  T(s, '설비 용량을 산정하여 조건 변경에 따른 공간 영향을 검토하는 도구,\n그리고 건축 실무자가 생성형 AI로 직접 만든 검토 도구의 지식을\n명시화·검증·공유 관점에서 분석한 연구는 확인되지 않음', { x: 0.85, y: 5.4, w: 11.6, h: 1.1, fontSize: 18, color: C.text1, valign: 'top' }); }

// 7 목적과 연구질문
{ const s = content('연구 목적과 연구 질문', '연구 목적은 초기 판단 오류로 인한 설계 변경 리스크를 줄이기 위해 바이브 코딩으로 만든 앱을 도입해 공간 설계의 정합성을 높이는 방법을 제시하는 것입니다. 연구 질문은 세 가지입니다.');
  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 1.45, w: 12.1, h: 1.45, fill: { color: C.text1 }, rectRadius: 0.1 });
  T(s, [{ text: '목적  ', options: { bold: true, color: C.accent2 } }, { text: '바이브 코딩 앱을 초기 계획설계에 도입하여\n공간 설계의 정합성을 높이고, 설계 변경 리스크를 줄이는 방법 제시', options: { color: C.background1 } }], { x: 0.9, y: 1.45, w: 11.5, h: 1.45, fontSize: 21, valign: 'middle' });
  const q = [['RQ1', '앱은 지식을 어떻게 반영하고\n명시화·검증·공유하는가', '3.4절, 4.1~4.4절'], ['RQ2', '기존 검토 방식에 비해\n무엇을 개선하는가', '4.5절'], ['RQ3', '앱들을 통합하여 지식 확장의\n완성도를 높이려면\n어떤 프레임워크가 필요한가', '5.4절']];
  q.forEach((c, i) => { const x = 0.6 + i * 4.1;
    s.addShape(pres.ShapeType.roundRect, { x, y: 3.25, w: 3.9, h: 3.2, fill: { color: C.background2 }, rectRadius: 0.1 });
    T(s, c[0], { x: x + 0.25, y: 3.4, w: 3.4, h: 0.65, fontSize: 30, bold: true, color: C.accent1 });
    T(s, c[1], { x: x + 0.25, y: 4.15, w: 3.4, h: 1.7, fontSize: 18, color: C.text1, valign: 'top' });
    T(s, '→ ' + c[2], { x: x + 0.25, y: 5.95, w: 3.4, h: 0.35, fontSize: 13, color: C.accent4 }); }); }

// 8 분석틀
{ const s = content('연구 방법: 사례분석과 분석틀 (4관점 12요소)', '방법은 문헌 조사와 사례분석을 결합한 정성적 연구입니다. 문헌에서 지식 명시화, 도구 구현, 검증, 공유·개선의 네 관점, 열두 요소의 분석틀을 도출했습니다. 소스코드, 커밋 이력, 설명서, 실행 화면을 분석하고 계산은 제가 독립적으로 계산한 값과 대조했습니다. 판정은 충족, 부분 충족, 미흡 세 단계입니다.');
  const a = [['지식 명시화', '앱이 아는 것을\n눈에 보이게 적었나?', '① 검토 영역\n② 입력 변수·단위\n③ 산식·규칙\n④ 근거·가정 표시'], ['도구 구현', '결과를 건축 공간으로\n보여 주나?', '⑤ 구현 형태\n⑥ 출력·공간 표현'], ['검증', '맞는지 확인할\n수 있나?', '⑦ 문서-코드 일치\n⑧ 기준 사례 시험\n⑨ 자체 점검'], ['공유·개선', '남도 쓰고\n고칠 수 있나?', '⑩ 배포·접근\n⑪ 의견 수집·처리\n⑫ 지식 원본 관리']];
  a.forEach((c, i) => { const x = 0.6 + i * 3.08;
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.45, w: 2.9, h: 0.65, fill: { color: C.accent1 }, rectRadius: 0.08 });
    T(s, c[0], { x, y: 1.45, w: 2.9, h: 0.65, fontSize: 19, bold: true, color: C.background1, align: 'center', valign: 'middle' });
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.2, w: 2.9, h: 3.0, fill: { color: C.background2 }, rectRadius: 0.08 });
    T(s, c[1], { x: x + 0.15, y: 2.3, w: 2.6, h: 0.95, fontSize: 16, bold: true, color: C.text1, valign: 'middle' });
    T(s, c[2], { x: x + 0.15, y: 3.3, w: 2.6, h: 1.8, fontSize: 15, color: C.text1, valign: 'top' }); });
  T(s, '근거: 소스코드 · 커밋 이력 · 설명서 · 실행 화면, 계산은 연구자가 독립 계산하여 대조', { x: 0.6, y: 5.5, w: 12.1, h: 0.4, fontSize: 16, color: C.accent4 });
  T(s, '판정   ● 충족    ◐ 부분 충족    ○ 미흡', { x: 0.6, y: 6.05, w: 12.1, h: 0.5, fontSize: 22, bold: true, color: C.text1 }); }

// 9 사례 앱 3개 + 개발 과정
{ const s = content('분석 대상: 개발 중인 앱 3개', '세 앱은 제가 협력업체와 협의한 검토 방법과 장비 사양을 로직으로 옮겨 Claude Code로 만든 웹앱입니다. GitHub에 올리고 Vercel로 배포하며, 동료 실무자가 사용하고 협력업체가 결과를 보며 의견을 주는 시범 적용 단계입니다. 완성품이 아니라 개발이 계속되는 단계여서 2026년 11월 버전을 분석했습니다.');
  const lk = (t, u) => ({ text: t, options: { color: '0563C1', underline: { style: 'sng' }, hyperlink: { url: u } } });
  const ap = [['A  전력 이중화 검토', '변압기·전기실 수 → 평면', '73 커밋 · 2026.08~11', 'power-redundancy-calculator.vercel.app', 'https://power-redundancy-calculator.vercel.app/'], ['B  데이터홀 단면 검토', '랙·냉각 장비 → 층고·건물 규모', '17 커밋 · 2026.05~10', 'datahall-section-check.vercel.app', 'https://datahall-section-check.vercel.app/'], ['C  발전기실 급·배기', '급·배기 풍량 → 입면·층고', '5 커밋 · 2026.06~09', 'generator-calc.vercel.app', 'https://generator-calc.vercel.app/']];
  ap.forEach((c, i) => { const x = 0.6 + i * 4.1;
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.45, w: 3.9, h: 2.35, fill: { color: C.background2 }, rectRadius: 0.1 });
    T(s, c[0], { x: x + 0.2, y: 1.55, w: 3.5, h: 0.5, fontSize: 19, bold: true, color: C.accent1 });
    T(s, c[1], { x: x + 0.2, y: 2.1, w: 3.5, h: 0.5, fontSize: 16, color: C.text1 });
    T(s, c[2], { x: x + 0.2, y: 2.65, w: 3.5, h: 0.35, fontSize: 14, color: C.accent4 });
    T(s, [lk(c[3], c[4])], { x: x + 0.2, y: 3.15, w: 3.55, h: 0.4, fontSize: 11.5 }); });
  T(s, '개발과 운영 과정', { x: 0.6, y: 4.1, w: 6, h: 0.4, fontSize: 18, bold: true, color: C.accent1 });
  const p = ['협력업체 협의\n장비 사양', 'Claude Code\n바이브 코딩', 'GitHub\n커밋', 'Vercel\n웹 배포', '시범 적용\n의견·개선'];
  p.forEach((t, i) => { const x = 0.6 + i * 2.45;
    s.addShape(pres.ShapeType.roundRect, { x, y: 4.6, w: 2.15, h: 1.05, fill: { color: i === 4 ? C.accent1 : C.accent5 }, rectRadius: 0.08 });
    T(s, t, { x, y: 4.6, w: 2.15, h: 1.05, fontSize: 15, bold: true, color: C.background1, align: 'center', valign: 'middle' });
    if (i < 4) s.addShape(pres.ShapeType.rightArrow, { x: x + 2.17, y: 4.95, w: 0.25, h: 0.32, fill: { color: C.accent4 } }); });
  T(s, '시범 적용(pilot) 단계: 동료 실무자가 사용하고, 협력업체가 검토 결과를 보며 의견을 줌. 공식 인증 전, 2026년 11월 버전 분석', { x: 0.6, y: 5.95, w: 12.1, h: 0.7, fontSize: 15, color: C.accent4 }); }

// 10~12 사례
const caseSlide = (title, notes, f, ar, bul, weak, source) => {
  const s = content(title, notes);
  const h = 4.9, w = Math.min(7.2, h * ar);
  img(s, f, 0.6, 1.4, w, h);
  const x = 0.6 + w + 0.35, W = 12.7 - x;
  T(s, '잘 되어 있는 점', { x, y: 1.4, w: W, h: 0.4, fontSize: 18, bold: true, color: C.accent1 });
  bullets(s, bul, x, 1.9, W, 2.8, 15);
  T(s, '보완할 점', { x, y: 4.75, w: W, h: 0.4, fontSize: 18, bold: true, color: 'B7791F' });
  bullets(s, weak, x, 5.2, W, 1.4, 14);
  src(s, source); return s; };
caseSlide('사례 A: 전력 이중화 검토 앱', '사례 A는 이중화 방식과 변압기 용량으로 전기실 수를 계산하고 계통을 그림으로 보여 줍니다. 계산식이 화면에 그대로 나오고, 제가 363개 결과를 직접 계산과 대조해 모두 일치했습니다. 보완할 점은 일부 기준값의 출처가 코드 안에만 있고 의견 이력이 개발자 쪽에 남지 않는다는 것입니다.', 'fig_app1_b_diagram_crop.png', 1.27,
  ['이중화 7종 + 직접 입력, 단위와 함께 입력', '계산식을 값과 함께 표시\n4,000 × 0.99 × 0.9 × 0.96 = 3,421.4kW', '33개 조건, 363개 결과를 직접 계산과 대조 → 불일치 없음', '부하율 100% 초과 등 자체 경고'],
  ['일부 기준값의 출처가 코드 메모에만 있음', '의견이 브라우저에만 저장됨', '냉각 산식이 사례 B와 중복'],
  '자료: 저자 개발 앱 배포본(power-redundancy-calculator.vercel.app) 화면 캡처, 2026.11.18');
caseSlide('사례 B: 데이터홀 단면 검토 앱', '사례 B는 랙과 냉각 장비를 배치했을 때 층고와 보 아래 공간이 충분한지를 단면과 평면으로 보여 줍니다. 필요 풍량과 보 아래 환기 공간을 단계별로 보여 주고, 풍속이 초과되면 붉게 경고합니다. 보완할 점은 계수와 기준값의 출처가 화면에 없다는 것입니다.', 'fig_caseB_crop.png', 1.557,
  ['입면·단면·평면을 자동 작성, 3D·DXF 내보내기', '단계별 계산: 필요 풍량 2,541,704CMH', '보 아래 환기 공간 8.3m − 6.33m = 1.97m', '풍속 초과(5.85m/s)와 기둥 간섭 경고'],
  ['계수와 기준값의 출처가 화면에 없음', '시험 코드 없음', '배포본에서 서버 수집 기능은 동작하지 않음'],
  '자료: 저자 개발 앱 저장소(github.com/sangchuls/datahall-section-check) 실행 화면 캡처, 2026.11.20');
caseSlide('사례 C: 발전기실 급·배기 계산 앱', '사례 C는 제조사 모델을 고르면 급기와 배기 개구면적을 계산하고 외기 그릴과 소음기, 댐퍼의 배치를 그림으로 보여 줍니다. 카탈로그 출처 쪽수를 화면에 표시해 근거 표시가 세 사례 중 가장 명확하고, 계산 함수를 한 파일에 두어 지식 원본 관리도 가장 명확합니다. 의견 수집 기능이 없는 것이 보완점입니다.', 'fig_caseC_crop.png', 1.485,
  ['제조사 4곳 모델 선택 → 값 자동 입력', '카탈로그 출처 쪽수를 화면에 표시 (근거 표시 가장 명확)', 'Cummins 3,000kW 8대: 급기 69.12㎡ · 배기 54.72㎡ 일치', '카탈로그 7.20㎡ vs 역산 7.17㎡ 나란히 점검', '계산 함수 한 파일 → 화면과 API가 공유'],
  ['의견 수집 기능 없음', '시험 코드 없음'],
  '자료: 저자 개발 앱 저장소(github.com/sangchuls/generator-calc) 실행 화면 캡처, 2026.11.20');

// 13 종합
{ const s = content('종합 비교: 모두 잘 갖췄고, 보강할 점은 다르다', '판정 결과입니다. 충족은 A가 7개, B가 6개, C가 8개입니다. 공통 강점은 산식을 화면에 드러내고 결과를 도면으로 연결하며 스스로 경고하는 것입니다. 공통 보완점은 기준 사례 시험이 저장소에 없다는 것이고, 근거 표시는 C가 충족, A가 부분, B가 미흡으로 차이가 났습니다.');
  s.addChart(pres.charts.BAR, [
    { name: '충족', labels: ['사례 A', '사례 B', '사례 C'], values: [7, 6, 8] },
    { name: '부분 충족', labels: ['사례 A', '사례 B', '사례 C'], values: [4, 4, 2] },
    { name: '미흡', labels: ['사례 A', '사례 B', '사례 C'], values: [0, 1, 1] }], {
    x: 0.6, y: 1.4, w: 6.2, h: 4.6, barDir: 'bar', barGrouping: 'stacked', chartColors: ['0F766E', 'F2A93B', 'B23A48'],
    showTitle: true, title: '판정 결과 (11개 요소)', titleFontSize: 16, titleColor: '1B2733', titleFontFace: '+mn-lt',
    showValue: true, dataLabelFormatCode: '0;;;', dataLabelPosition: 'ctr', dataLabelColor: 'FFFFFF', dataLabelFontSize: 16, dataLabelFontFace: '+mn-lt',
    showLegend: true, legendPos: 'b', legendFontSize: 14, legendColor: '1B2733', legendFontFace: '+mn-lt',
    catAxisLabelColor: '1B2733', catAxisLabelFontSize: 15, catAxisLabelFontFace: '+mn-lt', catAxisOrientation: 'maxMin', valAxisHidden: true, valGridLine: { style: 'none' }, catGridLine: { style: 'none' } });
  const k = [['공통 강점', '산식을 화면에 표시 · 결과를 도면으로 연결 · 스스로 경고', C.accent1], ['공통 보완', '기준 사례 시험이 저장소에 없음', 'B7791F'], ['차이', '근거 표시: C 충족 · A 부분 · B 미흡\n지식 원본: C 한 파일, A·B는 냉각 산식 중복', C.accent3]];
  k.forEach((c, i) => card(s, 7.2, 1.4 + i * 1.55, 5.5, 1.4, c[0], c[1], { hs: 18, bs: 15, hh: 0.4, hc: c[2] }));
  T(s, '※ 2026년 11월 시점 버전에 대한 판정, 개발이 계속되는 단계', { x: 0.6, y: 6.2, w: 12.1, h: 0.4, fontSize: 14, color: C.accent4 }); }

// 14 기존 방식 대비
{ const s = content('기존 방식 대비 앱의 개선', '기존에는 비슷한 사례에서 유추하거나 직관으로 예상하고, 협력업체의 협조가 가능할 때만 도움을 받았습니다. 앱은 검토 로직을 정량적이고 일관된 규칙으로 고정해 같은 입력에 같은 결과를 내고, 조건이 바뀌면 바로 다시 계산합니다. 다만 이 비교는 저자의 실무 경험에 근거한 정성적 비교이며 소요 시간과 정확도의 측정은 후속 연구입니다.');
  card(s, 0.6, 1.45, 5.4, 2.6, '기존 방식', '유사 사례에서 유추\n직관으로 예상\n협력업체 협조가 가능할 때만 도움', { fill: C.background2, hs: 22, bs: 18, hc: C.accent4 });
  s.addShape(pres.ShapeType.rightArrow, { x: 6.1, y: 2.5, w: 0.5, h: 0.5, fill: { color: C.accent1 } });
  card(s, 6.7, 1.45, 6.0, 2.6, '앱 도입 후', '정량적이고 일관된 검토 규칙\n같은 입력 → 같은 결과, 근거와 산식 표시\n조건이 바뀌면 즉시 재계산', { fill: 'E1F3EF', hs: 22, bs: 18, hc: C.accent1 });
  const r = [['A 전력 이중화', '전기실 규모·개수 → 평면'], ['B 데이터홀', '냉각 장비 대수·배치 → 건물 규모'], ['C 발전기 풍량', '그릴창 규모·건물 높이 → 입면·층고']];
  r.forEach((c, i) => { const x = 0.6 + i * 4.1;
    s.addShape(pres.ShapeType.roundRect, { x, y: 4.4, w: 3.9, h: 1.3, fill: { color: C.accent5 }, rectRadius: 0.1 });
    T(s, c[0], { x: x + 0.2, y: 4.5, w: 3.5, h: 0.45, fontSize: 18, bold: true, color: C.background1 });
    T(s, c[1], { x: x + 0.2, y: 5.0, w: 3.5, h: 0.6, fontSize: 15, color: C.accent6, valign: 'top' }); });
  T(s, '앱의 결과는 개략 검토용이며 업체의 상세 검토를 대체하지 않는다', { x: 0.6, y: 6.0, w: 12.1, h: 0.4, fontSize: 15, color: C.accent4 }); }

// 15 지식 유형
{ const s = content('앱이 담은 지식, 남은 지식', '지식 유형으로 보면 앱은 입력값과 기준인 선언적 지식, 산식과 계산 순서인 절차적 지식을 명시화했습니다. 하지만 앱 사이의 조건과 연결 관계인 상황적 지식, 무엇을 어떤 순서로 검토할지의 판단인 전략적 지식은 설계자의 경험에 남아 있습니다. 이것이 앱의 통합과 온톨로지가 필요한 이유입니다.');
  const k = [['선언적 지식', '입력값 · 기준', '앱에 담김', true], ['절차적 지식', '산식 · 계산 순서', '앱에 담김', true], ['상황적 지식', '조건이 바뀌면 다른 설비와 공간이\n어떻게 영향받는가', '설계자에게 남음 → 온톨로지', false], ['전략적 지식', '무엇을 어떤 순서로\n검토할 것인가', '설계자에게 남음 → AI 에이전트', false]];
  k.forEach((c, i) => { const x = 0.6 + i * 3.08;
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.5, w: 2.9, h: 3.9, fill: { color: c[3] ? 'E1F3EF' : 'FEF3C7' }, rectRadius: 0.1 });
    T(s, c[0], { x: x + 0.15, y: 1.65, w: 2.6, h: 0.55, fontSize: 21, bold: true, color: c[3] ? C.accent1 : 'B7791F', align: 'center' });
    T(s, c[1], { x: x + 0.15, y: 2.4, w: 2.6, h: 1.5, fontSize: 16, color: C.text1, align: 'center', valign: 'top' });
    T(s, c[3] ? '✔ ' + c[2] : '▶ ' + c[2], { x: x + 0.15, y: 4.2, w: 2.6, h: 1.0, fontSize: 15, bold: true, color: c[3] ? C.accent1 : 'B7791F', align: 'center', valign: 'middle' }); });
  T(s, '앱들을 통합하는 단계에서 상황적·전략적 지식을 다룬다', { x: 0.6, y: 5.8, w: 12.1, h: 0.5, fontSize: 20, bold: true, color: C.text1, align: 'center' }); }

// 16 제언
{ const s = content('제언: 앱의 신뢰성과 일관성을 높이는 4가지', '통합에 앞서 앱마다 지식 원본을 단일화하고, 화면의 값마다 근거 수준을 표시하고, 독립 계산 대조를 기준 사례 시험으로 저장소에 넣고, 여러 앱을 한 포털에서 제공하며 의견 이력을 모으는 것을 제언합니다.');
  const k = [['1', '지식 원본 단일화', '산식·기준값·출처를 문서로 두고 여러 앱이 공유\n(A·B의 중복 냉각 산식 해소)'], ['2', '근거 수준 표시', '값마다 확인·가정 구분, 확인 시기와 출처 유형 기록'], ['3', '기준 사례 시험', '독립 계산 대조를 시험으로 만들어\n저장소에 포함'], ['4', '공통 포털과 의견 이력', '앱을 한 곳에서 제공, 의견을 개발자 쪽에 모아\n처리 상태와 이력 관리']];
  k.forEach((c, i) => { const x = 0.6 + (i % 2) * 6.15, y = 1.5 + Math.floor(i / 2) * 2.55;
    s.addShape(pres.ShapeType.roundRect, { x, y, w: 5.95, h: 2.3, fill: { color: C.background2 }, rectRadius: 0.1 });
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.25, y: y + 0.3, w: 0.75, h: 0.75, fill: { color: C.accent1 } });
    T(s, c[0], { x: x + 0.25, y: y + 0.3, w: 0.75, h: 0.75, fontSize: 24, bold: true, color: C.background1, align: 'center', valign: 'middle' });
    T(s, c[1], { x: x + 1.25, y: y + 0.3, w: 4.5, h: 0.75, fontSize: 20, bold: true, color: C.text1, valign: 'middle' });
    T(s, c[2], { x: x + 0.3, y: y + 1.25, w: 5.4, h: 0.95, fontSize: 15, color: C.text1, valign: 'top' }); }); }

// 17 프레임워크와 로드맵
{ const s = content('후속 연구: 앱 통합과 온톨로지 기반 AI 에이전트', '연구의 최종 목표는 앱을 통합하고 온톨로지를 도입해 AI 에이전트가 앱을 직접 호출하며 결론과 함께 근거와 산술식을 제시하도록 하는 것입니다. 계산은 앱이 맡고, 온톨로지는 조건과 연결 관계를, 에이전트는 검토의 순서와 판단을 맡습니다. 본 과제는 1단계이고 2, 3단계는 연구 계획입니다. 효과 검증은 후속 연구입니다.');
  img(s, 'fig_5_1_agent.png', 0.6, 1.4, 12.1, 3.84);
  const rm = [['① 앱 개발 과정·원리 분석', '본 과제', true], ['② 앱 통합', '계획', false], ['③ 온톨로지화·AI 에이전트', '계획', false]];
  rm.forEach((r, i) => { const x = 0.6 + i * 4.2;
    s.addShape(pres.ShapeType.roundRect, { x, y: 5.55, w: 3.7, h: 0.75, fill: { color: r[2] ? C.accent1 : C.accent4 }, rectRadius: 0.08 });
    T(s, [{ text: r[0], options: { bold: true } }, { text: '  (' + r[1] + ')', options: {} }], { x: x + 0.1, y: 5.55, w: 3.5, h: 0.75, fontSize: 15, color: C.background1, valign: 'middle', align: 'center' });
    if (i < 2) s.addShape(pres.ShapeType.rightArrow, { x: x + 3.75, y: 5.75, w: 0.4, h: 0.35, fill: { color: C.text1 } }); });
  src(s, '자료: 저자 작성', 6.55); }

// 18 결론
{ const s = pres.addSlide({ masterName: 'DARK' });
  s.addText('결론', { placeholder: 'title', y: 0.45, h: 1.0, fontSize: 36 });
  const a = [['RQ1', '협력업체 협의와 장비 사양을 로직으로 옮겨 앱을 개발·배포하며, 산식과 결과는 명시되나 근거 표시와 시험은 사례별 차이'], ['RQ2', '유추와 직관을 정량적이고 일관된 검토 규칙으로 바꾸어 검토 방식을 체계화'], ['RQ3', '온톨로지와 AI 에이전트가 앱을 제어하고 결론과 근거·산식을 제시하는 통합 프레임워크 제안']];
  a.forEach((c, i) => { const y = 2.05 + i * 1.05;
    s.addShape(pres.ShapeType.roundRect, { x: 0.7, y, w: 11.9, h: 0.9, fill: { color: C.accent5 }, rectRadius: 0.1 });
    T(s, c[0], { x: 0.95, y, w: 1.1, h: 0.9, fontSize: 24, bold: true, color: C.accent2, valign: 'middle' });
    T(s, c[1], { x: 2.15, y, w: 10.2, h: 0.9, fontSize: 16, color: C.background1, valign: 'middle' }); });
  T(s, [{ text: '기대 효과  ', options: { bold: true, color: C.accent2 } }, { text: '검토 속도 ↑ · 정량적 근거 ↑ · 초기 계획의 정합성 ↑ · 설계 변경 리스크 ↓', options: { color: C.background1 } }], { x: 0.7, y: 5.35, w: 11.9, h: 0.5, fontSize: 19 });
  T(s, [{ text: '한계  ', options: { bold: true, color: C.accent2 } }, { text: '시범 적용 단계의 개발 중인 앱(2026.11 버전) · 효과(정확도·소요 시간)는 후속 연구에서 검증', options: { color: C.accent6 } }], { x: 0.7, y: 5.95, w: 11.9, h: 0.5, fontSize: 16 });
  T(s, '감사합니다', { x: 0.7, y: 6.55, w: 11.9, h: 0.6, fontSize: 28, bold: true, color: C.background1, align: 'right' });
  s.addNotes('정리하면, 협력업체와의 협의와 장비 사양을 로직으로 옮긴 앱은 유추와 직관에 의존하던 검토를 정량적이고 일관된 규칙으로 바꿨습니다. 앱 사이의 연결과 검토 순서는 온톨로지와 AI 에이전트가 맡는 통합 프레임워크를 제안했습니다. 앱은 시범 적용 단계이고 효과 검증은 후속 연구입니다. 감사합니다.'); }

// 19 접속 주소
{ const s = content('배포한 웹앱 접속 주소', '세 앱은 아래 주소에서 바로 접속할 수 있습니다. 시범 적용 단계의 개발 중인 앱이며 2026년 11월 버전입니다.');
  const lk = (t, u) => ({ text: t, options: { color: '0563C1', underline: { style: 'sng' }, hyperlink: { url: u } } });
  const ap = [['A', '전력 이중화 검토 앱', '변압기·전기실 수 → 평면', 'https://power-redundancy-calculator.vercel.app/'], ['B', '데이터홀 단면 검토 앱', '랙·냉각 장비 → 층고·건물 규모', 'https://datahall-section-check.vercel.app/'], ['C', '발전기실 급·배기 계산 앱', '급·배기 풍량 → 입면·층고', 'https://generator-calc.vercel.app/']];
  ap.forEach((c, i) => { const y = 1.5 + i * 1.65;
    s.addShape(pres.ShapeType.roundRect, { x: 0.6, y, w: 12.1, h: 1.4, fill: { color: C.background2 }, rectRadius: 0.1 });
    s.addShape(pres.ShapeType.ellipse, { x: 0.9, y: y + 0.3, w: 0.8, h: 0.8, fill: { color: C.accent1 } });
    T(s, c[0], { x: 0.9, y: y + 0.3, w: 0.8, h: 0.8, fontSize: 28, bold: true, color: C.background1, align: 'center', valign: 'middle' });
    T(s, c[1], { x: 2.0, y: y + 0.18, w: 10.4, h: 0.5, fontSize: 22, bold: true, color: C.text1, valign: 'middle' });
    T(s, c[2], { x: 2.0, y: y + 0.68, w: 10.4, h: 0.3, fontSize: 14, color: C.accent4 });
    T(s, [lk(c[3], c[3])], { x: 2.0, y: y + 0.98, w: 10.4, h: 0.35, fontSize: 18 }); });
  T(s, '※ 시범 적용 단계의 개발 중인 앱 (2026년 11월 버전)  ·  주소를 클릭하면 연결됩니다', { x: 0.6, y: 6.5, w: 12.1, h: 0.4, fontSize: 14, color: C.accent4 }); }

(async () => { await pres.writeFile({ fileName: 'presentation2.pptx' }); await applyTheme('presentation2.pptx', THEME); console.log('written'); })();
