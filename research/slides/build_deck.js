// 발표 자료(4장, 4분): 최종 보고서 내용만 사용. 실행: node build_deck.js
const pptxgen = require('pptxgenjs');
const SK = '/root/.claude/skills/synced/b09188c3-4889-4bd8-b551-d0a19e63b56f_facbcb1f-1d3a-4bc4-bdfb-98f8c08ef25b/pptx';
const { applyTheme } = require(SK + '/scripts/apply_theme.js');

const THEME = {
  name: 'DC Study', headFontFace: 'Malgun Gothic', bodyFontFace: 'Malgun Gothic',
  colors: { dk1: '1B2733', lt1: 'FFFFFF', dk2: '0F766E', lt2: 'F1F5F9',
            accent1: '0F766E', accent2: 'F2A93B', accent3: 'B23A48', accent4: '64748B', accent5: '2B3B4C', accent6: 'CBD5E1',
            hlink: '0F766E', folHlink: '64748B' }
};
const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';            // 13.33 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = '건축 초기 계획설계의 지식 확장을 위한 Micro-App 분석';
pres.author = '임상철';
const C = pres.SchemeColor;

pres.defineSlideMaster({
  title: 'DARK', background: { color: '1B2733' },
  objects: [{ placeholder: { options: { name: 'title', type: 'title', x: 0.7, y: 0.7, w: 8.3, h: 1.7, fontSize: 36, bold: true, color: C.background1, valign: 'top', align: 'left', margin: 0 }, text: '' } }]
});
pres.defineSlideMaster({
  title: 'CONTENT', background: { color: 'FFFFFF' },
  objects: [
    { placeholder: { options: { name: 'title', type: 'title', x: 0.6, y: 0.35, w: 12.1, h: 0.85, fontSize: 34, bold: true, color: C.text1, valign: 'middle', align: 'left', margin: 0 }, text: '' } },
    { text: { text: '한양대학교 공학대학원 | 건축계획연구', options: { x: 0.6, y: 7.0, w: 6, h: 0.3, fontSize: 10, color: C.accent4, margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 7.0, w: 0.5, h: 0.3, fontSize: 10, color: C.accent4, align: 'right' }
});

const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0 }, o));

// ================= Slide 1: 배경·문제의식 (일반 시설 → 특수 시설 → 데이터센터) =================
{
  const s = pres.addSlide({ masterName: 'DARK' });
  s.addText('건축 초기 계획설계의 지식 확장을 위한 Micro-App 분석', { placeholder: 'title' });
  T(s, '바이브 코딩으로 개발한 데이터센터 검토 앱 3개 사례를 중심으로', { x: 0.7, y: 2.2, w: 8.5, h: 0.5, fontSize: 18, color: C.accent6, objectName: '부제' });
  T(s, '1  초기 계획설계의 문제', { x: 0.7, y: 2.95, w: 6, h: 0.4, fontSize: 16, bold: true, color: C.accent2, objectName: '1단계 라벨' });
  const bx = [['협업·지원이 어려움', '경험·사례 유추로 직접 판단'], ['전문 검토와 불일치', '업체 검토 결과와 달라짐'], ['설계 변경 리스크', '공간계획 재작업 → 지식 확장 필요']];
  bx.forEach((b, i) => {
    const x = 0.7 + i * 4.15;
    s.addShape(pres.ShapeType.roundRect, { x, y: 3.4, w: 3.5, h: 1.15, fill: { color: C.accent5 }, rectRadius: 0.1, objectName: `문제 상자 ${i + 1}` });
    T(s, b[0], { x: x + 0.2, y: 3.48, w: 3.1, h: 0.45, fontSize: 18, bold: true, color: C.background1, objectName: `문제 제목 ${i + 1}` });
    T(s, b[1], { x: x + 0.2, y: 3.95, w: 3.1, h: 0.5, fontSize: 14, color: C.accent6, valign: 'top', objectName: `문제 설명 ${i + 1}` });
    if (i < 2) s.addShape(pres.ShapeType.rightArrow, { x: x + 3.6, y: 3.8, w: 0.45, h: 0.38, fill: { color: C.accent4 }, objectName: `화살표 ${i + 1}` });
  });
  T(s, '2  왜 데이터센터인가', { x: 0.7, y: 4.75, w: 6, h: 0.4, fontSize: 16, bold: true, color: C.accent2, objectName: '2단계 라벨' });
  const cmp = [['일반 시설', '사람이 상주하는 공간 중심\n설계 각론·법규·사례로 기획 가능', C.accent4], ['특수 시설 (산업·방송통신·데이터센터)', '설비·전기 공간이 압도적\n메인 층 : 기계·전기실 ≈ 1 : 1', C.accent1]];
  cmp.forEach((c, i) => {
    const x = 0.7 + i * 4.25;
    s.addShape(pres.ShapeType.roundRect, { x, y: 5.2, w: 4.0, h: 1.1, fill: { color: c[2] }, rectRadius: 0.1, objectName: `비교 상자 ${i + 1}` });
    T(s, c[0], { x: x + 0.15, y: 5.25, w: 3.7, h: 0.38, fontSize: 15, bold: true, color: C.background1, objectName: `비교 제목 ${i + 1}` });
    T(s, c[1], { x: x + 0.15, y: 5.63, w: 3.7, h: 0.62, fontSize: 13, color: C.background1, valign: 'top', objectName: `비교 설명 ${i + 1}` });
  });
  s.addShape(pres.ShapeType.roundRect, { x: 9.3, y: 4.75, w: 3.4, h: 1.55, fill: { color: C.accent5 }, rectRadius: 0.12, objectName: '통계 카드' });
  T(s, '415 → 945 TWh', { x: 9.3, y: 4.8, w: 3.4, h: 0.65, fontSize: 28, bold: true, color: C.accent2, align: 'center', valign: 'middle', objectName: '통계 수치' });
  T(s, '데이터센터 전력 소비 2024→2030년 전망', { x: 9.4, y: 5.45, w: 3.2, h: 0.45, fontSize: 12, color: C.background1, align: 'center', valign: 'top', objectName: '통계 설명' });
  T(s, '자료: IEA(2025); 비율은 김재성·이원재(2018)', { x: 9.4, y: 5.9, w: 3.2, h: 0.3, fontSize: 10, color: C.accent6, align: 'center', objectName: '통계 출처' });
  T(s, '목적: 바이브 코딩 앱을 초기 계획설계에 도입하여 공간 정합성을 높이고 설계 변경 리스크를 줄이는 방법 제시', { x: 0.7, y: 6.45, w: 12, h: 0.4, fontSize: 15, bold: true, color: C.background1, objectName: '연구 목적' });
  T(s, '공학대학원 건축학전공  |  임상철 (2026160492)  |  담당교수 김영주  |  2026년 11월', { x: 0.7, y: 6.95, w: 12, h: 0.3, fontSize: 12, color: C.accent6, objectName: '인적사항' });
  s.addNotes('초기 계획설계에서는 구조, 기계, 전기 같은 건축 외 분야의 협업과 지원을 받기 어려워, 건축가가 경험과 사례로 직접 판단합니다. 이 판단이 나중에 업체의 전문 검토와 달라지면 공간계획을 다시 해야 하는 설계 변경 리스크가 생깁니다. 사람이 상주하는 공간이 중심인 일반 시설은 설계 각론과 법규, 사례로 어느 정도 정확하게 기획할 수 있습니다. 하지만 산업시설이나 방송통신시설의 데이터센터처럼 설비·전기 공간이 압도적인 특수 시설은 그것만으로 한계가 있습니다. 국내 사례 연구에 따르면 데이터센터는 서버실 층과 기계·전기실의 면적이 약 1 대 1로 요구됩니다. 전력 수요도 2024년 약 415TWh에서 2030년 945TWh로 늘 전망입니다. 그래서 본 연구는 바이브 코딩으로 만든 검토 앱을 초기 계획설계에 도입해 공간 정합성을 높이고 설계 변경 리스크를 줄이는 방법을 제시합니다. (약 1분)');
}

// ================= Slide 2: 선행연구 공백 =================
{
  const s = pres.addSlide({ masterName: 'CONTENT' });
  s.addText('선행연구와 연구 공백', { placeholder: 'title' });
  const H = { bold: true, color: 'FFFFFF', fill: { color: '1B2733' }, fontSize: 15, valign: 'middle' };
  const cell = (t, o) => ({ text: t, options: Object.assign({ fontSize: 14, valign: 'middle', color: '1B2733' }, o || {}) });
  const rows = [
    [{ text: '연구 분야', options: H }, { text: '선행연구', options: H }, { text: '한계', options: H }],
    [cell('설계 지식 구조화·AI 설계지원', { bold: true }), cell('이병수 외(2016), 김혜정 외(2019), 이용주(2020)'), cell('건축 내부·형태·친환경 분야에 한정')],
    [cell('LLM·온톨로지·에이전트', { bold: true }), cell('Chen 외(2024), Dong 외(2025)'), cell('규정 해석·BIM 조작 중심')],
    [cell('데이터센터 공간구성', { bold: true }), cell('김재성·이원재(2018), 이명진(2021)'), cell('실 구성 정리, 용량 산정·변경 대응 도구 없음')],
    [cell('데이터센터 냉각', { bold: true }), cell('Cai & Gou(2026)'), cell('냉각 효율 중심, 전력·공간 연결 없음')],
    [cell('본 연구', { bold: true, color: '0F766E', fill: { color: 'E1F3EF' } }), cell('실무자가 AI로 만든 검토 앱 3개의 사례분석', { bold: true, fill: { color: 'E1F3EF' } }), cell('설비 용량 산정 → 공간 계획 연결', { bold: true, fill: { color: 'E1F3EF' } })],
  ];
  s.addTable(rows, { x: 0.6, y: 1.45, w: 12.1, colW: [3.4, 4.6, 4.1], rowH: 0.68, border: { type: 'solid', pt: 0.75, color: 'CBD5E1' }, margin: [0, 0.12, 0, 0.12], objectName: '선행연구 비교표' });
  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 5.85, w: 12.1, h: 0.95, fill: { color: C.background2 }, rectRadius: 0.1, objectName: '공백 상자' });
  T(s, [{ text: '연구 공백  ', options: { bold: true, color: C.accent1 } }, { text: '설비 용량 산정을 공간 계획으로 잇는, 건축 실무자가 생성형 AI로 직접 만든 검토 도구의 지식을 명시화·검증·공유 관점에서 분석한 연구는 확인되지 않음', options: { color: C.text1 } }], { x: 0.85, y: 5.9, w: 11.6, h: 0.85, fontSize: 16, valign: 'middle', objectName: '공백 문장' });
  s.addNotes('선행연구는 네 갈래입니다. 설계 지식의 구조화와 AI 설계지원 연구는 건축 내부나 형태, 친환경 분야에 한정되고, LLM과 온톨로지, 에이전트 연구는 규정 해석과 BIM 조작에 집중합니다. 데이터센터 분야에서는 국내 사례의 실 프로그램과 공간구성이 정리되었고, 냉각 전략도 종합 정리되었습니다. 하지만 설비 용량을 산정해 공간 계획으로 연결하는 검토 도구는 다루지 않았습니다. 특히 건축 실무자가 생성형 AI로 직접 만든 도구의 지식을 명시화, 검증, 공유 관점에서 분석한 연구는 제가 검토한 범위에서 확인되지 않았습니다. 이것이 본 연구의 공백입니다. (약 50초)');
}

// ================= Slide 3: 연구 질문과 연구 모형 =================
{
  const s = pres.addSlide({ masterName: 'CONTENT' });
  s.addText('연구 질문과 연구 모형', { placeholder: 'title' });
  const rq = ['RQ1  앱은 지식을 어떻게 반영·검증·공유하는가', 'RQ2  기존 검토 방식 대비 무엇을 개선하는가', 'RQ3  통합에 필요한 프레임워크는 무엇인가'];
  rq.forEach((r, i) => {
    const x = 0.6 + i * 4.1;
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.35, w: 3.9, h: 0.62, fill: { color: C.accent1 }, rectRadius: 0.08, objectName: `연구 질문 ${i + 1}` });
    T(s, r, { x: x + 0.12, y: 1.35, w: 3.66, h: 0.62, fontSize: 12, bold: true, color: C.background1, valign: 'middle', objectName: `연구 질문 문구 ${i + 1}` });
  });
  const cols = [
    { h: '이론적 근거', items: ['지식 창조 이론\n(Nonaka & Takeuchi, 1995)', '설계 지식의 도구화\n(이병수 외, 2016)', '설계과학연구 방법론\n(Peffers 외, 2007)'] },
    { h: '분석틀 (4관점·12요소)', items: ['지식 명시화  ①~④', '도구 구현  ⑤⑥', '검증  ⑦~⑨', '공유·개선  ⑩~⑫'] },
    { h: '사례분석 (3개)', items: ['A  전력 이중화 → 전기실', 'B  데이터홀 → 층고·규모', 'C  발전기 급·배기 → 입면', 'Claude Code → GitHub\n→ Vercel 배포·개선'] },
    { h: '결과와 제언', items: ['요소별 판정  ● ◐ ○', '기존 방식 대비 개선점', '통합 프레임워크 제안'] },
  ];
  const W = 2.6, G = 0.567, Y = 2.2, HH = 3.45;
  cols.forEach((c, i) => {
    const x = 0.6 + i * (W + G);
    s.addShape(pres.ShapeType.roundRect, { x, y: Y, w: W, h: HH, fill: { color: C.background2 }, rectRadius: 0.1, objectName: `열 ${i + 1} 배경` });
    s.addShape(pres.ShapeType.roundRect, { x, y: Y, w: W, h: 0.6, fill: { color: i === 1 ? C.accent1 : C.text1 }, rectRadius: 0.1, objectName: `열 ${i + 1} 머리` });
    T(s, c.h, { x: x + 0.1, y: Y, w: W - 0.2, h: 0.6, fontSize: 15, bold: true, color: C.background1, align: 'center', valign: 'middle', objectName: `열 ${i + 1} 제목` });
    const gap = 2.75 / c.items.length;
    c.items.forEach((it, k) => {
      T(s, it, { x: x + 0.15, y: Y + 0.68 + k * gap, w: W - 0.3, h: gap - 0.06, fontSize: 13.5, color: C.text1, valign: 'middle', bold: i === 1, objectName: `열 ${i + 1} 항목 ${k + 1}` });
    });
    if (i < 3) s.addShape(pres.ShapeType.rightArrow, { x: x + W + 0.08, y: Y + 1.5, w: G - 0.16, h: 0.4, fill: { color: C.accent4 }, objectName: `화살표 ${i + 1}` });
  });
  const rm = [['① 앱 개발 과정·원리 분석', '본 과제', true], ['② 앱 통합', '계획', false], ['③ 온톨로지화·AI 에이전트', '계획', false]];
  rm.forEach((r, i) => {
    const x = 0.6 + i * 4.2;
    s.addShape(pres.ShapeType.roundRect, { x, y: 5.85, w: 3.7, h: 0.62, fill: { color: r[2] ? C.accent1 : C.accent4 }, rectRadius: 0.08, objectName: `로드맵 ${i + 1}` });
    T(s, [{ text: r[0], options: { bold: true } }, { text: '  (' + r[1] + ')', options: {} }], { x: x + 0.1, y: 5.85, w: 3.5, h: 0.62, fontSize: 14, color: C.background1, valign: 'middle', align: 'center', objectName: `로드맵 문구 ${i + 1}` });
    if (i < 2) s.addShape(pres.ShapeType.rightArrow, { x: x + 3.75, y: 6.0, w: 0.4, h: 0.32, fill: { color: C.text1 }, objectName: `로드맵 화살표 ${i + 1}` });
  });
  const lk = (t, u) => ({ text: t, options: { color: '0563C1', underline: { style: 'sng' }, hyperlink: { url: u } } });
  T(s, [
    { text: '앱 접속  A ', options: { color: C.accent4 } }, lk('power-redundancy-calculator.vercel.app', 'https://power-redundancy-calculator.vercel.app/'),
    { text: '  B ', options: { color: C.accent4 } }, lk('datahall-section-check.vercel.app', 'https://datahall-section-check.vercel.app/'),
    { text: '  C ', options: { color: C.accent4 } }, lk('generator-calc.vercel.app', 'https://generator-calc.vercel.app/'),
  ], { x: 3.3, y: 6.6, w: 9.4, h: 0.3, fontSize: 11, align: 'right', objectName: '앱 링크' });
  T(s, '판정: ● 충족  ◐ 부분 충족  ○ 미흡', { x: 0.6, y: 6.6, w: 12.1, h: 0.35, fontSize: 14, color: C.accent4, objectName: '판정 기준' });
  s.addNotes('연구 질문은 세 가지입니다. 첫째, 앱은 지식을 어떻게 반영하고 명시화, 검증, 공유하는가. 둘째, 기존 검토 방식에 비해 무엇을 개선하는가. 셋째, 앱들을 통합하려면 어떤 프레임워크가 필요한가입니다. 방법은 문헌 조사와 사례분석입니다. 지식 창조 이론, 설계 지식 구조화 연구, 설계과학연구 방법론에서 지식 명시화, 도구 구현, 검증, 공유·개선의 네 관점, 열두 요소로 된 분석틀을 도출했습니다. 사례는 제가 개발해 동료 실무자, 협력업체와 제한적으로 시범 적용 중인 전력 이중화, 데이터홀 단면, 발전기실 급·배기 앱입니다. 협력업체와 협의한 검토 방법과 장비 사양을 로직으로 옮겼고, Claude Code로 수정해 GitHub에 커밋하고 Vercel로 배포하는 파이프라인으로 피드백을 받아 개선합니다. 이번 과제는 세 단계 중 첫 단계입니다. (약 1분)');
}

// ================= Slide 4: 결과와 제언 =================
{
  const s = pres.addSlide({ masterName: 'CONTENT' });
  s.addText('분석 결과: 연구 질문에 대한 답', { placeholder: 'title' });
  s.addChart(pres.charts.BAR, [
    { name: '충족', labels: ['사례 A', '사례 B', '사례 C'], values: [7, 6, 8] },
    { name: '부분 충족', labels: ['사례 A', '사례 B', '사례 C'], values: [4, 4, 2] },
    { name: '미흡', labels: ['사례 A', '사례 B', '사례 C'], values: [0, 1, 1] },
  ], {
    x: 0.6, y: 1.4, w: 5.6, h: 4.1, barDir: 'bar', barGrouping: 'stacked', chartColors: ['0F766E', 'F2A93B', 'B23A48'],
    showTitle: true, title: '판정 결과 (11개 요소)', titleFontSize: 15, titleColor: '1B2733', titleFontFace: '+mn-lt',
    showValue: true, dataLabelFormatCode: '0;;;', dataLabelPosition: 'ctr', dataLabelColor: 'FFFFFF', dataLabelFontSize: 14, dataLabelFontFace: '+mn-lt',
    showLegend: true, legendPos: 'b', legendFontSize: 13, legendColor: '1B2733', legendFontFace: '+mn-lt',
    catAxisLabelColor: '1B2733', catAxisLabelFontSize: 14, catAxisLabelFontFace: '+mn-lt', catAxisOrientation: 'maxMin',
    valAxisHidden: true, valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
  });
  const F = [['RQ1  지식의 명시화·검증', '산식·결과는 화면에 명시, 근거 표시와 기준 사례 시험은 사례별 차이'], ['RQ2  기존 방식 대비 개선', '유추·직관 → 정량적이고 일관된 규칙, 검토 방식의 체계화'], ['RQ3  통합 프레임워크', '온톨로지·AI 에이전트가 앱을 제어, 결론과 근거·산식을 함께 제시']];
  F.forEach((f, i) => {
    const y = 1.5 + i * 1.35;
    s.addShape(pres.ShapeType.ellipse, { x: 6.7, y: y + 0.12, w: 0.7, h: 0.7, fill: { color: C.accent1 }, objectName: `발견 ${i + 1} 원` });
    T(s, String(i + 1), { x: 6.7, y: y + 0.12, w: 0.7, h: 0.7, fontSize: 20, bold: true, color: C.background1, align: 'center', valign: 'middle', objectName: `발견 ${i + 1} 번호` });
    T(s, f[0], { x: 7.65, y, w: 5.05, h: 0.45, fontSize: 18, bold: true, color: C.text1, valign: 'middle', objectName: `발견 ${i + 1} 제목` });
    T(s, f[1], { x: 7.65, y: y + 0.47, w: 5.05, h: 0.6, fontSize: 14, color: C.accent4, valign: 'top', objectName: `발견 ${i + 1} 설명` });
  });
  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 5.75, w: 12.1, h: 1.2, fill: { color: C.text1 }, rectRadius: 0.1, objectName: '제언 상자' });
  T(s, [{ text: '제언  ', options: { bold: true, color: C.accent2 } }, { text: '지식 원본 단일화 · 근거 수준 표기 · 기준 사례 시험 · 공통 포털과 의견 이력', options: { color: C.background1 } }], { x: 0.9, y: 5.85, w: 11.5, h: 0.5, fontSize: 17, valign: 'middle', objectName: '제언 문장' });
  T(s, [{ text: '기대 효과  ', options: { bold: true, color: C.accent2 } }, { text: '공간 정합성↑ · 설계 변경 리스크↓ · 산업·방송통신시설 등 다른 특수 시설로 확장', options: { color: C.background1 } }], { x: 0.9, y: 6.4, w: 11.5, h: 0.45, fontSize: 17, valign: 'middle', objectName: '기대 효과 문장' });
  s.addNotes('결과입니다. 연구 질문 1에 대해, 세 앱 모두 산식과 결과를 화면에 드러내고 도면으로 연결하며 스스로 경고합니다. 다만 근거 표시와 기준 사례 시험은 사례마다 수준이 달랐고, 같은 냉각 산식이 두 앱에 중복되어 있었습니다. 연구 질문 2에 대해, 앱은 사례 유추와 직관, 업체의 비공식 도움에 의존하던 검토를 정량적이고 일관된 규칙으로 바꾸어 검토 방식을 체계화했습니다. 연구 질문 3에 대해, 앱 사이의 연결과 검토 순서는 아직 설계자에게 남아 있어, 온톨로지와 AI 에이전트가 앱을 직접 제어하고 결론과 근거, 산식을 함께 제시하는 프레임워크를 제안합니다. 이를 통해 초기 계획설계의 정합성을 높이고 설계 변경 리스크를 줄일 것으로 기대하며, 분석틀은 다른 특수 시설로도 확장할 수 있습니다. 다만 정확도 향상의 효과는 후속 연구에서 검증합니다. (약 1분)');
}

(async () => {
  await pres.writeFile({ fileName: 'presentation.pptx' });
  await applyTheme('presentation.pptx', THEME);
  console.log('written');
})();
