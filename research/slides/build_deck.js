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

// ================= Slide 1: 배경·문제의식 (일반 문제 → 데이터센터) =================
{
  const s = pres.addSlide({ masterName: 'DARK' });
  s.addText('건축 초기 계획설계의 지식 확장을 위한 Micro-App 분석', { placeholder: 'title' });
  T(s, '바이브 코딩으로 개발한 데이터센터 검토 앱 3개 사례를 중심으로', { x: 0.7, y: 2.2, w: 8.5, h: 0.5, fontSize: 18, color: C.accent6, objectName: '부제' });
  // 1단계: 일반 문제
  T(s, '1  초기 계획설계의 문제', { x: 0.7, y: 3.0, w: 6, h: 0.4, fontSize: 16, bold: true, color: C.accent2, objectName: '1단계 라벨' });
  const bx = [['협업·지원이 어려움', '사례 유추·비공식 업체 도움'], ['전문 검토 결과', '직접 판단과 달라질 수 있음'], ['공간계획 변경', '재작업 발생 → 지식 확장 필요']];
  bx.forEach((b, i) => {
    const x = 0.7 + i * 4.15;
    s.addShape(pres.ShapeType.roundRect, { x, y: 3.45, w: 3.5, h: 1.25, fill: { color: C.accent5 }, rectRadius: 0.1, objectName: `문제 상자 ${i + 1}` });
    T(s, b[0], { x: x + 0.2, y: 3.55, w: 3.1, h: 0.45, fontSize: 18, bold: true, color: C.background1, objectName: `문제 제목 ${i + 1}` });
    T(s, b[1], { x: x + 0.2, y: 4.05, w: 3.1, h: 0.55, fontSize: 15, color: C.accent6, valign: 'top', objectName: `문제 설명 ${i + 1}` });
    if (i < 2) s.addShape(pres.ShapeType.rightArrow, { x: x + 3.6, y: 3.9, w: 0.45, h: 0.38, fill: { color: C.accent4 }, objectName: `화살표 ${i + 1}` });
  });
  // 2단계: 왜 데이터센터인가
  T(s, '2  왜 데이터센터인가', { x: 0.7, y: 4.95, w: 6, h: 0.4, fontSize: 16, bold: true, color: C.accent2, objectName: '2단계 라벨' });
  s.addShape(pres.ShapeType.roundRect, { x: 0.7, y: 5.4, w: 8.2, h: 0.85, fill: { color: C.accent5 }, rectRadius: 0.1, objectName: '이유 상자' });
  T(s, '설비·전기 시설이 공간을 좌우하는 대표 시설: 조건이 수치로 표현됨', { x: 0.9, y: 5.4, w: 7.8, h: 0.85, fontSize: 16, color: C.background1, valign: 'middle', objectName: '이유 문장' });
  // 통계 카드 (수요 근거)
  s.addShape(pres.ShapeType.roundRect, { x: 9.3, y: 4.95, w: 3.4, h: 1.65, fill: { color: C.accent5 }, rectRadius: 0.12, objectName: '통계 카드' });
  T(s, '415 → 945 TWh', { x: 9.3, y: 5.0, w: 3.4, h: 0.7, fontSize: 28, bold: true, color: C.accent2, align: 'center', valign: 'middle', objectName: '통계 수치' });
  T(s, '데이터센터 전력 소비 2024→2030년 전망', { x: 9.4, y: 5.7, w: 3.2, h: 0.5, fontSize: 12, color: C.background1, align: 'center', valign: 'top', objectName: '통계 설명' });
  T(s, '자료: IEA(2025), Energy and AI', { x: 9.4, y: 6.2, w: 3.2, h: 0.3, fontSize: 10, color: C.accent6, align: 'center', objectName: '통계 출처' });
  T(s, '목적: AI로 만든 검토 앱의 지식 명시화·검증·공유 방식을 분석하고 개선 방향 제시', { x: 0.7, y: 6.4, w: 8.4, h: 0.45, fontSize: 15, bold: true, color: C.background1, objectName: '연구 목적' });
  T(s, '공학대학원 건축학전공  |  임상철 (2026160492)  |  담당교수 김영주  |  2026년 11월', { x: 0.7, y: 6.95, w: 12, h: 0.3, fontSize: 12, color: C.accent6, objectName: '인적사항' });
  s.addNotes('초기 계획설계에서는 구조, 기계, 전기 같은 건축 외 분야의 협업과 지원을 받기 어려워, 건축가가 경험으로 직접 판단해야 하는 것이 현실입니다. 이 검토들은 원래 전기·설비 업체의 검토가 필요한 항목인데, 초기에는 동일한 시스템과 사례를 유추하거나 비공식적으로 업체의 도움을 받아야 했습니다. 이 판단이 나중에 전문 검토 결과와 달라지면 공간계획이 바뀌고 재작업이 생깁니다. 그래서 설계자의 지식 확장이 필요합니다. 특히 설비·전기 시설이 공간을 좌우하는 시설의 초기 설계에서 두드러지는데, 데이터센터가 그 대표적인 시설입니다. 전력, 냉각, 비상 전원 방식이 설비실 규모를 직접 결정하고 조건이 수치로 표현됩니다. 수요도 2024년 약 415TWh에서 2030년 약 945TWh로 늘 전망입니다. 최근에는 실무자도 생성형 AI로 검토 도구를 직접 만듭니다. 이 도구가 어떤 지식을 어떤 근거로 담는지 정리되지 않았기 때문에, 제가 개발한 앱 3개를 분석틀로 분석했습니다. (약 1분)');
}

// ================= Slide 2: 선행연구 공백 =================
{
  const s = pres.addSlide({ masterName: 'CONTENT' });
  s.addText('선행연구와 연구 공백', { placeholder: 'title' });
  const H = { bold: true, color: 'FFFFFF', fill: { color: '1B2733' }, fontSize: 15, valign: 'middle' };
  const cell = (t, o) => ({ text: t, options: Object.assign({ fontSize: 14, valign: 'middle', color: '1B2733' }, o || {}) });
  const rows = [
    [{ text: '선행연구', options: H }, { text: '핵심 기여', options: H }, { text: '한계', options: H }],
    [cell('이병수 외(2016)', { bold: true }), cell('매스 규모산정 지식을 BIM으로 구조화'), cell('건축 내부 지식에 한정')],
    [cell('김혜정 외(2019)', { bold: true }), cell('AI 활용 초기 디자인 대안 생성'), cell('형태 대안 생성에 한정')],
    [cell('이용주(2020)', { bold: true }), cell('전문가 시스템 기반 친환경 설계지원'), cell('친환경 분야에 한정')],
    [cell('Chen 외(2024)', { bold: true }), cell('LLM·온톨로지 기반 BIM 규정 검토'), cell('규정 해석 중심')],
    [cell('Dong 외(2025)', { bold: true }), cell('LLM 다중 에이전트의 BIM 설계 조정'), cell('설비·공간 통합 판단 미포함')],
    [cell('본 연구', { bold: true, color: '0F766E', fill: { color: 'E1F3EF' } }), cell('실무자가 AI로 만든 검토 앱의 지식 명시화·검증·공유 분석', { bold: true, fill: { color: 'E1F3EF' } }), cell('저자 개발 앱 3개로 한정', { fill: { color: 'E1F3EF' } })],
  ];
  s.addTable(rows, { x: 0.6, y: 1.45, w: 12.1, colW: [2.7, 6.0, 3.4], rowH: 0.6, border: { type: 'solid', pt: 0.75, color: 'CBD5E1' }, margin: [0, 0.12, 0, 0.12], objectName: '선행연구 비교표' });
  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 5.85, w: 12.1, h: 0.95, fill: { color: C.background2 }, rectRadius: 0.1, objectName: '공백 상자' });
  T(s, [{ text: '연구 공백  ', options: { bold: true, color: C.accent1 } }, { text: '건축 실무자가 생성형 AI로 직접 만든 전문 분야 도구의 지식을 명시화·검증·공유 관점에서 분석한 연구는 검토한 범위에서 확인되지 않음', options: { color: C.text1 } }], { x: 0.85, y: 5.9, w: 11.6, h: 0.85, fontSize: 16, valign: 'middle', objectName: '공백 문장' });
  s.addNotes('선행연구는 설계 지식의 구조화, AI의 대안 생성, 전문가 시스템, LLM과 온톨로지를 활용한 규정 검토, 다중 에이전트의 BIM 조정까지 가능성을 보였습니다. 다만 대상이 건축 내부 지식이나 친환경 같은 특정 분야, 규정 해석, BIM 조작에 한정되어 있습니다. 건축 실무자가 생성형 AI로 직접 만든 전문 분야 도구의 지식을 명시화, 검증, 공유 관점에서 분석한 연구는 제가 검토한 범위에서는 확인되지 않았습니다. 이것이 본 연구의 공백입니다. (약 1분)');
}

// ================= Slide 3: 연구 모형 =================
{
  const s = pres.addSlide({ masterName: 'CONTENT' });
  s.addText('연구 모형: 분석틀로 사례 3개를 분석', { placeholder: 'title' });
  const cols = [
    { h: '이론적 근거', items: ['지식 창조 이론\n(Nonaka & Takeuchi, 1995)', '설계 지식의 도구화\n(이병수 외, 2016)', '설계과학연구 방법론\n(Peffers 외, 2007)'] },
    { h: '분석틀 (4관점·12요소)', items: ['지식 명시화  ①~④', '도구 구현  ⑤⑥', '검증  ⑦~⑨', '공유·개선  ⑩~⑫'] },
    { h: '사례분석 (3개)', items: ['A  전력 이중화 → 전기실·평면', 'B  데이터홀 → 냉각 장비·규모', 'C  발전기 풍량 → 입면·층고', '실무 활용 · 웹 배포 · 피드백 개선 중'] },
    { h: '결과와 제언', items: ['요소별 판정  ● ◐ ○', '사례 간 공통점·차이', '지식 도구화 개선 방향'] },
  ];
  const W = 2.6, G = 0.567;
  cols.forEach((c, i) => {
    const x = 0.6 + i * (W + G);
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.5, w: W, h: 4.1, fill: { color: C.background2 }, rectRadius: 0.1, objectName: `열 ${i + 1} 배경` });
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.5, w: W, h: 0.7, fill: { color: i === 1 ? C.accent1 : C.text1 }, rectRadius: 0.1, objectName: `열 ${i + 1} 머리` });
    T(s, c.h, { x: x + 0.1, y: 1.5, w: W - 0.2, h: 0.7, fontSize: 15, bold: true, color: C.background1, align: 'center', valign: 'middle', objectName: `열 ${i + 1} 제목` });
    const gap = 3.1 / c.items.length;
    c.items.forEach((it, k) => {
      T(s, it, { x: x + 0.2, y: 2.35 + k * (i === 0 ? 1.08 : gap), w: W - 0.4, h: i === 0 ? 1.0 : gap - 0.1, fontSize: 14, color: C.text1, valign: 'middle', bold: i === 1, objectName: `열 ${i + 1} 항목 ${k + 1}` });
    });
    if (i < 3) s.addShape(pres.ShapeType.rightArrow, { x: x + W + 0.08, y: 3.35, w: G - 0.16, h: 0.4, fill: { color: C.accent4 }, objectName: `화살표 ${i + 1}` });
  });
  const rm = [['① 앱 개발 과정·원리 분석', '본 과제', true], ['② 앱 통합', '계획', false], ['③ 온톨로지화·AI 에이전트', '계획', false]];
  rm.forEach((r, i) => {
    const x = 0.6 + i * 4.2;
    s.addShape(pres.ShapeType.roundRect, { x, y: 5.85, w: 3.7, h: 0.62, fill: { color: r[2] ? C.accent1 : C.accent4 }, rectRadius: 0.08, objectName: `로드맵 ${i + 1}` });
    T(s, [{ text: r[0], options: { bold: true } }, { text: '  (' + r[1] + ')', options: {} }], { x: x + 0.1, y: 5.85, w: 3.5, h: 0.62, fontSize: 14, color: C.background1, valign: 'middle', align: 'center', objectName: `로드맵 문구 ${i + 1}` });
    if (i < 2) s.addShape(pres.ShapeType.rightArrow, { x: x + 3.75, y: 6.0, w: 0.4, h: 0.32, fill: { color: C.text1 }, objectName: `로드맵 화살표 ${i + 1}` });
  });
  T(s, [
    { text: '앱 접속  A ', options: { color: C.accent4 } },
    { text: 'power-redundancy-calculator.vercel.app', options: { color: '0563C1', underline: { style: 'sng' }, hyperlink: { url: 'https://power-redundancy-calculator.vercel.app/' } } },
    { text: '   |   소스  ', options: { color: C.accent4 } },
    { text: 'github.com/sangchuls', options: { color: '0563C1', underline: { style: 'sng' }, hyperlink: { url: 'https://github.com/sangchuls' } } },
  ], { x: 5.2, y: 6.6, w: 7.5, h: 0.3, fontSize: 12, align: 'right', objectName: '앱 링크' });
  T(s, '판정: ● 충족  ◐ 부분 충족  ○ 미흡', { x: 0.6, y: 6.6, w: 12.1, h: 0.35, fontSize: 14, color: C.accent4, objectName: '판정 기준' });
  s.addNotes('연구는 앱 개발 과정과 원리의 분석, 앱 통합, 온톨로지화의 세 단계이고, 이번 과제는 첫 단계입니다. 방법은 문헌 조사와 사례분석입니다. 지식 창조 이론, 설계 지식 구조화 연구, 설계과학연구 방법론에서 네 가지 관점, 열두 개 세부 요소의 분석틀을 도출했습니다. 지식 명시화, 도구 구현, 검증, 공유·개선입니다. 이 틀로 제가 개발한 전력 이중화, 데이터홀 단면, 발전기실 급·배기 앱 세 개를 소스코드, 커밋 이력, 실행 화면, 독립 계산 대조로 분석해 충족, 부분 충족, 미흡으로 판정했습니다. 이 앱들의 로직은 평소 협력사와 협의하고 문의하며 메모한 파편화된 전문 지식을 통합하고 전문가의 검토 원리를 분석해서 만들었습니다. 지금은 실무 검토에 활용하고 웹으로 배포하며, 바이브 코딩으로 만든 앱을 저장소와 배포로 잇는 파이프라인으로 사용자 피드백을 받아 개선하고 있습니다. (약 1분)');
}

// ================= Slide 4: 결과와 제언 =================
{
  const s = pres.addSlide({ masterName: 'CONTENT' });
  s.addText('분석 결과와 제언', { placeholder: 'title' });
  s.addChart(pres.charts.BAR, [
    { name: '충족', labels: ['사례 A', '사례 B', '사례 C'], values: [7, 6, 7] },
    { name: '부분 충족', labels: ['사례 A', '사례 B', '사례 C'], values: [4, 4, 3] },
    { name: '미흡', labels: ['사례 A', '사례 B', '사례 C'], values: [0, 1, 1] },
  ], {
    x: 0.6, y: 1.4, w: 5.6, h: 4.1, barDir: 'bar', barGrouping: 'stacked', chartColors: ['0F766E', 'F2A93B', 'B23A48'],
    showTitle: true, title: '판정 결과 (11개 요소)', titleFontSize: 15, titleColor: '1B2733', titleFontFace: '+mn-lt',
    showValue: true, dataLabelFormatCode: '0;;;', dataLabelPosition: 'ctr', dataLabelColor: 'FFFFFF', dataLabelFontSize: 14, dataLabelFontFace: '+mn-lt',
    showLegend: true, legendPos: 'b', legendFontSize: 13, legendColor: '1B2733', legendFontFace: '+mn-lt',
    catAxisLabelColor: '1B2733', catAxisLabelFontSize: 14, catAxisLabelFontFace: '+mn-lt', catAxisOrientation: 'maxMin',
    valAxisHidden: true, valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
  });
  const F = [['기준 사례 시험', '세 사례 모두 저장소에 없음'], ['근거·가정 표시', '사례별 편차: C 충족, A 부분, B 미흡'], ['냉각 산식', '사례 A·B에 각각 중복 구현']];
  F.forEach((f, i) => {
    const y = 1.5 + i * 1.35;
    s.addShape(pres.ShapeType.ellipse, { x: 6.7, y: y + 0.12, w: 0.7, h: 0.7, fill: { color: C.accent1 }, objectName: `발견 ${i + 1} 원` });
    T(s, String(i + 1), { x: 6.7, y: y + 0.12, w: 0.7, h: 0.7, fontSize: 20, bold: true, color: C.background1, align: 'center', valign: 'middle', objectName: `발견 ${i + 1} 번호` });
    T(s, f[0], { x: 7.65, y, w: 5.05, h: 0.45, fontSize: 18, bold: true, color: C.text1, valign: 'middle', objectName: `발견 ${i + 1} 제목` });
    T(s, f[1], { x: 7.65, y: y + 0.47, w: 5.05, h: 0.5, fontSize: 15, color: C.accent4, valign: 'top', objectName: `발견 ${i + 1} 설명` });
  });
  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 5.75, w: 12.1, h: 1.2, fill: { color: C.text1 }, rectRadius: 0.1, objectName: '제언 상자' });
  T(s, [{ text: '제언  ', options: { bold: true, color: C.accent2 } }, { text: '지식 원본 단일화 · 근거 수준 표기 · 기준 사례 시험 · 공통 포털과 의견 이력', options: { color: C.background1 } }], { x: 0.9, y: 5.85, w: 11.5, h: 0.5, fontSize: 17, valign: 'middle', objectName: '제언 문장' });
  T(s, [{ text: '후속 연구  ', options: { bold: true, color: C.accent2 } }, { text: '앱 통합 → 온톨로지 프레임워크·테스트 앱, 타 용도 확장', options: { color: C.background1 } }], { x: 0.9, y: 6.4, w: 11.5, h: 0.45, fontSize: 17, valign: 'middle', objectName: '후속 연구 문장' });
  s.addNotes('결과입니다. 세 사례 모두 산식을 화면에 드러내고 결과를 도면으로 연결하며 스스로 경고합니다. 하지만 기준 사례 시험은 세 사례 모두 저장소에 없었고, 근거·가정 표시는 사례마다 달랐으며, 같은 냉각 산식이 두 앱에 중복 구현되어 있었습니다. 그래서 지식 원본의 단일화, 근거 수준 표기, 기준 사례 시험, 공통 포털과 의견 이력 관리를 제언합니다. 지식 유형으로 보면 앱은 선언적, 절차적 지식을 명시화했지만, 앱 사이의 조건과 연결 관계인 상황적 지식과 검토 순서인 전략적 지식은 설계자에게 남아 있습니다. 그래서 후속 연구로는 온톨로지와 AI 에이전트가 계산 도구의 결과를 통합해 검토의견서를 작성하는 프레임워크를 제안합니다. 후속 연구로는 이런 도메인 앱을 통합해 지식을 온톨로지로 구조화하는 프레임워크와 테스트 앱을 제안합니다. 분석틀은 시설 유형에 종속되지 않아 병원이나 연구시설 같은 다른 용도로 확장해 검증할 수 있습니다. 다만 이번 연구는 데이터센터에 한정됩니다. (약 1분)');
}

(async () => {
  await pres.writeFile({ fileName: 'presentation.pptx' });
  await applyTheme('presentation.pptx', THEME);
  console.log('written');
})();
