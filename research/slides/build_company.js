// 발표 자료 v2 (장수 제한 없음). 실행: node build_deck2.js
const pptxgen = require('pptxgenjs');
const SK = '/root/.claude/skills/synced/b09188c3-4889-4bd8-b551-d0a19e63b56f_facbcb1f-1d3a-4bc4-bdfb-98f8c08ef25b/pptx';
const { applyTheme } = require(SK + '/scripts/apply_theme.js');
const FIG = '/home/user/pango-gmp/research/figures/';
const THEME = { name: 'DC Study', headFontFace: 'Malgun Gothic', bodyFontFace: 'Malgun Gothic',
  colors: { dk1: '222222', lt1: 'FFFFFF', dk2: '5E5A55', lt2: 'EFEFEF', accent1: '8E8A84', accent2: 'DAD6D0', accent3: '3A3836', accent4: '8C8C8C', accent5: '5E5A55', accent6: 'C8C5C0', hlink: '3A3836', folHlink: '8C8C8C' } };
const pres = new pptxgen(); pres.layout = 'LAYOUT_WIDE';
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = '건축 초기 계획설계의 지식 확장을 위한 Micro-App 분석'; pres.author = '임상철';
const C = pres.SchemeColor;
pres.defineSlideMaster({ title: 'COVER', background: { color: 'FFFFFF' }, objects: [] });
pres.defineSlideMaster({ title: 'CONTENT', background: { color: 'FFFFFF' }, objects: [
  { placeholder: { options: { name: 'title', type: 'title', x: 1.5, y: 0.4, w: 11.2, h: 0.75, fontSize: 30, bold: true, color: '222222', valign: 'middle', align: 'left', margin: 0 }, text: '' } },
  { line: { x: 0.6, y: 1.3, w: 12.1, h: 0, line: { color: 'D4D4D4', width: 1 } } },
  { text: { text: '• • •', options: { x: 12.0, y: 0.3, w: 0.7, h: 0.3, fontSize: 12, color: 'C8C8C8', align: 'right', margin: 0 } } },
  { text: { text: 'AI 컨퍼런스 2026 | AI 활용사례', options: { x: 0.6, y: 7.05, w: 6, h: 0.3, fontSize: 10, color: '8C8C8C', margin: 0 } } }],
  slideNumber: { x: 12.2, y: 7.05, w: 0.5, h: 0.3, fontSize: 10, color: '8C8C8C', align: 'right' } });
const T = (s, text, o) => s.addText(text, Object.assign({ isTextBox: true, margin: 0 }, o));
const SUB = {};
const tw = (t, fs) => [...t].reduce((a, ch) => a + (/[\u3131-\uD79D]/.test(ch) ? 1.0 : (ch === ' ' ? 0.33 : 0.55)), 0) * fs / 72 * 1.02;
const bubble = (s, x, y, w = 0.7, h = 0.5) => {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: '8E8A84' }, rectRadius: 0.08 });
  s.addShape(pres.ShapeType.rtTriangle, { x: x + 0.12, y: y + h - 0.02, w: 0.16, h: 0.14, fill: { color: '8E8A84' }, flipV: true });
  T(s, '•••', { x, y, w, h, fontSize: 13, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' }); };
const content = (title, notes) => { const s = pres.addSlide({ masterName: 'CONTENT' }); s.addText(title, { placeholder: 'title' });
  bubble(s, 0.6, 0.53);
  const sub = SUB[title]; if (sub) T(s, sub, { x: 1.5 + Math.min(tw(title, 30), 9.5) + 0.2, y: 0.62, w: 3, h: 0.4, fontSize: 12, color: '8C8C8C', valign: 'bottom' });
  if (notes) s.addNotes(notes); return s; };
const card = (s, x, y, w, h, head, body, o = {}) => {
  s.addShape(pres.ShapeType.roundRect, { x, y, w, h, fill: { color: o.fill || C.background2 }, rectRadius: 0.1 });
  T(s, head, { x: x + 0.2, y: y + 0.12, w: w - 0.4, h: o.hh || 0.5, fontSize: o.hs || 18, bold: true, color: o.hc || C.text1, valign: 'middle', align: o.al || 'center' });
  if (body) T(s, body, { x: x + 0.2, y: y + (o.hh || 0.5) + 0.15, w: w - 0.4, h: h - (o.hh || 0.5) - 0.25, fontSize: o.bs || 15, color: o.bc || '555555', valign: 'top', align: o.al || 'center' });
};
const img = (s, f, x, y, w, h) => s.addImage({ path: FIG + f, x, y, w, h, sizing: { type: 'contain', w, h } });
const src = (s, t, y = 6.7) => T(s, t, { x: 0.6, y, w: 12.1, h: 0.28, fontSize: 11, color: C.accent4 });
const bullets = (s, items, x, y, w, h, fs = 16) => s.addText(items.map((t, i) => ({ text: t, options: { bullet: { indent: 18 }, breakLine: i < items.length - 1, paraSpaceAfter: 8 } })), { x, y, w, h, fontSize: fs, color: C.text1, valign: 'top', margin: 0 });

const PH = (t) => ({ text: t, options: { color: 'B23A48', bold: true } }); // 직접 채울 값
// 1 표지
{ const s = pres.addSlide({ masterName: 'COVER' });
  bubble(s, 6.27, 0.9, 0.8, 0.58);
  T(s, '설계자가 직접 만든 AI 검토 앱', { x: 0.6, y: 1.75, w: 12.1, h: 0.95, fontSize: 44, bold: true, color: '9A968F', align: 'center', valign: 'middle' });
  T(s, '데이터센터 초기 설계, 업체 회신 전에 답을 얻다', { x: 0.6, y: 2.65, w: 12.1, h: 0.95, fontSize: 40, bold: true, color: '222222', align: 'center', valign: 'middle' });
  T(s, '바이브 코딩으로 만든 전력 · 데이터홀 · 발전기실 검토 앱 3종', { x: 0.6, y: 3.85, w: 12.1, h: 0.45, fontSize: 19, color: '333333', align: 'center' });
  s.addShape(pres.ShapeType.line, { x: 0.6, y: 5.75, w: 12.1, h: 0, line: { color: 'C8C8C8', width: 1 } });
  T(s, [{ text: '[부서명]', options: { color: 'B23A48', bold: true } }, { text: '   임상철', options: { bold: true } }, { text: '     |     ', options: { color: 'BBBBBB' } }, { text: '트랙 1 · AI 활용사례' }, { text: '     |     ', options: { color: 'BBBBBB' } }, { text: '2026.12.18' }], { x: 0.6, y: 6.0, w: 12.1, h: 0.4, fontSize: 13, color: '333333', align: 'center' });
  s.addNotes('안녕하세요. 데이터센터 초기 설계에서 협력업체 회신을 기다리던 검토를 설계자가 직접 만든 AI 앱으로 바로 확인하게 된 사례를 소개하겠습니다.'); }

// 2 문제
{ const s = content('이런 경험, 있으셨죠?', '데이터센터 초기 설계에서는 전기실, 발전기실, 데이터홀 규모가 건물 전체를 좌우합니다. 그런데 이 단계에는 아직 전기·설비 업체와 계약이 안 되어 있어, 비슷한 사례로 추정하거나 업체에 비공식으로 부탁해야 했습니다. 나중에 정식 검토 결과가 다르면 규모와 계획을 다시 해야 했습니다.');
  const st = [['계약 전 초기 설계', '전기·설비 업체가\n아직 참여하지 않음'], ['경험과 사례로 추정', '비공식 문의 후\n회신을 기다림'], ['정식 검토와 차이', '전기실·발전기실·층고\n규모가 달라짐'], ['설계 변경·재작업', '일정과 비용 부담']];
  st.forEach((b, i) => { const x = 0.6 + i * 3.1;
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.7, w: 2.75, h: 2.3, fill: { color: i === 3 ? '3A3836' : 'EFEFEF' }, rectRadius: 0.1 });
    T(s, b[0], { x: x + 0.1, y: 1.85, w: 2.55, h: 0.6, fontSize: 18, bold: true, color: i === 3 ? 'FFFFFF' : '222222', align: 'center' });
    T(s, b[1], { x: x + 0.1, y: 2.55, w: 2.55, h: 1.2, fontSize: 15, color: i === 3 ? 'DDDDDD' : '555555', align: 'center' });
    if (i < 3) s.addShape(pres.ShapeType.rightArrow, { x: x + 2.78, y: 2.65, w: 0.3, h: 0.36, fill: { color: '8C8C8C' } }); });
  T(s, '︾', { x: 0.6, y: 4.3, w: 12.1, h: 0.4, fontSize: 20, color: '9A9A9A', align: 'center' });
  s.addShape(pres.ShapeType.rect, { x: 0.6, y: 4.8, w: 12.1, h: 1.1, fill: { color: 'FFFFFF' }, line: { color: '8C8C8C', width: 1 } });
  T(s, '업체 검토 전에, 설계자가 근거 있는 수치를 바로 확인할 수 없을까?', { x: 0.6, y: 4.8, w: 12.1, h: 1.1, fontSize: 22, bold: true, color: '222222', align: 'center', valign: 'middle' }); }

// 3 해결
{ const s = content('해결: 설계자가 AI로 검토 앱을 직접 만들다', '그래서 협력업체와 협의하며 메모해 둔 검토 방법과 장비 사양을 모아, Claude Code라는 AI 코딩 도구로 검토 앱을 직접 만들었습니다. 코딩 전문가가 아니어도 말로 요구사항을 설명하면 AI가 코드를 만들어 주는 방식, 이른바 바이브 코딩입니다.');
  const c = [['흩어진 실무 지식', '협력업체 협의 메모\n장비 사양 · 계산 근거\n전문가의 검토 순서'], ['AI로 앱 제작', '말로 요구사항 설명\n→ Claude Code가 코드 작성\n(바이브 코딩)'], ['웹에서 바로 사용', '입력하면 즉시 계산\n결과를 도면·다이어그램으로\n링크만 있으면 누구나']];
  c.forEach((b, i) => { const x = 0.6 + i * 4.1;
    card(s, x, 1.6, 3.9, 3.1, b[0], b[1], { hs: 21, bs: 16, hh: 0.6 });
    if (i < 2) s.addShape(pres.ShapeType.rightArrow, { x: x + 3.92, y: 2.95, w: 0.17, h: 0.35, fill: { color: '8C8C8C' } }); });
  s.addShape(pres.ShapeType.rect, { x: 0.6, y: 5.1, w: 12.1, h: 0.95, fill: { color: 'FFFFFF' }, line: { color: '8C8C8C', width: 1 } });
  T(s, '코딩 전문가가 아닌 설계 실무자가, 자기 업무 지식으로 만든 도구', { x: 0.6, y: 5.1, w: 12.1, h: 0.95, fontSize: 20, bold: true, color: '222222', align: 'center', valign: 'middle' }); }

// 4 앱 3종
{ const s = content('만든 앱 3종: 초기 설계의 핵심 검토', '앱은 세 가지입니다. 전력 이중화 앱은 전기실 규모와 개수를, 데이터홀 단면 앱은 랙과 냉각 장비 배치로 층고와 건물 규모를, 발전기실 급·배기 앱은 그릴창 면적과 층고를 확인합니다. 모두 건물 규모를 좌우하는 검토입니다.');
  const lk = (t, u) => ({ text: t, options: { color: '3A3836', bold: true, underline: { style: 'sng' }, hyperlink: { url: u } } });
  const ap = [['A', '전력 이중화 검토', '이중화 방식·변압기 용량 입력\n→ 전기실 규모·개수, 급전 구조', '평면', 'https://power-redundancy-calculator.vercel.app/'], ['B', '데이터홀 단면 검토', '랙·냉각 장비·기둥 간격 입력\n→ 층고, 보 아래 공간, 장비 배치', '단면·규모', 'https://datahall-section-check.vercel.app/'], ['C', '발전기실 급·배기 계산', '발전기 모델·대수 입력\n→ 급·배기 그릴 면적, 풍속', '입면·층고', 'https://generator-calc.vercel.app/']];
  ap.forEach((c, i) => { const x = 0.6 + i * 4.1;
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.55, w: 3.9, h: 4.4, fill: { color: 'EFEFEF' }, rectRadius: 0.1 });
    s.addShape(pres.ShapeType.ellipse, { x: x + 1.55, y: 1.75, w: 0.8, h: 0.8, fill: { color: '8E8A84' } });
    T(s, c[0], { x: x + 1.55, y: 1.75, w: 0.8, h: 0.8, fontSize: 26, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
    T(s, c[1], { x: x + 0.2, y: 2.7, w: 3.5, h: 0.5, fontSize: 20, bold: true, color: '222222', align: 'center' });
    T(s, c[2], { x: x + 0.2, y: 3.3, w: 3.5, h: 1.1, fontSize: 15, color: '555555', align: 'center' });
    T(s, '공간 영향: ' + c[3], { x: x + 0.2, y: 4.5, w: 3.5, h: 0.4, fontSize: 15, bold: true, color: '5E5A55', align: 'center' });
    T(s, [lk('앱 열기 ↗', c[4])], { x: x + 0.2, y: 5.2, w: 3.5, h: 0.4, fontSize: 15, align: 'center' }); }); }

// 5~7 시연
const demo = (title, notes, f, ar, pts) => { const s = content(title, notes);
  const h = 5.0, w = Math.min(7.6, h * ar); img(s, f, 0.6, 1.45, w, h);
  const x = 0.6 + w + 0.35, W = 12.7 - x;
  pts.forEach((p, i) => { const y = 1.5 + i * 1.25;
    s.addShape(pres.ShapeType.ellipse, { x, y: y + 0.05, w: 0.5, h: 0.5, fill: { color: '8E8A84' } });
    T(s, String(i + 1), { x, y: y + 0.05, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
    T(s, p[0], { x: x + 0.65, y, w: W - 0.65, h: 0.45, fontSize: 17, bold: true, color: '222222' });
    T(s, p[1], { x: x + 0.65, y: y + 0.45, w: W - 0.65, h: 0.7, fontSize: 13.5, color: '555555', valign: 'top' }); });
  return s; };
demo('시연 A: 전력 이중화 검토', '이중화 방식과 변압기 용량을 넣으면 필요한 변압기와 전기실 구성이 바로 나오고, 급전 구조가 다이어그램으로 그려집니다. 산식도 화면에 그대로 보여 줘서 왜 그 값이 나왔는지 설명할 수 있습니다.', 'fig_app1_b_diagram_crop.png', 1.27,
  [['입력', '이중화 방식 7종, IT 용량, 변압기 용량'], ['즉시 계산', '변압기 대수, 부하율, 전기실 구성'], ['도면으로 표현', '급전 다이어그램, 기준층 평면, 데이터홀'], ['근거 표시', '4,000 × 0.99 × 0.9 × 0.96 = 3,421.4kW']]);
demo('시연 B: 데이터홀 단면 검토', '랙 배치와 냉각 장비를 넣으면 필요한 풍량과 보 아래 환기 공간이 계산되고, 입면·단면·평면이 자동으로 그려집니다. 풍속이 넘치거나 기둥과 간섭하면 바로 경고합니다.', 'fig_caseB_crop.png', 1.557,
  [['입력', '서버 수량, 랙 배치, 기둥 간격, 랙 부하'], ['즉시 계산', '필요 풍량, 보 아래 환기 공간, 층고'], ['도면으로 표현', '입면·단면·평면 자동 작성, DXF 내보내기'], ['자동 경고', '풍속 초과, 기둥 간섭']]);
demo('시연 C: 발전기실 급·배기 계산', '제조사와 모델을 고르면 사양이 자동으로 들어가고, 대수에 따른 급기·배기 그릴 면적과 풍속이 계산됩니다. 카탈로그 출처를 화면에 함께 보여 줍니다.', 'fig_caseC_crop.png', 1.485,
  [['입력', '발전기 제조사·모델 선택, 대수'], ['즉시 계산', '급기·배기 개구 면적, 풍속'], ['도면으로 표현', '외기 그릴 → 소음기 → 댐퍼 계통도'], ['근거 표시', '제조사 카탈로그 출처 쪽수 표시']]);

// 8 효과
{ const s = content('달라진 점', '기존에는 비슷한 사례로 추정하거나 업체 회신을 기다렸다면, 이제는 조건을 넣으면 바로 결과와 근거를 확인합니다. 같은 입력이면 누가 해도 같은 결과가 나오고, 조건이 바뀌어도 즉시 다시 계산할 수 있습니다. 숫자는 실제 경험 기준으로 채워 주세요.');
  card(s, 0.6, 1.55, 5.6, 2.6, '기존', '유사 사례 추정 · 직관\n업체 비공식 문의 → 회신 대기\n담당자마다 다른 결과', { hs: 22, bs: 17, hh: 0.6, fill: 'EFEFEF' });
  s.addShape(pres.ShapeType.rightArrow, { x: 6.3, y: 2.6, w: 0.6, h: 0.5, fill: { color: '8E8A84' } });
  card(s, 7.0, 1.55, 5.7, 2.6, '앱 도입 후', '조건 입력 → 즉시 결과와 산식\n같은 입력, 같은 결과 (일관성)\n조건 변경 시 바로 재계산', { hs: 22, bs: 17, hh: 0.6, fill: 'F7F7F7' });
  const m = [['검토 회신', '[ ○일 ]  →  즉시'], ['대안 비교', '[ ○개 ]  →  [ ○개 ]'], ['사용 인원', '[ ○명 ] (팀원·협력업체)']];
  m.forEach((c, i) => { const x = 0.6 + i * 4.1;
    s.addShape(pres.ShapeType.rect, { x, y: 4.5, w: 3.9, h: 1.5, fill: { color: 'FFFFFF' }, line: { color: '8C8C8C', width: 1 } });
    T(s, c[0], { x, y: 4.6, w: 3.9, h: 0.45, fontSize: 16, color: '5E5A55', align: 'center' });
    T(s, [PH(c[1])], { x, y: 5.05, w: 3.9, h: 0.8, fontSize: 24, align: 'center', valign: 'middle' }); });
  T(s, '※ 붉은 [ ] 값은 실제 경험 기준으로 채워 주세요 (측정값이 아니면 "체감 기준" 표기)', { x: 0.6, y: 6.2, w: 12.1, h: 0.35, fontSize: 12, color: '8C8C8C' }); }

// 9 만드는 방법
{ const s = content('누구나 따라 할 수 있는 제작 과정', '제작 과정은 다섯 단계입니다. 업무 지식을 정리하고, Claude Code에 말로 설명해 앱을 만들고, GitHub에 저장해 Vercel로 웹에 올립니다. 쓰면서 받은 의견을 다시 AI로 반영하는 순환 구조입니다.');
  const p = [['지식 정리', '협의 메모·사양·\n계산 근거'], ['AI로 제작', 'Claude Code\n바이브 코딩'], ['저장', 'GitHub\n버전 관리'], ['배포', 'Vercel\n웹 링크'], ['사용·개선', '팀원·협력업체\n의견 반영']];
  p.forEach((t, i) => { const x = 0.6 + i * 2.47;
    s.addShape(pres.ShapeType.ellipse, { x: x + 0.55, y: 1.7, w: 1.1, h: 1.1, fill: { color: i === 4 ? '3A3836' : '8E8A84' } });
    T(s, String(i + 1), { x: x + 0.55, y: 1.7, w: 1.1, h: 1.1, fontSize: 26, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
    T(s, t[0], { x, y: 3.0, w: 2.2, h: 0.5, fontSize: 18, bold: true, color: '222222', align: 'center' });
    T(s, t[1], { x, y: 3.5, w: 2.2, h: 0.9, fontSize: 14, color: '555555', align: 'center' });
    if (i < 4) s.addShape(pres.ShapeType.rightArrow, { x: x + 1.85, y: 2.1, w: 0.5, h: 0.3, fill: { color: 'C8C5C0' } }); });
  T(s, '↺  사용하며 받은 의견을 다시 AI로 반영하는 순환', { x: 0.6, y: 4.6, w: 12.1, h: 0.45, fontSize: 17, bold: true, color: '5E5A55', align: 'center' });
  s.addShape(pres.ShapeType.rect, { x: 0.6, y: 5.3, w: 12.1, h: 0.9, fill: { color: 'FFFFFF' }, line: { color: '8C8C8C', width: 1 } });
  T(s, '핵심은 코딩 실력이 아니라, 실무자의 업무 지식', { x: 0.6, y: 5.3, w: 12.1, h: 0.9, fontSize: 20, bold: true, color: '222222', align: 'center', valign: 'middle' }); }

// 10 다음 단계
{ const s = content('다음 단계: 앱을 연결하는 AI 검토 비서', '지금은 앱이 각각 따로 계산합니다. 다음 단계는 앱들을 연결해, 수전용량이나 냉각 방식이 바뀌면 전기실, 발전기실, 데이터홀에 미치는 영향을 함께 보여 주고, AI가 근거와 산식을 담은 검토의견서까지 만들어 주는 것입니다.');
  const r = [['지금', '앱 3종이 각각 계산', '시범 적용 (팀원·협력업체)', false], ['다음', '앱 통합 · 공통 지식 정리', '사내 공유 · 의견 수집', false], ['목표', 'AI 에이전트가 앱을 호출', '근거·산식이 담긴\n종합 검토의견서 자동 작성', true]];
  r.forEach((c, i) => { const x = 0.6 + i * 4.1;
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.6, w: 3.9, h: 3.4, fill: { color: c[3] ? '3A3836' : 'EFEFEF' }, rectRadius: 0.1 });
    T(s, c[0], { x: x + 0.2, y: 1.75, w: 3.5, h: 0.6, fontSize: 24, bold: true, color: c[3] ? 'DAD6D0' : '8E8A84', align: 'center' });
    T(s, c[1], { x: x + 0.2, y: 2.5, w: 3.5, h: 0.9, fontSize: 18, bold: true, color: c[3] ? 'FFFFFF' : '222222', align: 'center' });
    T(s, c[2], { x: x + 0.2, y: 3.45, w: 3.5, h: 1.2, fontSize: 15, color: c[3] ? 'DDDDDD' : '555555', align: 'center' });
    if (i < 2) s.addShape(pres.ShapeType.rightArrow, { x: x + 3.92, y: 3.1, w: 0.17, h: 0.35, fill: { color: '8C8C8C' } }); });
  s.addShape(pres.ShapeType.rect, { x: 0.6, y: 5.35, w: 12.1, h: 0.9, fill: { color: 'FFFFFF' }, line: { color: '8C8C8C', width: 1 } });
  T(s, '수전용량이 바뀌면? → 전기실 · 발전기실 · 데이터홀 영향을 한 번에', { x: 0.6, y: 5.35, w: 12.1, h: 0.9, fontSize: 19, bold: true, color: '222222', align: 'center', valign: 'middle' }); }

// 11 제안
{ const s = content('함께 해 주세요', '이 방식은 데이터센터뿐 아니라 설비 비중이 큰 다른 시설에도 적용할 수 있습니다. 사내 공유와 피드백, 그리고 각 부서의 검토 지식을 함께 모으는 협업을 제안드립니다.');
  const k = [['써 보고 의견 주기', '링크로 바로 사용\n불편한 점·오류 제보'], ['검토 지식 공유', '부서별 검토 기준·\n산식·체크리스트'], ['다른 시설로 확장', '산업시설 · 방송통신시설 등\n설비 비중이 큰 시설']];
  k.forEach((c, i) => card(s, 0.6 + i * 4.1, 1.6, 3.9, 2.9, c[0], c[1], { hs: 21, bs: 16, hh: 0.6 }));
  s.addShape(pres.ShapeType.rect, { x: 0.6, y: 4.9, w: 12.1, h: 1.1, fill: { color: 'FFFFFF' }, line: { color: '8C8C8C', width: 1 } });
  T(s, '설계자의 경험을, 모두가 쓰는 도구로', { x: 0.6, y: 4.9, w: 12.1, h: 1.1, fontSize: 24, bold: true, color: '222222', align: 'center', valign: 'middle' });
  T(s, '문의: [연락처]   ·   HALO Community 등록 사례: [링크]', { x: 0.6, y: 6.2, w: 12.1, h: 0.4, fontSize: 13, color: 'B23A48', bold: true, align: 'center' }); }

// 12 링크
{ const s = content('지금 바로 써 보세요', '세 앱 모두 링크로 바로 열 수 있습니다. 감사합니다.');
  const lk = (t, u) => ({ text: t, options: { color: '3A3836', bold: true, underline: { style: 'sng' }, hyperlink: { url: u } } });
  const ap = [['A', '전력 이중화 검토', 'https://power-redundancy-calculator.vercel.app/'], ['B', '데이터홀 단면 검토', 'https://datahall-section-check.vercel.app/'], ['C', '발전기실 급·배기 계산', 'https://generator-calc.vercel.app/']];
  ap.forEach((c, i) => { const y = 1.55 + i * 1.5;
    s.addShape(pres.ShapeType.roundRect, { x: 0.6, y, w: 12.1, h: 1.25, fill: { color: 'EFEFEF' }, rectRadius: 0.1 });
    s.addShape(pres.ShapeType.ellipse, { x: 0.9, y: y + 0.22, w: 0.8, h: 0.8, fill: { color: '8E8A84' } });
    T(s, c[0], { x: 0.9, y: y + 0.22, w: 0.8, h: 0.8, fontSize: 26, bold: true, color: 'FFFFFF', align: 'center', valign: 'middle' });
    T(s, c[1], { x: 2.0, y: y + 0.15, w: 10, h: 0.5, fontSize: 21, bold: true, color: '222222' });
    T(s, [lk(c[2], c[2])], { x: 2.0, y: y + 0.65, w: 10, h: 0.45, fontSize: 17 }); });
  T(s, '감사합니다', { x: 0.6, y: 6.15, w: 12.1, h: 0.6, fontSize: 28, bold: true, color: '222222', align: 'center' }); }

(async () => { await pres.writeFile({ fileName: 'company.pptx' }); await applyTheme('company.pptx', THEME); console.log('written'); })();
