// 앱 ① 기준 계산 사례 테스트: 실제 앱(index.html)을 브라우저로 실행해 화면 값을
// 독립적으로 작성한 산식(아래 oracle)과 비교한다. 사용법:
//   NODE_PATH=$(npm root -g) node calc_reference_test.js <앱 폴더 절대경로>
const { chromium } = require('playwright');
const ceil = x => Math.ceil(x - 1e-9);
// 독립 산식: 설명서 §4 (코드를 보지 않고 산식만으로 작성)
function oracle(loadMW, rating, pf, ups, safety) {
  const eff = rating * pf * safety * ups, L = loadMW * 1000;
  const rows = { '2oo3': [2, 1, 'd'], '3oo4': [3, 1, 'd'], 'N+1': [3, 1, 's'], 'N+2C': [3, 2, 's'], '2N': [1, 1, 'f'] };
  const out = {};
  for (const [k, [n, c, t]] of Object.entries(rows)) {
    const per = t === 'f' ? L : L / n;
    const u = Math.max(1, ceil(per / eff));
    out[k] = { perMW: per / 1000, perUnits: u, total: u * (n + c) };
  }
  return { eff, out };
}
const num = s => parseFloat(String(s).replace(/[^0-9.]/g, ''));
(async () => {
  const dir = process.argv[2];
  const b = await chromium.launch(); const p = await b.newPage();
  await p.route(/^https?:/, r => r.abort());
  await p.goto('file://' + dir + '/index.html'); await p.waitForTimeout(500);
  const set = async (id, v) => p.evaluate(([id, v]) => { const e = document.getElementById(id); e.value = v; e.dispatchEvent(new Event('input', { bubbles: true })); e.dispatchEvent(new Event('change', { bubbles: true })); }, [id, String(v)]);
  const cases = [];
  for (const rating of [2000, 2500, 3000, 4000])
    for (const safety of [0.9, 0.85, 0.8, 0.75])
      for (const pod of [3000, 6000]) cases.push({ rating, safety, pod, pf: 0.99, ups: 0.96 });
  cases.push({ rating: 4000, safety: 0.9, pod: 6000, pf: 1.0, ups: 0.98 });
  let pass = 0, fail = [];
  for (const c of cases) {
    await set('podCapacity', c.pod); await set('rating', c.rating); await set('safety', c.safety); await set('pf', c.pf); await set('upsEfficiency', c.ups);
    await p.waitForTimeout(60);
    const r = await p.evaluate(() => ({
      load: parseFloat(document.getElementById('load').value), cap: document.getElementById('stepCap').textContent,
      ref: [...document.querySelectorAll('#refTableBody tr')].map(t => [...t.cells].map(x => x.textContent.trim()))
    }));
    const o = oracle(r.load, c.rating, c.pf, c.ups, c.safety);
    const check = (name, got, exp) => { if (Math.abs(got - exp) > 0.06) fail.push({ c, name, got, exp }); else pass++; };
    check('유효용량(kW)', num(r.cap), o.eff);
    const keys = ['2oo3', '3oo4', 'N+1', 'N+2C', '2N'];
    r.ref.forEach((row, i) => { check(keys[i] + ' 모듈당 대수', num(row[3]), o.out[keys[i]].perUnits); check(keys[i] + ' 총 대수', num(row[4]), o.out[keys[i]].total); });
  }
  console.log(`cases=${cases.length} checks_passed=${pass} failed=${fail.length}`);
  console.log(JSON.stringify(fail.slice(0, 10), null, 1));
  await b.close(); process.exit(fail.length ? 1 : 0);
})();
