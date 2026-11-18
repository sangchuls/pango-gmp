// 앱 ①·② 배포본 화면 캡처 (논문 그림용). NODE_PATH=$(npm root -g) node capture.js <출력폴더>
const { chromium } = require('playwright');
const out = process.argv[2];
(async () => {
  const b = await chromium.launch();
  const ctx = await b.newContext({ viewport: { width: 1280, height: 900 }, deviceScaleFactor: 2, locale: 'ko-KR' });
  const p = await ctx.newPage();
  const log = [];
  // 앱 ①
  await p.goto('https://power-redundancy-calculator.vercel.app/', { waitUntil: 'networkidle' });
  await p.waitForTimeout(1500);
  const cards = p.locator('.card');
  log.push('app1 cards=' + await cards.count());
  const shot = async (loc, name) => { await loc.scrollIntoViewIfNeeded(); await p.waitForTimeout(400); await loc.screenshot({ path: `${out}/${name}.png` }); log.push(name); };
  await shot(cards.nth(3), 'fig_app1_a_result');
  await shot(cards.nth(4), 'fig_app1_b_diagram');
  await shot(cards.nth(8), 'fig_app1_c_datahall');
  await p.click('#fbOpenBtn'); await p.waitForTimeout(600);
  const modal = p.locator('#fbModal > *').first();
  await shot(modal, 'fig_app1_d_feedback');
  // 앱 ②
  const B = 'https://ai-factory-study.vercel.app/';
  for (const [hash, name] of [['', 'fig_app2_a_home'], ['#/pod/basic', 'fig_app2_b_pod'], ['#/glossary', 'fig_app2_c_glossary'], ['#/sources', 'fig_app2_d_sources']]) {
    await p.goto(B + hash, { waitUntil: 'networkidle' }); await p.waitForTimeout(1200);
    await p.evaluate(() => window.scrollTo(0, 0));
    await p.screenshot({ path: `${out}/${name}.png` }); log.push(name);
  }
  console.log(log.join('\n')); await b.close();
})();
