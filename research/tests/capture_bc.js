// 사례 B(데이터홀 단면 검토)·C(발전기실 급·배기) 화면 캡처: 저장소 소스를 로컬 서버로 실행
const { chromium } = require('playwright');
const out = process.argv[2];
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1440, height: 1000 }, deviceScaleFactor: 2 });
  const errs = []; p.on('pageerror', e => errs.push(e.message));
  for (const [url, name] of [['http://127.0.0.1:8766/index.html', 'fig_caseB_main'], ['http://127.0.0.1:8765/index.html', 'fig_caseC_main']]) {
    await p.goto(url, { waitUntil: 'networkidle', timeout: 60000 }).catch(e => errs.push(name + ' goto ' + e.message));
    await p.waitForTimeout(2500);
    await p.screenshot({ path: `${out}/${name}.png` });
    const h = await p.evaluate(() => document.body.scrollHeight);
    console.log(name, 'height', h);
  }
  console.log('errors', errs.slice(0, 5)); await b.close();
})();
