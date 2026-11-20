// 사례 C 캡처: CDN(unpkg·cdnjs) 요청을 npm 레지스트리에서 받은 같은 라이브러리 파일로 대체
const { chromium } = require('playwright');
const L = '/tmp/claude-0/-home-user/b5e6f573-7b81-5e78-b347-81f5624c7faa/scratchpad/cdnlibs';
const MAP = [[/react-dom@18.*react-dom\.development\.js/, L + '/react-dom-18.3.1/package/umd/react-dom.development.js'],
             [/react@18.*react\.development\.js/, L + '/react-18.3.1/package/umd/react.development.js'],
             [/@babel\/standalone/, L + '/babel-standalone-8.1.2/package/babel.min.js'],
             [/xlsx\.full\.min\.js/, L + '/xlsx-0.18.5/package/dist/xlsx.full.min.js']];
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1440, height: 1000 }, deviceScaleFactor: 2 });
  await p.route(/^https:\/\/(unpkg\.com|cdnjs\.cloudflare\.com)\//, r => {
    const m = MAP.find(([re]) => re.test(r.request().url()));
    return m ? r.fulfill({ path: m[1], contentType: 'application/javascript' }) : r.abort();
  });
  const errs = []; p.on('pageerror', e => errs.push(e.message));
  await p.goto('http://127.0.0.1:8765/index.html', { waitUntil: 'networkidle' });
  await p.waitForTimeout(4000);
  await p.screenshot({ path: process.argv[2] + '/fig_caseC_main.png', fullPage: false });
  const h = await p.evaluate(() => document.body.scrollHeight);
  await p.screenshot({ path: process.argv[2] + '/fig_caseC_full.png', fullPage: true });
  console.log('height', h, 'errors', errs.slice(0, 3)); await b.close();
})();
