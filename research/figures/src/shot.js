const { chromium } = require('playwright');
(async()=>{const b=await chromium.launch();const p=await b.newPage({deviceScaleFactor:2,viewport:{width:1140,height:600}});
for (const [f,o] of [['framework.html','fig_3_1_framework.png'],['wireframe.html','fig_4_3_wireframe.png']]){
 await p.goto('file://'+__dirname+'/'+f); await p.waitForTimeout(300); await p.locator('.wrap').screenshot({path:__dirname+'/../'+o}); }
await b.close()})()
