const { chromium } = require('playwright');
(async()=>{const b=await chromium.launch();const p=await b.newPage({deviceScaleFactor:2,viewport:{width:1140,height:500}});
await p.goto('file://'+__dirname+'/agent.html');await p.waitForTimeout(300);await p.locator('.wrap').screenshot({path:__dirname+'/../fig_5_1_agent.png'});await b.close()})()
