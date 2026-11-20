const { chromium } = require('playwright');
(async()=>{const b=await chromium.launch();const p=await b.newPage({deviceScaleFactor:2,viewport:{width:1140,height:500}});
await p.goto('file://'+__dirname+'/vars.html');await p.waitForTimeout(300);await p.locator('.wrap').screenshot({path:__dirname+'/../fig_5_2_vars.png'});await b.close()})()
