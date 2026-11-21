const { chromium } = require('playwright');
(async()=>{const b=await chromium.launch();const p=await b.newPage({deviceScaleFactor:2,viewport:{width:1140,height:600}});
await p.goto('file://'+__dirname+'/rooms.html');await p.waitForTimeout(500);await p.locator('.wrap').screenshot({path:__dirname+'/../fig_2_2_rooms.png'});await b.close()})()
