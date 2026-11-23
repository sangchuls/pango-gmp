const { chromium } = require('playwright');
(async()=>{const b=await chromium.launch();const p=await b.newPage({viewport:{width:1920,height:1080},deviceScaleFactor:1});
await p.goto('file://'+__dirname+'/cover.html');await p.waitForTimeout(600);await p.screenshot({path:__dirname+'/HALO_대표이미지_16x9.png'});await b.close()})()
