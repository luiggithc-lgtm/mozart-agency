const {chromium}=require('playwright');
(async()=>{const b=await chromium.launch();
for(const [src,out,w,h] of [['src/mockup-home.html','entregables/muestra-home-movitecnica.png',1600,1000],['src/storyboard-reel.html','entregables/muestra-storyboard-reel.png',1920,1080]]){
 const p=await b.newPage({viewport:{width:w,height:h},deviceScaleFactor:2});
 await p.goto('file://'+process.cwd()+'/'+src);await p.waitForTimeout(300);
 await p.screenshot({path:out});await p.close();}
await b.close();})();
