const {chromium}=require('playwright');
(async()=>{const b=await chromium.launch();const p=await b.newPage({viewport:{width:1920,height:1080}});
await p.goto('file://'+process.cwd()+'/src/propuesta.html');await p.waitForLoadState('networkidle');
await p.pdf({path:'entregables/Propuesta-Mozart-Movitecnica.pdf',width:'1920px',height:'1080px',printBackground:true});
const n=await p.$$eval('section',s=>s.length);
for(let i=0;i<n;i++){const el=(await p.$$('section'))[i];await el.screenshot({path:process.env.SP+`/slide-${String(i+1).padStart(2,'0')}.png`});}
console.log('slides',n);await b.close();})();
