const {chromium}=require('playwright-core');
const fs=require('node:fs');
const path=require('node:path');
const {pathToFileURL}=require('node:url');
(async()=>{
 const root=path.resolve(__dirname,'..');
 const browser=await chromium.launch({channel:'msedge',headless:true});
 try {
  const page=await browser.newPage({viewport:{width:1400,height:1050}});
  await page.goto(pathToFileURL(path.join(root,'index.html')).href);
  await page.addScriptTag({path:path.join(root,'node_modules/pdfjs-dist/build/pdf.min.js')});
  const b64=fs.readFileSync(path.join(root,'artifacts/estoque-teste.pdf')).toString('base64');
  await page.evaluate(async ({b64,worker})=>{
   pdfjsLib.GlobalWorkerOptions.workerSrc=worker;
   const pdf=await pdfjsLib.getDocument({data:Uint8Array.from(atob(b64),c=>c.charCodeAt(0))}).promise;
   const first=await pdf.getPage(1);const viewport=first.getViewport({scale:1.6});
   document.body.replaceChildren();document.body.style.cssText='margin:20px;background:#e8e8e8';
   const canvas=document.createElement('canvas');canvas.width=viewport.width;canvas.height=viewport.height;canvas.style.background='#fff';document.body.append(canvas);
   await first.render({canvasContext:canvas.getContext('2d'),viewport}).promise;
  },{b64,worker:pathToFileURL(path.join(root,'node_modules/pdfjs-dist/build/pdf.worker.min.js')).href});
  await page.screenshot({path:path.join(root,'artifacts/pdf-preview.png'),fullPage:true});
 } finally {await browser.close()}
})().catch(e=>{console.error(e);process.exitCode=1});
