'use strict';
/* v14: exercise the production PNG-to-PDF path without installing browser deps.
   Production renderer output is also shared with the live A4 preview. */
const assert = require('node:assert/strict');
const fs = require('node:fs');
const {execFileSync} = require('node:child_process');

const src = fs.readFileSync('assets/app.js','utf8');
const css = fs.readFileSync('assets/pdf.css','utf8');
const html = fs.readFileSync('index.html','utf8');
const from=src.indexOf('  // Binary PNG avoids');
const to=src.indexOf('  function renderResults(',from);
assert(from>0 && to>from,'PDF implementation must be present');
const calls = [], instances = [], stageCalls=[];
let fontChecks=0;
const pngSignature=new Uint8Array([137,80,78,71,13,10,26,10]);
let pngBlobs=0;
let canvasFreed=0;
class PDF {
  constructor(options) {
    assert.deepEqual(options,{orientation:'landscape',unit:'mm',format:'a4',compress:true});
    instances.push(this);
    this.pages=1;
  }
  addPage(...args){assert.deepEqual(args,['a4','landscape']);this.pages++;}
  addImage(bytes,format,x,y,w,h,_alias,compression){
    assert.equal(format,'PNG');
    assert.equal(x,0);assert.equal(y,0);assert.equal(w,297);assert.equal(h,210);
    assert.equal(compression,'FAST');
    assert(bytes instanceof Uint8Array,'image bytes must not be base64 JPEG');
    assert.deepEqual([...bytes], [...pngSignature]);
    calls.push({bytes:bytes.length,format});
  }
  output(type){assert.equal(type,'arraybuffer');return new Uint8Array([1,2,3]).buffer;}
}
const document={fonts:{get ready(){fontChecks++;return Promise.resolve();}}};
const stage={replaceChildren(...children){stageCalls.push(children);}};
const tick=async()=>{};
const makePrintPage=(sheet,rows,one,total,signatures)=>({sheet,rows,one,total,signatures});
const html2canvas=async(_element,options)=>{
  assert.equal(options.backgroundColor,'#ffffff');
  assert.equal(options.scale,3);
  assert.equal(options.windowWidth,1123);
  assert.equal(options.windowHeight,794);
  assert.equal(options.useCORS,false);
  assert.equal(options.logging,false);
  let width=3368,height=2381;
  const canvas={
    toBlob(callback,mime){
      assert.equal(mime,'image/png');
      pngBlobs++;
      callback({arrayBuffer:async()=>pngSignature.buffer});
    },
    toDataURL(){throw Error('Lossy/base64 path must not run on modern browsers');},
    get width(){return width;}, set width(value){if(value===0)canvasFreed++;width=value;},
    get height(){return height;}, set height(value){height=value;}
  };
  return canvas;
};
const api=new Function('window','stage','html2canvas','tick','makePrintPage','document',
  src.slice(from,to)+'\nreturn {createPDF,encodePrintPNG};')(
  {jspdf:{jsPDF:PDF}},stage,html2canvas,tick,makePrintPage,document
);
(async()=>{
  const sheet={pages:[[{}],[{}]]};
  const output=await api.createPDF(sheet,2);
  assert(output instanceof Uint8Array);
  assert.deepEqual([...output],[1,2,3]);
  assert.equal(instances[0].pages,2);
  assert.equal(calls.length,2);
  assert.equal(pngBlobs,2);
  assert.equal(fontChecks,2);
  assert.equal(canvasFreed,2,'each page canvas is released');
  assert.deepEqual(stageCalls.at(-1),[],'the hidden print stage is cleaned');

  // Failed PNG encoding must not leave stale print DOM/canvas resources.
  let brokenCanvas={width:10,height:10,toBlob(callback){callback(null);}};
  let failed=false;
  try{await api.encodePrintPNG(brokenCanvas);}catch(e){
    failed=true;assert.match(e.message,/kodlanamadı/);
  }
  assert(failed);

  // Workbook palette and the shared PDF/preview CSS must match exactly.
  const py = [
    'import json,adaptive_layout as a',
    'print(json.dumps({k:getattr(a,k) for k in '+
      '["BURGUNDY","PALE","GRAY","EXAM_GRAY","WHITE","TEXT","LIGHT_LINE","CLOSED_TEXT","CLOSED_TEXT_ON_BURGUNDY"]}))'
  ].join('\n');
  const palette=JSON.parse(execFileSync('python',['-c',py],{encoding:'utf8'}));
  const vars={
    BURGUNDY:'--excel-burgundy',
    PALE:'--excel-pale',
    GRAY:'--excel-gray',
    EXAM_GRAY:'--excel-exam-gray',
    WHITE:'--excel-white',
    TEXT:'--excel-text',
    LIGHT_LINE:'--excel-line',
    CLOSED_TEXT:'--excel-closed',
    CLOSED_TEXT_ON_BURGUNDY:'--excel-closed-on-burgundy'
  };
  for(const [key,cssName] of Object.entries(vars)){
    assert(css.includes(cssName+':#'+palette[key]),key+' color mismatch');
  }
  assert.match(css,/\.print-page \.print-head\s*\{\s*background:var\(--excel-burgundy\)/);
  assert.match(css,/\.print-page \.week-top th\s*\{\s*background:var\(--excel-burgundy\)/);
  assert.match(css,/\.print-page \.student-row:nth-child\(even\) td:not\(\.exam-cell\)\{background:var\(--excel-gray\)\}/);
  assert.match(css,/\.print-page td\.num,\.print-page td\.rank\{font-weight:400\}/);
  assert.match(css,/print-color-adjust:exact/);
  assert.match(html,/assets\/pdf\.css\?v=14/);
  assert.match(html,/assets\/app\.js\?v=14/);
  assert.doesNotMatch(src.slice(from,to),/image\/jpeg|toDataURL\('image\/jpeg'/i);
  console.log('v14 high-resolution PNG PDF path, canvas cleanup, Excel color parity: OK');
})().catch(error=>{console.error(error);process.exitCode=1;});
