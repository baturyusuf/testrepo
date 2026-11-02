/* Compare the actual production v11 and v12 HTML page output at startWeek=1. */
'use strict';
const assert=require('node:assert/strict');
const fs=require('node:fs');
const {execFileSync}=require('node:child_process');
const legacy=execFileSync('git',['show','77e6add481767829bc76f6e4564ac214c0be4928:assets/app.js'],{encoding:'utf8'});
const latest=fs.readFileSync('assets/app.js','utf8');
const doc={createElement:()=>({style:{setProperty(){}},innerHTML:''})};
const escape=value=>String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
function pageFrom(source) {
  const start=source.indexOf('function makePrintPage(');
  const end=source.indexOf('  // The lightweight preview calls',start);
  assert(start>=0 && end>start);
  return new Function('document','html',source.slice(start,end)+'\nreturn makePrintPage;')(doc,escape);
}
const renderOld=pageFrom(legacy),renderNew=pageFrom(latest);
const week=(n,span)=>({week:n,kind:'week',label:n+'. HAFTA',span});
const exam=n=>({week:n,kind:'exam',label:'VİZE · '+n+'. HAFTA',span:1});
const rows=Array.from({length:17},(_,i)=>({rank:i+1,number:'12345',name:'ÖĞRENCİ '+i}));
for(const groups of [
  [week(1,1),week(2,1),exam(3),week(4,1)],
  [week(1,3),week(2,3),exam(3),week(4,3)],
  [week(1,5),exam(2),week(3,5)]
]){
  const sheet={groups,name:'01',title:'YOKLAMA',course:'D',instructor:'H',department:'B',
               fixedWidthsMm:[7.5,21,58],markWidthMm:15,fontSizePt:8,fontFamily:'Arial'};
  const original=renderOld(sheet,rows,1,1,groups.length).innerHTML;
  assert.equal(renderNew({...sheet,startWeek:1},rows,1,1,groups.length).innerHTML,original);
  assert.equal(renderNew(sheet,rows,1,1,groups.length).innerHTML,original);
}
console.log('v11 → v12 first-week PDF/preview HTML: identical');
