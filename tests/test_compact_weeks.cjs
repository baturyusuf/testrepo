'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const H = require('../assets/helpers.js');
const renderSource=fs.readFileSync('assets/app.js','utf8');
const from=renderSource.indexOf('function makePrintPage(');
const to=renderSource.indexOf('  // The lightweight preview calls',from);
const render=new Function('document','html',renderSource.slice(from,to)+'\nreturn makePrintPage;')(
  {createElement:()=>({style:{setProperty(){}},innerHTML:''})},
  value=>String(value)
);
const settings=[16,3,8,16,15,58];
const v12=H.paperLayout(...settings);
const compact=H.paperLayout(...settings,4);
assert(compact.length < v12.length, 'Completed weeks should save an A4 section');
const all=compact.flatMap(p=>p.groups);
assert.deepEqual(all.map(g=>g.week),Array.from({length:16},(_,i)=>i+1));
for(const g of all){
  if(g.week<4){assert.equal(g.span,1);assert.equal(g.widthMm,9)}
  else if(g.kind==='week'){assert.equal(g.span,3);assert(g.widthMm>=15)}
  else{assert.equal(g.span,1);assert(g.widthMm>=15)}
}
const first=compact[0];
const rows=Array.from({length:17},(_,i)=>({rank:i+1,number:'100',name:'ÖĞRENCİ '+i}));
const markup=render({title:'TEST',course:'C',instructor:'D',department:'E',
  groups:first.groups,startWeek:4,markWidthMm:first.markWidthMm,
  fixedWidthsMm:first.fixedWidthsMm,fontSizePt:8,fontFamily:'Aptos'},rows,1,1,3).innerHTML;
assert.equal((markup.match(/closed-block/g)||[]).length,1);
assert.match(markup,/class="exam-cell closed-block closed-vertical" colspan="3" rowspan="17"/);
assert.match(markup,/4\. HAFTADAN ÖNCEKİ HAFTALAR/);
const match=markup.match(/<colgroup>(.*?)<\/colgroup>/);
assert(match);
const widths=[...match[1].matchAll(/width:([\d.]+)px/g)].map(m=>Number(m[1]));
assert.equal(widths.length,3+first.columns);
assert(widths.slice(3,6).every(x=>Math.abs(x-9*96/25.4)<0.02));
assert(widths.slice(6).some(x=>x>15*96/25.4-1));
for(let start=2;start<=15;start++){
 const pages=H.paperLayout(...settings,start);
 assert(pages.every(p=>p.totalWidthMm<=289.005));
 assert(pages.flatMap(p=>p.groups).every(g=>
   g.kind==='exam'||g.week>=start ? g.widthMm>=15 : g.widthMm===9));
 assert(pages.every(p=>p.groups.length>0));
}
assert.deepEqual(H.paperLayout(...settings),H.paperLayout(...settings,1));
console.log('v13 compact week renderer and A4 capacity: OK');
