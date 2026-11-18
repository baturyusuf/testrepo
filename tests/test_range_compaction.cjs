'use strict';
const assert=require('node:assert/strict'),fs=require('node:fs'),H=require('../assets/helpers.js');
const source=fs.readFileSync('assets/app.js','utf8');
const from=source.indexOf('function makePrintPage('),end=source.indexOf('  // The lightweight preview calls',from);
const render=new Function('document','html',source.slice(from,end)+'\nreturn makePrintPage;')(
  {createElement:()=>({style:{setProperty(){}},innerHTML:''})},String);
const rows=Array.from({length:17},(_,i)=>({rank:i+1,number:'100',name:'ÖĞRENCİ '+i}));
function check(start,mid=8){
 const sections=H.paperLayout(16,2,mid,16,10,58,start);
 assert(sections.every(p=>p.totalWidthMm<=289.01));
 for(const p of sections.filter(p=>H.hasOpenWeek(p.groups,start))){
  const markup=render({title:'T',course:'C',instructor:'I',department:'D',name:'01',
    groups:p.groups,startWeek:start,markWidthMm:p.markWidthMm,fixedWidthsMm:p.fixedWidthsMm,
    fontSizePt:8,fontFamily:'Arial'},rows,1,1,2).innerHTML;
  for(const g of p.groups.filter(g=>g.kind==='closed')){
    assert(g.endWeek-g.week+1>=4&&g.span===1&&g.widthMm===26);
    assert(markup.includes(g.label));
    assert(markup.includes(`${g.week}-${g.endWeek} arasındaki haftalar bu belgede imzalanamaz`));
  }
  assert.equal((markup.match(/<tr class="student-row"/g)||[]).length,17);
 }
 return sections.flatMap(p=>p.groups);
}
const nine=check(9);
assert(nine.some(g=>g.kind==='closed'&&g.week===1&&g.endWeek===7));
assert(nine.some(g=>g.kind==='exam'&&g.week===8));
assert(check(4).every(g=>g.kind!=='closed'));
assert(check(5).some(g=>g.kind==='closed'&&g.week===1&&g.endWeek===4));
const split=check(9,4);
assert(split.some(g=>g.kind==='closed'&&g.week===5&&g.endWeek===8));
assert(split.some(g=>g.kind==='exam'&&g.week===4));
assert.match(fs.readFileSync('assets/pdf.css','utf8'),/td\.closed-compact \.closed-text/);
console.log('Issue #8 horizontal compact render: OK');
