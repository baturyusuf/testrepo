/* The same print-page renderer is used for PDF and live preview. */
'use strict';
const assert = require('node:assert/strict');
const fs = require('node:fs');
const h = require('../assets/helpers.js');
const source = fs.readFileSync('assets/app.js','utf8');
const from = source.indexOf('function makePrintPage(');
const until = source.indexOf('  // The lightweight preview calls',from);
assert(from>=0 && until>from);
const document = {createElement:()=>({style:{setProperty(){}},innerHTML:''})};
const escapeHtml = value => String(value ?? '').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const render = new Function('document','html',source.slice(from,until)+'\nreturn makePrintPage;')(document,escapeHtml);
const rows = Array.from({length:17},(_,i)=>({rank:i+1,number:'1000',name:'ÖĞRENCİ '+i}));
const week = (n,span=1)=>({week:n,kind:'week',label:n+'. HAFTA',span});
const exam = n=>({week:n,kind:'exam',label:'VİZE SINAVI · '+n+'. HAFTA',span:1});
const page=(groups,startWeek,width=15)=>render({
  title:'TEST',course:'D',instructor:'H',department:'B',name:'01',
  groups,startWeek,fixedWidthsMm:[7.5,21,58],markWidthMm:width,
  fontSizePt:8,fontFamily:'Arial',rowHeightPt:23.5
},rows,1,1,1).innerHTML;
const count=(html,re)=>(html.match(re)||[]).length;

const groups=[week(1,3),week(2,3),week(3,3),exam(4),week(5,3)];
const initial=page(groups,1);
assert.equal(initial,page(groups,undefined),'Week 1 must render the same as before');
assert.doesNotMatch(initial,/closed-/);
assert.equal(count(initial,/<tr class="student-row"/g),17);
const closed=page(groups,3);
assert.equal(count(closed,/closed-block/g),1);
assert.match(closed,/colspan="6" rowspan="17"/);
assert.match(closed,/3\. HAFTADAN ÖNCEKİ HAFTALAR<br>BU LİSTEDE İMZALANMAZ/);
assert.equal(count(closed,/<th class="week-header closed-week"/g),2);
assert.equal(count(closed,/<th class="closed-week">/g),6);
assert.equal(count(closed,/<td class=""><\/td>/g),17*6);
assert.equal(count(closed,/<td class="exam-cell"><\/td>/g),17);
const split=page(groups,6);
assert.equal(count(split,/closed-block/g),2);
assert.match(split,/colspan="9" rowspan="17"/);
assert.match(split,/colspan="3" rowspan="17"/);
assert.doesNotMatch(split,/exam-header[^>]*closed-week/);
assert.match(page([week(1),week(2)],2),/closed-vertical" colspan="1" rowspan="17"/);
assert.match(page([week(1),week(2),week(3)],3),/closed-vertical" colspan="2" rowspan="17"/);
assert.doesNotMatch(page([week(1),week(2),week(3),week(4)],4),/closed-vertical/);
assert.doesNotMatch(page([exam(1),week(2)],2),/closed-/);
assert.equal(h.lastSignatureWeek(16,8,16),15);
assert.equal(h.lastSignatureWeek(16,8,15),16);
assert.equal(h.lastSignatureWeek(10,9,10),8);
const plan=h.paperLayout(16,1,8,16,15,58);
assert.deepEqual(plan.map(p=>[p.start,p.end]),[[1,8],[9,16]]);
assert.deepEqual([1,8,9,10].map(n=>plan.filter(p=>h.hasOpenWeek(p.groups,n)).length),[2,1,1,1]);
assert.match(fs.readFileSync('index.html','utf8'),/<select id="start-week"/);
assert.match(source,/startWeek:firstOpenWeek/);
assert.match(fs.readFileSync('assets/preview.js','utf8'),/'start-week'/);
assert.match(fs.readFileSync('assets/pdf.css','utf8'),/closed-vertical/);
console.log('Start week PDF/preview regression: OK');
