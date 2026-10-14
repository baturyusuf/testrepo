/* Render the production PDF/preview page function with a minimal DOM stub. */
const assert = require('node:assert/strict');
const fs = require('node:fs');
const src = fs.readFileSync('assets/app.js','utf8');
const start = src.indexOf('function makePrintPage(');
const end = src.indexOf('  // The lightweight preview calls', start);
assert(start >= 0 && end > start, 'Production page renderer not located');

const doc = {
  createElement: () => {
    const result = {
      style: {setProperty(){}},
      innerHTML: ''
    };
    return result;
  }
};
const escapeHtml = value => String(value ?? '').replace(/[&<>"']/g, c => ({
  '&':'&amp;', '<':'&lt;', '>':'&gt;', '"':'&quot;', "'":'&#39;'
}[c]));
const render = new Function('document','html',
  src.slice(start,end) + '\nreturn makePrintPage;')(doc,escapeHtml);

function htmlFor(groups) {
  return render({
    title:'2026-2027 GÜZ YOKLAMA', course:'BLG103 - Programlama',
    instructor:'Deneme Öğretim Üyesi', department:'Bilgisayar Mühendisliği',
    name:'Haftalar', groups, fixedWidthsMm:[7.5,21,58],
    markWidthMm:15, fontSizePt:8, fontFamily:'Arial', rowHeightPt:23.5
  }, [{rank:1,number:'9999999',name:'Öğrenci 1'}],1,1,1).innerHTML;
}

const single=htmlFor([
  {week:1,kind:'week',label:'1. HAFTA',span:1},
  {week:2,kind:'week',label:'2. HAFTA',span:1},
  {week:3,kind:'exam',label:'VİZE SINAVI · 3. HAFTA',span:1}
]);
assert.match(single, /class="week-top single-signature"/);
assert.match(single, /<th class="week-header"\s*>1\. HAFTA<\/th>/);
assert.match(single, /<th class="week-header"\s*>2\. HAFTA<\/th>/);
assert.doesNotMatch(single,/İMZA/,'1 signature must not show numbered subheadings');
assert.doesNotMatch(single,/class="week-sub"/,'1 signature must not reserve a subheader row');
assert.match(single, /VİZE SINAVI/);

for (const count of [2,3,4,5]) {
  const multi=htmlFor([{week:1,kind:'week',label:'1. HAFTA',span:count}]);
  assert.match(multi, new RegExp('class="week-header" colspan="'+count+'"'));
  assert.match(multi,/class="week-sub"/);
  for(let i=1;i<=count;i++) assert(multi.includes(i+'. İMZA'));
}
const mixed=htmlFor([
  {week:1,kind:'week',label:'1. HAFTA',span:1},
  {week:2,kind:'week',label:'2. HAFTA',span:3}
]);
assert.match(mixed,/class="week-header" rowspan="2">1\. HAFTA<\/th>/);
assert.equal((mixed.match(/1\. İMZA/g)||[]).length,1);
console.log('Single vs multiple weekly signature PDF and preview header render: OK');
