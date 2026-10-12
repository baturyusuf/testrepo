const assert = require('node:assert/strict');
const h = require('../assets/helpers.js');
let checked = 0;
for (let weeks=10; weeks<=20; weeks++) {
  for (let marks=1; marks<=5; marks++) {
    for (const width of [8,12,15,20]) {
      for (const nameMax of [30,58,80]) {
        const exam = Math.min(8,weeks-1), final=weeks;
        const pages=h.paperLayout(weeks,marks,exam,final,width,nameMax);
        const all=pages.flatMap(p=>p.groups);
        assert.deepEqual(all.map(g=>g.week), Array.from({length:weeks},(_,i)=>i+1));
        assert.equal(all.filter(g=>g.kind==='exam').length,2);
        for (const page of pages) {
          assert.equal(page.columns,page.groups.reduce((n,g)=>n+g.span,0));
          assert.deepEqual(page.fixedWidthsMm,[7.5,21,nameMax]);
          assert(page.markWidthMm >= width-0.0002);
          assert(page.totalWidthMm <= 289.005);
          assert(page.groups.every(g=>g.kind==='exam'?g.span===1:g.span===marks));
          assert(page.groups.every(g=>g.kind==='exam'?/SINAVI/.test(g.label):/HAFTA/.test(g.label)));
        }
        checked++;
      }
    }
  }
}
console.log(`${checked} A4 preview / Excel planner agreement cases passed`);