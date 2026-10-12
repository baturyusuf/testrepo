const assert = require('node:assert/strict');
const h = require('../assets/helpers.js');
let trials=0;
for(let weeks=10;weeks<=20;weeks++) for(let signatures=1;signatures<=5;signatures++) {
  for(const [midterm,final] of [[8,Math.min(16,weeks)],[1,weeks],[weeks-1,weeks],[2,Math.min(9,weeks)]]) {
    for(const [minMark,nameMax] of [[8,58],[15,58],[20,30],[20,80]]) {
      const pages=h.paperLayout(weeks,signatures,midterm,final,minMark,nameMax);
      assert.equal(pages[0].start,1);
      assert.equal(pages.at(-1).end,weeks);
      assert.equal(pages.reduce((n,p)=>n+p.end-p.start+1,0),weeks);
      const avail=289-7.5-21-nameMax;
      assert(pages.every(p=>p.columns*p.minWidthMm <= avail+.02));
      assert(pages.every(p=>p.minWidthMm >= minMark-.001 && p.minWidthMm<=Math.max(18,minMark)+.001));
      trials++;
    }
  }
}
assert(h.paperLayout(20,5,8,16,15).length>h.paperLayout(20,5,8,16,8).length);
assert.throws(()=>h.paperLayout(10,3,10,8),RangeError);
assert.throws(()=>h.paperLayout(10,3,8,10,40),RangeError);
console.log(trials,'adaptive layout combinations passed');