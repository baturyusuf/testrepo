const assert = require('node:assert/strict');
const h = require('../assets/helpers.js');
let trials = 0;
for (let weeks = 10; weeks <= 20; weeks++) {
  for (let signatures = 1; signatures <= 5; signatures++) {
    for (const [midterm,final] of [[8,Math.min(16,weeks)],[1,weeks],[weeks-1,weeks],[2,Math.min(9,weeks)]]) {
      const pages = h.paperLayout(weeks, signatures, midterm, final);
      assert.deepEqual(pages[0].start, 1);
      assert.deepEqual(pages.at(-1).end, weeks);
      assert.equal(pages.flatMap(p=>Array.from({length:p.end-p.start+1},(_,i)=>p.start+i)).length, weeks);
      assert(pages.every(p=>p.columns*p.minWidthMm <= 195.001));
      assert(pages.every(p=>p.minWidthMm >= 8.8 && p.minWidthMm <= 18));
      trials++;
    }
  }
}
assert.equal(h.paperLayout(20,5,8,16).length,5);
assert.equal(h.paperLayout(16,3,8,16).length,2);
assert.equal(h.paperLayout(16,1,8,16).length,1);
assert.throws(()=>h.paperLayout(10,3,10,8),RangeError);
console.log(trials,'adaptive layout combinations passed');