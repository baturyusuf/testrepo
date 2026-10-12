const assert = require('node:assert/strict');
const h = require('../assets/helpers.js');
function reference(weeks, sig, mid, final, mm, nameMax) {
  const spans=Array.from({length:weeks},(_,i)=>i+1===mid||i+1===final?1:sig);
  const max=289-7.5-21-nameMax;
  const dp=Array(weeks+1).fill(null);dp[weeks]={pages:0,score:0};
  for(let i=weeks-1;i>=0;i--){
    let used=0;
    for(let j=i;j<weeks;j++){
      used+=spans[j];
      if(used*mm>max+1e-9)break;
      const r=dp[j+1];if(!r)continue;
      const test={pages:r.pages+1,score:r.score+used*used};
      if(!dp[i]||test.pages<dp[i].pages||test.pages===dp[i].pages&&test.score<dp[i].score)dp[i]=test;
    }
  }
  return dp[0];
}
assert.deepEqual(h.paperLayout(16,1,8,16,15,58).map(x=>x.columns),[8,8]);
let n=0;
for(let weeks=10;weeks<=20;weeks++)
 for(let marks=1;marks<=5;marks++)
  for(let mm of [8,12,15,20])
   for(let name of [30,58,80]){
     const mid=Math.min(8,weeks-1),final=weeks;
     const sheets=h.paperLayout(weeks,marks,mid,final,mm,name);
     const referencePlan=reference(weeks,marks,mid,final,mm,name);
     assert.equal(sheets.length,referencePlan.pages);
     assert.equal(sheets.reduce((s,p)=>s+p.columns*p.columns,0),referencePlan.score);
     assert.deepEqual(sheets.flatMap(p=>p.groups.map(g=>g.week)),
       Array.from({length:weeks},(_,i)=>i+1));
     assert(sheets.every(p=>p.columns*p.markWidthMm<=289-7.5-21-name+0.02));
     n++;
   }
console.log('Balanced A4 sections: '+n+' JS checks passed');
