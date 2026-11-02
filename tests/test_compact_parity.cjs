/* Check the independent browser planner against Python for v13 scenarios. */
'use strict';
const assert=require('node:assert/strict');
const {execFileSync}=require('node:child_process');
const H=require('../assets/helpers.js');
const data=[];
for(const weeks of [10,12,16,20]) for(const signatures of [1,2,3,5])
for(const minMark of [8,15,20]) for(const nameMax of [30,58,80]){
 const midterm=Math.min(8,weeks-1),final=weeks;
 for(const startWeek of [1,2,Math.floor(weeks/2),weeks-1]){
  if(startWeek===midterm||startWeek===final)continue;
  data.push({weeks,signatures,midterm,final,minMark,nameMax,startWeek});
 }
}
const py=[
  'import json,sys,adaptive_layout as a',
  'data=json.load(sys.stdin)',
  'result=[]',
  'for d in data:',
  ' fixed=(7.5,21,d["nameMax"])',
  ' sections=a.page_plan(d["weeks"],d["signatures"],d["midterm"],d["final"],d["minMark"],fixed,d["startWeek"])',
  ' result.append([{"weeks":[(g.week,g.span) for g in section],"mark":a.mark_width_mm(section,d["minMark"],fixed,d["startWeek"]) if a.has_open_week(section,d["startWeek"]) else 0} for section in sections])',
  'print(json.dumps(result))'
].join('\n');
const python=JSON.parse(execFileSync('python',['-c',py],{input:JSON.stringify(data),encoding:'utf8',maxBuffer:2**23}));
for(let i=0;i<data.length;i++){
 const d=data[i];
 const js=H.paperLayout(d.weeks,d.signatures,d.midterm,d.final,d.minMark,d.nameMax,d.startWeek);
 const expected=python[i];
 assert.deepEqual(js.map(p=>p.groups.map(g=>[g.week,g.span])),expected.map(p=>p.weeks),JSON.stringify(d));
 for(let j=0;j<js.length;j++)
  if(expected[j].mark)assert(Math.abs(js[j].markWidthMm-expected[j].mark)<=.0001,JSON.stringify(d));
}
console.log(data.length+' Python/browser compact page plans agree');
