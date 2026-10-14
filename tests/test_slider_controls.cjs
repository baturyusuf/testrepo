const assert = require('node:assert/strict');
const fs = require('node:fs');
const html = fs.readFileSync('index.html', 'utf8');
const app = fs.readFileSync('assets/app.js', 'utf8');
const preview = fs.readFileSync('assets/preview.js', 'utf8');
const css = fs.readFileSync('assets/app.css', 'utf8');

for (const [id,min,max,step,initial,unit] of [
  ['mark-width',8,20,1,15,'mm'],
  ['font-size',6,14,1,8,'pt'],
  ['name-max-width',30,80,2,58,'mm']
]) {
  const tag = html.match(new RegExp('<input[^>]+id="'+id+'"[^>]*>'));
  assert(tag, 'Missing input for '+id);
  assert(tag[0].includes('type="range"'), 'Not a native slider: '+id);
  for (const [key,val] of Object.entries({min,max,step,value:initial})) {
    assert(tag[0].includes(key+'="'+val+'"'), id+' invalid '+key);
  }
  assert(html.includes('id="'+id+'-value"'), 'Missing current value for '+id);
  assert(html.includes(initial+' '+unit+'</output>'), 'Wrong default displayed for '+id);
}
assert(html.includes('<select id="font-family"'), 'Keep categorical font choice');
assert(app.includes('syncSlider(slider, unit)'), 'Slider numeric display updater missing');
assert(app.includes("slider.addEventListener('input'"), 'Slider must listen to live input');
assert(app.includes('invalidateResults();'), 'Stale outputs must be invalidated');
assert(preview.includes("addEventListener('input', scheduleSliderPreview)"), 'Preview must render before drag release');
assert(preview.includes('requestAnimationFrame(() =>'), 'Preview updates must be coalesced');
assert(css.includes('.layout-controls input[type="range"]'), 'Slider styling missing');
console.log('Three real-time print-setting sliders: OK');
