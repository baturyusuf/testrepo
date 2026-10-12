/* OBS Yoklama Oluşturucu. Student data never leaves the browser. */
(() => {
  'use strict';
  const $ = (selector) => document.querySelector(selector);
  const input = $('#file-input');
  const dropzone = $('#dropzone');
  const filesUI = $('#file-list');
  const courseUI = $('#course');
  const courseCodeUI = $('#course-code');
  const courseNameUI = $('#course-name');
  const instructorUI = $('#instructor');
  const yearUI = $('#academic-year');
  const semesterUI = $('#semester');
  const helpers = window.YoklamaHelpers;
  if (!helpers) throw new Error('Yoklama tarih ve dosya yardımcıları yüklenemedi.');
  const generateBtn = $('#generate');
  const results = $('#results');
  const status = $('#status');
  const stage = $('#print-stage');
  const encoder = new TextEncoder();
  const uploaded = [];
  let busy = false;
  let built = [];
  let runtime = null;
  let generatedCourse = '';

  // Local device clock, not the server clock. September–January: Güz; February–August: Bahar.
  const defaults = helpers.academicDefaults(new Date());
  for (let start = defaults.startYear - 4; start <= defaults.startYear + 5; start++) {
    const academicYear = `${start}-${start + 1}`;
    yearUI.add(new Option(academicYear, academicYear));
  }
  yearUI.value = defaults.year;
  semesterUI.value = defaults.semester;
  function syncSidebarYear() { $('#sidebar-year').textContent = yearUI.value.replace('-', '–'); $('#footer-year').textContent = yearUI.value.replace('-', '–'); }
  yearUI.addEventListener('change', syncSidebarYear);
  syncSidebarYear();

  const humanBytes = (bytes) => bytes < 1024 * 1024 ? `${Math.max(1, Math.round(bytes / 1024))} KB` : `${(bytes / 1048576).toFixed(1)} MB`;
  const html = (value) => String(value ?? '').replace(/[&<>"']/g, ch => ({'&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'}[ch]));
  const safeName = (value) => String(value).replace(/\\/g, '/').split('/').pop().replace(/[^\p{L}\p{N}_\.\- ]/gu, '_').replace(/^[. ]+|[. ]+$/g, '').slice(0,110) || 'obs.xlsx';
  const tick = () => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve)));
  const signatureCount = () => Number($('#signature-select').value);

  function renderFiles() {
    $('#upload-count').textContent = `(${uploaded.length})`;
    $('#clear-files').hidden = uploaded.length === 0;
    filesUI.innerHTML = uploaded.length ? uploaded.map((file, i) => `<div class="file-item"><span class="file-ok" aria-hidden="true">✓</span><span class="file-icon" aria-hidden="true">X</span><span class="file-info"><b>${html(file.name)}</b><small>${file.studentCount ? `${file.studentCount} öğrenci` : humanBytes(file.size)}</small></span><button class="remove-btn" type="button" title="Dosyayı kaldır" aria-label="${html(file.name)} dosyasını kaldır" data-remove="${i}">×</button></div>`).join('') : '<div class="empty-list">Henüz Excel dosyası seçilmedi.</div>';
    generateBtn.disabled = busy || uploaded.length === 0 || !courseCodeUI.value.trim() || !courseNameUI.value.trim() || !instructorUI.value.trim();
    filesUI.querySelectorAll('[data-remove]').forEach(button => button.addEventListener('click', () => {
      uploaded.splice(Number(button.dataset.remove), 1);
      invalidateResults();
      renderFiles();
    }));
  }

  function invalidateResults() {
    if (busy) return;
    built = [];
    $('#results-placeholder').hidden = false;
    $('#results-ready').hidden = true;
    $('#result-details').hidden = true;
    results.hidden = false;
    status.hidden = true;
  }

  function addFiles(list) {
    for (const file of [...list]) {
      if (!file.name.toLowerCase().endsWith('.xlsx')) {
        showStatus('Dosya türü desteklenmiyor', `${file.name}: yalnızca .xlsx desteklenmektedir.`, true);
        continue;
      }
      // Keep same file names distinct when they originate from different folders.
      if (uploaded.some(existing => existing.name === file.name && existing.size === file.size && existing.lastModified === file.lastModified)) continue;
      uploaded.push(file);
    }
    input.value = '';
    renderFiles();
  }

  dropzone.addEventListener('click', () => input.click());
  dropzone.addEventListener('keydown', (event) => {
    if (event.key === 'Enter' || event.key === ' ') { event.preventDefault(); input.click(); }
  });
  input.addEventListener('change', () => addFiles(input.files));
  for (const eventName of ['dragenter', 'dragover']) dropzone.addEventListener(eventName, event => {
    event.preventDefault(); dropzone.classList.add('dragging');
  });
  for (const eventName of ['dragleave', 'drop']) dropzone.addEventListener(eventName, event => {
    event.preventDefault(); dropzone.classList.remove('dragging');
  });
  dropzone.addEventListener('drop', event => addFiles(event.dataTransfer.files));
  [courseCodeUI,courseNameUI].forEach(field => field.addEventListener('input', () => {
    courseUI.value = `${courseCodeUI.value.trim()} - ${courseNameUI.value.trim()}`;
    invalidateResults();
    renderFiles();
  }));
  instructorUI.addEventListener('input', () => { invalidateResults(); renderFiles(); });
  [yearUI, semesterUI, $('#signature-select')].forEach(el => el.addEventListener('change', invalidateResults));
  $('#clear-files').addEventListener('click', () => { uploaded.length = 0; invalidateResults(); renderFiles(); });

  const help = $('#help-panel');
  const helpToggle = $('#help-toggle');
  function closeHelp() { help.hidden = true; helpToggle.setAttribute('aria-expanded','false'); helpToggle.focus(); }
  helpToggle.addEventListener('click', () => { help.hidden = false; helpToggle.setAttribute('aria-expanded','true'); $('#help-close').focus(); });
  $('#help-close').addEventListener('click', closeHelp);
  help.addEventListener('click', event => { if(event.target === help) closeHelp(); });
  document.addEventListener('keydown', event => { if(event.key === 'Escape' && !help.hidden) closeHelp(); });

  function showStatus(title, detail, failed = false, completed = false) {
    status.hidden = completed && !failed;
    status.classList.toggle('failed', failed);
    status.classList.toggle('success', completed && !failed);
    $('#status-title').textContent = title;
    $('#status-detail').textContent = detail;
  }
  function warnLibraries() {
    if (typeof loadPyodide !== 'function' || typeof JSZip === 'undefined' || typeof html2canvas !== 'function' || !window.jspdf) {
      throw new Error('Gerekli web kütüphaneleri yüklenemedi. İnternet bağlantısını kontrol edip sayfayı yenileyin.');
    }
  }

  async function getRuntime() {
    if (runtime) return runtime;
    showStatus('İşleme motoru hazırlanıyor', 'Python ve Excel kitaplıkları tarayıcıya yükleniyor. İlk açılışta internet gerekir.');
    await tick();
    warnLibraries();
    const py = await loadPyodide({indexURL: 'https://cdn.jsdelivr.net/pyodide/v0.27.7/full/'});
    await py.loadPackage('micropip');
    await py.runPythonAsync("import micropip; await micropip.install('openpyxl==3.1.5')");
    py.FS.mkdirTree('/home/pyodide/work/templates');
    py.FS.mkdirTree('/home/pyodide/work/uploads');
    for (const src of ['yoklama_olustur.py', 'web_bridge.py']) {
      const response = await fetch(`./${src}`);
      if (!response.ok) throw new Error(`Uygulama dosyası yüklenemedi: ${src}`);
      py.FS.writeFile(`/home/pyodide/${src}`, await response.text());
    }
    for (const count of [1, 2, 3]) {
      const file = `Yoklama-${count}-Sign-Template-Ready-70-Students.xlsx`;
      const response = await fetch(`./templates/${file}`);
      if (!response.ok) throw new Error(`Excel şablonu yüklenemedi: ${file}`);
      py.FS.writeFile(`/home/pyodide/work/templates/${file}`, new Uint8Array(await response.arrayBuffer()));
    }
    runtime = py;
    return py;
  }

  function download(blob, filename) {
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 15000);
  }
  function blobFor(data, mime) { return new Blob([data], {type: mime}); }
  function uniquePath(path, seen) {
    let name = path;
    let suffix = 2;
    while (seen.has(name)) {
      const dot = path.lastIndexOf('.');
      name = `${path.slice(0, dot < 0 ? undefined : dot)}-${suffix++}${dot < 0 ? '' : path.slice(dot)}`;
    }
    seen.add(name);
    return name;
  }

  function makePrintPage(sheet, rows, pageNum, pageTotal, signatures) {
    const el = document.createElement('div');
    el.className = 'print-page';
    const groups = sheet.groups;
    const numericCols = groups.reduce((sum, item) => sum + item.span, 0);
    const rankW = 3.9, noW = 10.2, nameW = 21.1;
    const sigWidth = (100 - rankW - noW - nameW) / numericCols;
    const firstLine = groups.map(group => group.kind === 'week'
      ? `<th class="week-header" colspan="${group.span}">${html(group.label)}</th>`
      : `<th class="exam-header" colspan="${group.span}" rowspan="2">${html(group.label)}</th>`).join('');
    const secondLine = groups.map(group => group.kind === 'week'
      ? Array.from({length: group.span}, (_, i) => `<th>${i + 1}. DERS</th>`).join('') : '').join('');
    const body = rows.map(row => `<tr class="student-row" style="height:${signatures === 3 ? '23.2px' : signatures === 2 ? '33.1px' : '30.7px'}"><td class="rank">${html(row.rank)}</td><td class="num">${html(row.number)}</td><td class="name">${html(row.name)}</td>${groups.map(group => Array.from({length:group.span}, () => `<td class="${group.kind === 'exam' ? 'exam-cell' : ''}"></td>`).join('')).join('')}</tr>`).join('');
    const colDefs = `<col style="width:${rankW}%"><col style="width:${noW}%"><col style="width:${nameW}%">` + Array.from({length: numericCols}, () => `<col style="width:${sigWidth}%">`).join('');
    el.innerHTML = `<div class="print-head">${html(sheet.title)}</div><div class="print-meta"><div class="meta-cell"><strong>Ders Kodu ve İsmi</strong><span>${html(sheet.course)}</span></div><div class="meta-cell"><strong>Dersin Öğretim Elemanı ve İmzası</strong><span>${html(sheet.instructor)}</span></div></div><div class="print-department">${html(sheet.department)}</div><table><colgroup>${colDefs}</colgroup><thead><tr class="week-top"><th rowspan="2">SIRA</th><th rowspan="2">NUMARA</th><th rowspan="2">AD SOYAD</th>${firstLine}</tr><tr class="week-sub">${secondLine}</tr></thead><tbody>${body}</tbody></table><span class="page-mark">${html(sheet.name)} · ${pageNum} / ${pageTotal}</span>`;
    return el;
  }

  async function createPDF(sheet, signatures) {
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({orientation: 'landscape', unit: 'mm', format: 'a4', compress: true});
    for (let page = 0; page < sheet.pages.length; page++) {
      const element = makePrintPage(sheet, sheet.pages[page], page + 1, sheet.pages.length, signatures);
      stage.replaceChildren(element);
      await tick();
      const canvas = await html2canvas(element, {backgroundColor:'#ffffff', scale: 1.7, useCORS:false, logging:false, windowWidth:1122, windowHeight:794});
      if (page) pdf.addPage('a4', 'landscape');
      pdf.addImage(canvas.toDataURL('image/jpeg', 0.89), 'JPEG', 0, 0, 297, 210, undefined, 'FAST');
      canvas.width = 0; canvas.height = 0;
    }
    stage.replaceChildren();
    return new Uint8Array(pdf.output('arraybuffer'));
  }

  function renderResults(data) {
    results.hidden = false;
    $('#results-summary').textContent = `${data.results.length} dosya işlendi · ${built.filter(file=>file.ext === 'pdf').length} PDF üretildi`;
    $('#results-placeholder').hidden = true;
    $('#results-ready').hidden = false;
    $('#result-details').hidden = data.results.length === 0 && data.errors.length === 0;
    const xls = built.filter(file => file.ext === 'xlsx');
    const pdf = built.filter(file => file.ext === 'pdf');
    $('#download-excel').disabled = xls.length === 0;
    $('#download-pdf').disabled = pdf.length === 0;
    $('#download-all').disabled = built.length === 0;
    $('#ready-title').textContent = data.errors.length ? 'Belgeler uyarılarla oluşturuldu.' : 'Yoklama listeleri başarıyla oluşturuldu!';
    for(const item of data.results) {
      const uploadedFile = uploaded.find(file => file.name === item.input);
      if(uploadedFile) uploadedFile.studentCount = item.count;
    }
    $('#result-list').innerHTML = data.results.map((item, i) => {
      const related = built.filter(file => file.inputIndex === i);
      const pdfs = related.filter(file => file.ext === 'pdf');
      return `<div class="result-item"><span class="file-icon">✓</span><div class="result-details"><b>${html(item.filename)}</b><small>${item.count} öğrenci · ${item.blankRows} boş satır · ${item.sheets.reduce((n,s)=>n+s.pages.length,0)} sayfa</small>${item.warnings.map(w => `<div class="warn">⚠ ${html(w)}</div>`).join('')}</div><div class="result-actions">${related.filter(file => file.ext === 'xlsx').map(file=>`<button class="download-btn" data-download="${html(file.id)}">↓ Excel</button>`).join('')}${pdfs.map(file=>`<button class="download-btn" data-download="${html(file.id)}">↓ PDF · ${html(file.label)}</button>`).join('')}</div></div>`;
    }).join('');
    $('#error-list').innerHTML = data.errors.map(error=>`<div class="error-item"><b>${html(error.file)}:</b> ${html(error.message)}</div>`).join('');
    // Keep the page position stable; the completed summary is below the generate button.
    document.querySelectorAll('[data-download]').forEach(button => button.addEventListener('click', () => {
      const item = built.find(file=>file.id === button.dataset.download);
      if(item) download(blobFor(item.bytes, item.mime), item.name);
    }));
  }

  async function generate() {
    if (busy || !uploaded.length || !courseCodeUI.value.trim() || !courseNameUI.value.trim() || !instructorUI.value.trim()) return;
    busy = true;
    results.hidden = true;
    status.hidden = false;
    built = [];
    renderFiles();
    const signatures = signatureCount();
    const course = courseUI.value.trim();
    const academicYear = yearUI.value;
    const semester = semesterUI.value;
    try {
      const py = await getRuntime();
      showStatus('Dosyalar okunuyor', `${uploaded.length} OBS listesi tarayıcıda işleniyor.`);
      await tick();
      const descriptors = [];
      for (let i = 0; i < uploaded.length; i++) {
        const f = uploaded[i];
        const path = `/home/pyodide/work/uploads/${i+1}`;
        py.FS.mkdirTree(path);
        py.FS.writeFile(`${path}/${safeName(f.name)}`, new Uint8Array(await f.arrayBuffer()));
        descriptors.push({name:f.name});
      }
      const payload = JSON.stringify({signatures, course, instructor:instructorUI.value.trim(), academicYear, semester, files:descriptors});
      py.globals.set('browser_payload', payload);
      const response = await py.runPythonAsync('import web_bridge; web_bridge.generate(browser_payload)');
      const manifest = JSON.parse(response);
      const seen = new Set();
      for (let i = 0; i < manifest.results.length; i++) {
        const item = manifest.results[i];
        const bytes = py.FS.readFile(item.file);
        built.push({id:`xlsx-${i}`, inputIndex:i, ext:'xlsx',name:item.filename,zipName:uniquePath(`Yoklamalar/${item.filename}`,seen),mime:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', bytes:new Uint8Array(bytes)});
        for (const sheet of item.sheets) {
          showStatus('PDF dosyaları oluşturuluyor', `${item.input} · ${sheet.name} · ${sheet.pages.length} A4 sayfa`);
          await tick();
          try {
            const pdfBytes = await createPDF(sheet, signatures);
            const filename = item.sheets.length === 1 ? item.filename.replace(/\.xlsx$/i,'.pdf') : item.filename.replace(/\.xlsx$/i,` - ${sheet.name}.pdf`);
            built.push({id:`pdf-${i}-${sheet.name}`,inputIndex:i,ext:'pdf',label:sheet.name,name:filename,zipName:uniquePath(`Yoklamalar/PDF/${filename}`,seen),mime:'application/pdf',bytes:pdfBytes});
          } catch(error) {
            stage.replaceChildren();
            manifest.errors.push({file:`${item.filename} (${sheet.name} PDF)`,message:error.message || String(error)});
          }
        }
      }
      generatedCourse = course;
      renderResults(manifest);
      if (manifest.errors.length) showStatus('İşlem uyarılarla tamamlandı', `${built.length} çıktı hazır, ${manifest.errors.length} hata mevcut.`, true);
      else showStatus('Yoklamalar hazır', `${built.length} Excel/PDF dosyası üretildi. İsterseniz tümünü ZIP olarak indirebilirsiniz.`, false, true);
    } catch(error) {
      console.error('OBS generator error:', error);
      showStatus('İşlem tamamlanamadı', error.message || String(error), true);
    } finally {
      busy = false;
      renderFiles();
    }
  }

  async function downloadSelection(ext) {
    const selected = ext === 'all' ? built : built.filter(item => item.ext === ext);
    if (!selected.length || typeof JSZip === 'undefined') return;
    const base = helpers.archiveFilename(generatedCourse).replace(/\.zip$/i, '');
    try {
      if (selected.length === 1 && ext !== 'all') {
        const item = selected[0];
        download(blobFor(item.bytes, item.mime), item.name);
        showStatus('İndirme hazır', `${item.name} indirilmeye hazır.`, false, true);
        return;
      }
      showStatus('ZIP hazırlanıyor', `${selected.length} çıktı bir araya getiriliyor.`);
      await tick();
      const zip = new JSZip();
      for (const item of selected) {
        const filename = ext === 'all' ? item.zipName : item.name;
        zip.file(filename, item.bytes);
      }
      const data = await zip.generateAsync({type:'blob',compression:'DEFLATE',compressionOptions:{level:4}});
      const suffix = ext === 'xlsx' ? ' - Excel' : ext === 'pdf' ? ' - PDF' : '';
      download(data, `${base}${suffix}.zip`);
      showStatus('İndirme hazır', `${selected.length} dosya indirilmeye hazır.`, false, true);
    } catch(error) {
      console.error('Download creation error:', error);
      showStatus('İndirme hazırlanamadı', error.message || String(error), true);
    }
  }

  generateBtn.addEventListener('click', generate);
  $('#download-all').addEventListener('click', () => downloadSelection('all'));
  $('#download-excel').addEventListener('click', () => downloadSelection('xlsx'));
  $('#download-pdf').addEventListener('click', () => downloadSelection('pdf'));
  renderFiles();
})();