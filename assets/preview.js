/* Client-only A4 layout preview. Uses the same page renderer as PDF export,
 * with the same physical widths and 17 synthetic students; no XLSX parsing. */
(() => {
  'use strict';
  const $ = s => document.querySelector(s);
  const viewport = $('#preview-viewport');
  const canvas = $('#preview-canvas');
  const selector = $('#preview-section');
  const prev = $('#preview-prev');
  const next = $('#preview-next');
  const zoomButton = $('#preview-zoom');
  const measures = $('#preview-measurements');
  const zoomLabel = $('#preview-scale');
  const A4_WIDTH_PX = 297 * 96 / 25.4;
  const A4_HEIGHT_PX = 210 * 96 / 25.4;
  const sampleRows = Array.from({length:17}, (_,i) => ({rank:i+1,number:'9999999',name:`Öğrenci ${i+1}`,nameFontPt:null}));
  let selectedSection = 0;
  let actualSize = false;
  let sheetView = null;

  function currentValues() {
    return {
      weeks:Number($('#week-count').value),
      signatures:Number($('#signature-select').value),
      midterm:Number($('#midterm-week').value),
      final:Number($('#final-week').value),
      markMin:Number($('#mark-width').value),
      nameMax:Number($('#name-max-width').value),
      fontPt:Number($('#font-size').value),
      fontFamily:$('#font-family').value,
      year:$('#academic-year').value,
      semester:$('#semester').value,
      code:$('#course-code').value.trim(),
      course:$('#course-name').value.trim(),
      instructor:$('#instructor').value.trim()
    };
  }

  function toSheet(values, plan, index) {
    const first = plan.start, last = plan.end;
    const label = first === last ? `${first}. Hafta` : `${first}-${last}. Haftalar`;
    return {
      name:`${String(index+1).padStart(2,'0')} - ${label}`,
      title:`İSTANBUL RUMELİ ÜNİVERSİTESİ ${values.year} EĞİTİM ÖĞRETİM YILI ${values.semester.toLocaleUpperCase('tr-TR')} DÖNEMİ DERS YOKLAMA LİSTESİ`,
      course:[values.code||'DERS101',values.course||'Ders Adı'].join(' - '),
      instructor:values.instructor || 'Öğretim Elemanı',
      department:'ÖRNEK BÖLÜM / PROGRAM',
      groups:plan.groups,
      fixedWidthsMm:plan.fixedWidthsMm,
      markWidthMm:plan.markWidthMm,
      totalWidthMm:plan.totalWidthMm,
      fontSizePt:values.fontPt,
      fontFamily:values.fontFamily,
      rowHeightPt:23.5,
      printableWidthMm:289,
      pages:[sampleRows]
    };
  }

  function fitPage() {
    if (!sheetView) return;
    const available = Math.max(1,viewport.clientWidth - 24);
    const scale = actualSize ? 1 : Math.min(.78, available / A4_WIDTH_PX);
    viewport.classList.toggle('is-actual-size',actualSize);
    canvas.style.width = `${A4_WIDTH_PX * scale}px`;
    canvas.style.height = `${A4_HEIGHT_PX * scale}px`;
    sheetView.style.transform = `scale(${scale})`;
    zoomLabel.textContent = `${Math.round(scale*100)}% ekran ölçeği`;
    zoomButton.textContent = actualSize ? 'Ekrana sığdır' : '%100 göster';
    zoomButton.setAttribute('aria-pressed', String(actualSize));
  }

  function render() {
    const values = currentValues();
    try {
      const plan = window.YoklamaHelpers.paperLayout(
        values.weeks, values.signatures, values.midterm, values.final,
        values.markMin, values.nameMax);
      selectedSection = Math.min(selectedSection,plan.length-1);
      selector.replaceChildren(...plan.map((page,i) => {
        const label = `Bölüm ${i+1}/${plan.length} · ${page.start}–${page.end}. haftalar`;
        return new Option(label,String(i));
      }));
      selector.value = String(selectedSection);
      prev.disabled = selectedSection <= 0;
      next.disabled = selectedSection >= plan.length-1;
      const segment = plan[selectedSection];
      const simulated = toSheet(values,segment,selectedSection);
      if (typeof window.YoklamaPrintPage !== 'function')
        throw new Error('Önizleme çizim motoru bulunamadı. Sayfayı yenileyin.');
      sheetView = window.YoklamaPrintPage(simulated,sampleRows,1,1,values.signatures);
      canvas.replaceChildren(sheetView);
      fitPage();
      measures.textContent = `Bölüm ${selectedSection+1}/${plan.length} · ${segment.start}–${segment.end}. haftalar · ${segment.columns} imza/sınav sütunu · Hücre ${segment.markWidthMm.toFixed(1)} mm · Ad Soyad ${values.nameMax} mm`;
      viewport.classList.remove('a4-preview-error');
    } catch (error) {
      sheetView = null;
      canvas.replaceChildren();
      canvas.style.height = '0px';
      canvas.style.width = '0px';
      measures.textContent = `Önizleme oluşturulamadı: ${error.message}`;
      zoomLabel.textContent = '';
      selector.replaceChildren();
      prev.disabled = next.disabled = true;
      viewport.classList.add('a4-preview-error');
    }
  }

  selector.addEventListener('change', () => {selectedSection = Number(selector.value);render();});
  prev.addEventListener('click', () => {selectedSection=Math.max(0,selectedSection-1);render();});
  next.addEventListener('click', () => {selectedSection++;render();});
  zoomButton.addEventListener('click', () => {actualSize=!actualSize;fitPage();});
  for (const id of ['week-count','signature-select','midterm-week','final-week',
      'mark-width','font-size','font-family','name-max-width','academic-year','semester']) {
    $( `#${id}` ).addEventListener('change', () => {selectedSection=0;render();});
  }
  for (const id of ['course-code','course-name','instructor']) {
    $( `#${id}` ).addEventListener('input',render);
  }
  if (typeof ResizeObserver !== 'undefined') new ResizeObserver(fitPage).observe(viewport);
  else window.addEventListener('resize',fitPage);
  render();
})();