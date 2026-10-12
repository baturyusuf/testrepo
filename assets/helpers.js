/* Pure, dependency-free helpers for the OBS attendance website. */
(function (root) {
  'use strict';

  // The institution's academic year changes in September. The January exam
  // period belongs to Güz; February to August defaults to Bahar.
  function academicDefaults(date) {
    const month = date.getMonth();
    const currentYear = date.getFullYear();
    const startYear = month >= 8 ? currentYear : currentYear - 1;
    return {
      startYear,
      year: `${startYear}-${startYear + 1}`,
      semester: month >= 8 || month === 0 ? 'Güz' : 'Bahar'
    };
  }

  // Keep Turkish letters, remove filesystem-reserved symbols and path traversal.
  function archiveFilename(course) {
    const cleaned = String(course ?? '')
      .normalize('NFC')
      .replace(/[\\/:*?"<>|\x00-\x1F]/g, ' ')
      .replace(/\s+/g, ' ')
      .replace(/^[. ]+|[. ]+$/g, '')
      .slice(0, 120)
      .trim();
    return `${cleaned || 'Ders'} Yoklamaları.zip`;
  }

  // Physical A4 planner mirrored by adaptive_layout.py (in mm).
  function paperLayout(weeks, signatures, midterm, final, minWidthMm = 15, nameMaxWidthMm = 58) {
    [weeks, signatures, midterm, final, minWidthMm, nameMaxWidthMm] =
      [weeks, signatures, midterm, final, minWidthMm, nameMaxWidthMm].map(Number);
    if (![weeks,signatures,midterm,final].every(Number.isInteger) ||
        weeks < 10 || weeks > 20 || signatures < 1 || signatures > 5 ||
        midterm < 1 || midterm >= final || final > weeks ||
        !Number.isFinite(minWidthMm) || minWidthMm < 8 || minWidthMm > 20 ||
        !Number.isFinite(nameMaxWidthMm) || nameMaxWidthMm < 30 || nameMaxWidthMm > 80)
      throw new RangeError('Sayfa, imza genişliği veya vize/final seçimi geçersiz.');
    // Conservative layout preview assumes the Name column reaches its maximum.
    // Actual sheet calculation uses the longest student name, and may fit more weeks.
    const capacity = 289 - 7.5 - 21 - nameMaxWidthMm;
    const groups = Array.from({length:weeks}, (_,i) => ({
      week:i+1, span:(i+1===midterm || i+1===final) ? 1 : signatures
    }));
    // Identical lexicographic optimization to adaptive_layout.py:
    // fewest sections, then lowest sum of squared signature-column counts.
    const dp = Array(weeks+1).fill(null);
    dp[weeks] = {count:0,squareSum:0,sections:[]};
    for (let start=weeks-1;start>=0;start--) {
      let used=0, best=null;
      for (let end=start+1;end<=weeks;end++) {
        used+=groups[end-1].span;
        if (used*minWidthMm > capacity+1e-9) break;
        const remainder=dp[end];
        if (!remainder) continue;
        const candidate={
          count:remainder.count+1,
          squareSum:remainder.squareSum+used*used,
          sections:[groups.slice(start,end),...remainder.sections]
        };
        if (!best || candidate.count<best.count ||
            (candidate.count===best.count && candidate.squareSum<best.squareSum)) best=candidate;
      }
      dp[start]=best;
    }
    if(!dp[0])throw new RangeError('Geçerli bir A4 sayfa planı oluşturulamadı.');
    const pages=dp[0].sections;
    return pages.map(g=>{
      const columns = g.reduce((n,entry)=>n+entry.span,0);
      const markWidthMm = Math.round(Math.min(Math.max(18,minWidthMm),capacity/columns)*10000)/10000;
      return {
        start:g[0].week,end:g[g.length-1].week,
        columns,
        minWidthMm:markWidthMm, markWidthMm,
        fixedWidthsMm:[7.5,21,nameMaxWidthMm],
        totalWidthMm:7.5+21+nameMaxWidthMm+columns*markWidthMm,
        groups:g.map(entry=>({
          week:entry.week,span:entry.span,
          kind:entry.week===midterm || entry.week===final?'exam':'week',
          label:entry.week===midterm?`VİZE SINAVI · ${entry.week}. HAFTA`
                :entry.week===final?`FİNAL SINAVI · ${entry.week}. HAFTA`:`${entry.week}. HAFTA`
        }))
      };
    });
  }
  const api = Object.freeze({academicDefaults, archiveFilename, paperLayout});
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.YoklamaHelpers = api;
})(typeof window !== 'undefined' ? window : globalThis);