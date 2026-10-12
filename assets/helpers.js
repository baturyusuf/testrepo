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
  function paperLayout(weeks, signatures, midterm, final) {
    [weeks, signatures, midterm, final] = [weeks, signatures, midterm, final].map(Number);
    if (![weeks,signatures,midterm,final].every(Number.isInteger) ||
        weeks < 10 || weeks > 20 || signatures < 1 || signatures > 5 ||
        midterm < 1 || midterm >= final || final > weeks)
      throw new RangeError('Vize finalden önce, final haftası dönem içinde olmalıdır.');
    const pages = [], capacity = 195, minWidth = 8.8;
    let group = [], used = 0;
    for (let week = 1; week <= weeks; week++) {
      const span = week === midterm || week === final ? 1 : signatures;
      if (group.length && (used + span) * minWidth > capacity + 1e-9) {
        pages.push(group); group = []; used = 0;
      }
      group.push({week,span}); used += span;
    }
    if (group.length) pages.push(group);
    return pages.map(g=>({
      start:g[0].week,end:g[g.length-1].week,
      columns:g.reduce((n,entry)=>n+entry.span,0),
      minWidthMm:Math.min(18,capacity/g.reduce((n,entry)=>n+entry.span,0))
    }));
  }
  const api = Object.freeze({academicDefaults, archiveFilename, paperLayout});
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.YoklamaHelpers = api;
})(typeof window !== 'undefined' ? window : globalThis);