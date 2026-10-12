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
    const pages = [], capacity = 289 - 7.5 - 21 - nameMaxWidthMm;
    let group = [], used = 0;
    for (let week = 1; week <= weeks; week++) {
      const span = week === midterm || week === final ? 1 : signatures;
      if (span * minWidthMm > capacity)
        throw new RangeError('Bir haftalık imzalar bu A4 genişliğine sığmıyor.');
      if (group.length && (used + span) * minWidthMm > capacity + 1e-9) {
        pages.push(group); group = []; used = 0;
      }
      group.push({week,span}); used += span;
    }
    if (group.length) pages.push(group);
    return pages.map(g=>({
      start:g[0].week,end:g[g.length-1].week,
      columns:g.reduce((n,entry)=>n+entry.span,0),
      minWidthMm:Math.min(Math.max(18,minWidthMm),capacity/g.reduce((n,entry)=>n+entry.span,0))
    }));
  }
  const api = Object.freeze({academicDefaults, archiveFilename, paperLayout});
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.YoklamaHelpers = api;
})(typeof window !== 'undefined' ? window : globalThis);