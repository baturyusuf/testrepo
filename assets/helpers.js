/* Pure, dependency-free helpers for the OBS attendance website. */
(function (root) {
  'use strict';

  // Academic year changes in September. January belongs to Güz.
  // February through August defaults to Bahar.
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

  // Keep Turkish letters; remove Windows/macOS/Linux filename-problem characters.
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

  const api = Object.freeze({ academicDefaults, archiveFilename });
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.YoklamaHelpers = api;
})(typeof window !== 'undefined' ? window : globalThis);
