const assert = require('node:assert/strict');
const helpers = require('../assets/helpers.js');

const date = (year, month, day) => new Date(year, month - 1, day, 12);
for (const [parts, academicYear, semester] of [
  [[2026, 8, 31], '2025-2026', 'Bahar'],
  [[2026, 9, 1], '2026-2027', 'Güz'],
  [[2026, 10, 12], '2026-2027', 'Güz'],
  [[2027, 1, 31], '2026-2027', 'Güz'],
  [[2027, 2, 1], '2026-2027', 'Bahar'],
  [[2027, 8, 31], '2026-2027', 'Bahar'],
  [[2027, 9, 1], '2027-2028', 'Güz']
]) {
  const actual = helpers.academicDefaults(date(...parts));
  assert.equal(actual.year, academicYear);
  assert.equal(actual.semester, semester);
}
assert.equal(helpers.archiveFilename('BLG103 - Programlama I'), 'BLG103 - Programlama I Yoklamaları.zip');
assert.equal(helpers.archiveFilename('AİM103 / Yapay Zekâ: I?'), 'AİM103 Yapay Zekâ I Yoklamaları.zip');
assert.equal(helpers.archiveFilename(' ../ '), 'Ders Yoklamaları.zip');
console.log('Academic defaults and ZIP naming: OK');
