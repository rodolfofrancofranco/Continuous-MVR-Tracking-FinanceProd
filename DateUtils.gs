// Utility helpers for safe ISO date parsing and month keys
function parseIsoDate(s) {
  if (!s) return null;
  try {
    // If already a Date object, return as-is
    if (s instanceof Date) return isNaN(s.getTime()) ? null : s;
    var d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  } catch (e) {
    return null;
  }
}

function formatYearMonthKey(d) {
  if (!d) return '';
  return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0');
}
