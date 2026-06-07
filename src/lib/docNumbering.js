const KEY = 'nir_doc_nums';

function load() {
  try { return JSON.parse(localStorage.getItem(KEY) || '{}'); } catch { return {}; }
}

// Assign next number for given ISO date string (e.g. "2025-06-07")
// Returns formatted string like "07.06.25-1"
export function assignDocNumber(isoDate) {
  const nums = load();
  const next = (nums[isoDate] || 0) + 1;
  nums[isoDate] = next;
  localStorage.setItem(KEY, JSON.stringify(nums));
  const [y, m, d] = isoDate.split('-');
  return `${d}.${m}.${y.slice(2)}-${next}`;
}

// Peek at what the next number would be without assigning
export function peekDocNumber(isoDate) {
  const nums = load();
  const next = (nums[isoDate] || 0) + 1;
  const [y, m, d] = isoDate.split('-');
  return `${d}.${m}.${y.slice(2)}-${next}`;
}
