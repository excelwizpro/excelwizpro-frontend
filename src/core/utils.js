// core/utils.js
export function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

export function getEl(id) {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Missing element: #${id}`);
  return el;
}

export function columnIndexToLetter(index) {
  let n = index + 1;
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}

export function normalizeName(name) {
  return String(name || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "_");
}

export function parseQueryParams() {
  try {
    return Object.fromEntries(new URLSearchParams(window.location.search));
  } catch {
    return {};
  }
}
