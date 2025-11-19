// ui/errorPanel.js
let panel = null;

export function ensureErrorPanel() {
  if (panel) return panel;
  panel = document.createElement("div");
  panel.className = "exwz-error-panel";
  Object.assign(panel.style, {
    borderRadius: "8px",
    padding: "8px 10px",
    marginTop: "8px",
    background: "#fde7e9",
    color: "#c22",
    fontSize: "0.85rem",
    display: "none"
  });
  const container = document.querySelector("main.container") || document.body;
  container.appendChild(panel);
  return panel;
}

export function showError(message, details) {
  const p = ensureErrorPanel();
  p.style.display = "block";
  p.innerHTML = `<strong>Error:</strong> ${message}`;
  if (details) {
    const pre = document.createElement("pre");
    pre.style.marginTop = "4px";
    pre.style.whiteSpace = "pre-wrap";
    pre.textContent = details;
    p.appendChild(pre);
  }
}

export function clearError() {
  if (!panel) return;
  panel.style.display = "none";
  panel.innerHTML = "";
}
