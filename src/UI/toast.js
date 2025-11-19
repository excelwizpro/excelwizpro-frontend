// ui/toast.js
let toastContainer = null;

function ensureToastContainer() {
  if (toastContainer) return toastContainer;
  toastContainer = document.createElement("div");
  toastContainer.className = "exwz-toast-container";
  Object.assign(toastContainer.style, {
    position: "fixed",
    bottom: "12px",
    right: "12px",
    zIndex: 9999,
    display: "flex",
    flexDirection: "column",
    gap: "6px",
    maxWidth: "260px"
  });
  document.body.appendChild(toastContainer);
  return toastContainer;
}

export function showToast(msg, kind = "info") {
  const c = ensureToastContainer();
  const t = document.createElement("div");
  t.className = "exwz-toast";
  t.textContent = msg;

  const base = {
    padding: "8px 10px",
    borderRadius: "6px",
    fontSize: "0.85rem",
    color: "#111",
    boxShadow: "0 2px 10px rgba(0,0,0,0.15)",
    background: "#f3f3f3"
  };

  const palette = {
    info: { background: "#e5f1ff", color: "#084f94" },
    warn: { background: "#fff4ce", color: "#976f00" },
    error: { background: "#fde7e9", color: "#c22" },
    success: { background: "#e6ffed", color: "#0c7a0c" }
  };

  Object.assign(t.style, base, palette[kind] || palette.info);
  c.appendChild(t);
  setTimeout(() => t.remove(), 2600);
}
