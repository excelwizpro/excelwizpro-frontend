// UI/toast.js

let container = null;

function getContainer() {
  if (!container) {
    container = document.createElement("div");
    container.className = "exwz-toast-container";
    Object.assign(container.style, {
      position: "fixed",
      bottom: "14px",
      right: "14px",
      zIndex: 99999,
      display: "flex",
      flexDirection: "column",
      gap: "6px",
      maxWidth: "260px",
      pointerEvents: "none"
    });
    document.body.appendChild(container);
  }
  return container;
}

export function showToast(msg, kind = "info") {
  const c = getContainer();
  const t = document.createElement("div");

  t.textContent = msg;
  t.className = "exwz-toast";

  const base = {
    padding: "8px 10px",
    borderRadius: "6px",
    fontSize: "0.85rem",
    fontFamily: "Inter, sans-serif",
    boxShadow: "0 2px 10px rgba(0,0,0,0.25)",
    pointerEvents: "auto"
  };

  const style = {
    info: { background: "#e5f1ff", color: "#084f94" },
    success: { background: "#e6ffed", color: "#0c7a0c" },
    warn: { background: "#fff4ce", color: "#976f00" },
    error: { background: "#fde7e9", color: "#c22" }
  };

  Object.assign(t.style, base, style[kind] || style.info);

  c.appendChild(t);
  setTimeout(() => t.remove(), 2400);
}
