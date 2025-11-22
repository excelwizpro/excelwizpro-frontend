// core/diagnostics.js
export function getOfficeDiagnostics() {
  try {
    return {
      host: Office.context?.host || "unknown",
      platform: Office.context?.diagnostics?.platform || "unknown",
      version: Office.context?.diagnostics?.version || "unknown",
      build: Office.context?.diagnostics?.build || "n/a"
    };
  } catch {
    return { host: "unknown", platform: "unknown", version: "unknown" };
  }
}
