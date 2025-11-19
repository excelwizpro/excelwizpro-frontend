// core/config.js
export const EXWZ_VERSION = "13.0.0";

export const DEFAULT_API_BASE = "https://excelwizpro-finalapi.onrender.com";

// Column map cache TTL (enterprise-friendly)
export const COLUMN_MAP_TTL_MS = 90 * 1000; // 90s
export const MAX_DATA_ROWS_PER_COLUMN = 50000;

// Backend / network
export const DEFAULT_REQUEST_TIMEOUT = 20000;
export const MAX_BACKEND_RETRIES = 3;
