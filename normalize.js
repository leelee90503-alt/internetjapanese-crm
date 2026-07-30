// Normalize account number: strip prefixes like "Acc-" and keep digits only.
// e.g. "Acc-123456" -> "123456", "ACC 123456" -> "123456"
export function normalizeAccountNumber(raw) {
  if (raw === null || raw === undefined) return "";
  const digitsOnly = String(raw).replace(/[^0-9]/g, "");
  return digitsOnly;
}

// Normalize phone number (for comparison): keep digits only.
export function normalizePhone(raw) {
  if (raw === null || raw === undefined) return "";
  return String(raw).replace(/[^0-9]/g, "");
}

// Normalize name (for comparison): strip whitespace + lowercase.
export function normalizeName(raw) {
  if (raw === null || raw === undefined) return "";
  return String(raw).trim().replace(/\s+/g, "").toLowerCase();
}

// Convert an Excel date cell (string or numeric serial) to a YYYY-MM-DD string.
export function normalizeDate(raw) {
  if (raw === null || raw === undefined || raw === "") return null;
  if (raw instanceof Date) {
    return toISODate(raw);
  }
  if (typeof raw === "number") {
    // Excel date serial number (based on the 1900 date system)
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const d = new Date(excelEpoch.getTime() + raw * 86400000);
    return toISODate(d);
  }
  const s = String(raw).trim();
  const parsed = new Date(s);
  if (!isNaN(parsed.getTime())) return toISODate(parsed);
  return null;
}

function toISODate(d) {
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, "0");
  const day = String(d.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

export function fmtMoney(n) {
  if (n === null || n === undefined || n === "") return "-";
  const num = Number(n);
  if (isNaN(num)) return String(n);
  return "$" + num.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

export function escapeHtml(str) {
  if (str === null || str === undefined) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}
