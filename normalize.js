// 계정번호(Account Number) 정규화: "Acc-" 등 접두어를 제거하고 숫자만 남긴다.
// 예: "Acc-123456" -> "123456", "ACC 123456" -> "123456"
export function normalizeAccountNumber(raw) {
  if (raw === null || raw === undefined) return "";
  const digitsOnly = String(raw).replace(/[^0-9]/g, "");
  return digitsOnly;
}

// 전화번호 정규화(비교용): 숫자만 남긴다.
export function normalizePhone(raw) {
  if (raw === null || raw === undefined) return "";
  return String(raw).replace(/[^0-9]/g, "");
}

// 이름 정규화(비교용): 공백 제거 + 소문자화
export function normalizeName(raw) {
  if (raw === null || raw === undefined) return "";
  return String(raw).trim().replace(/\s+/g, "").toLowerCase();
}

// 엑셀 날짜 셀(문자열 또는 숫자 serial)을 YYYY-MM-DD 문자열로 변환
export function normalizeDate(raw) {
  if (raw === null || raw === undefined || raw === "") return null;
  if (raw instanceof Date) {
    return toISODate(raw);
  }
  if (typeof raw === "number") {
    // 엑셀 날짜 serial number (1900 날짜 시스템 기준)
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
