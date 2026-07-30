import { supabase } from "./supabaseClient.js";
import { normalizeAccountNumber, normalizePhone } from "./normalize.js";

// 중복 의심 고객 판단: 이름 + 전화번호 + 계정번호 조합이 기존 데이터와 일치하면
// "중복 의심"으로만 표시한다. 시스템이 자동으로 병합/제외하지 않으며,
// 최종 판단은 항상 검토·확인 화면에서 직원이 내린다.
//
// - 이름+전화번호가 기존 고객(customers)과 일치 -> 이름·전화번호 일치
// - 계정번호가 기존 오더 서비스 항목(order_service_lines)의 계정번호와 일치 -> 계정번호 일치
export async function checkDuplicateSuspect({ name, phone, accountNumber }) {
  const reasons = [];
  let matchedCustomerId = null;
  let matchedCustomerLabel = null;

  const normPhone = normalizePhone(phone);
  const normAccount = normalizeAccountNumber(accountNumber);

  if (name && normPhone) {
    const { data, error } = await supabase
      .from("customers")
      .select("id, name, phone")
      .ilike("name", name.trim())
      .limit(5);
    if (!error && data) {
      const hit = data.find((c) => normalizePhone(c.phone) === normPhone);
      if (hit) {
        reasons.push("이름+전화번호 일치");
        matchedCustomerId = hit.id;
        matchedCustomerLabel = `${hit.name} (${hit.phone || "-"})`;
      }
    }
  }

  if (normAccount) {
    const { data, error } = await supabase
      .from("order_service_lines")
      .select("id, account_number, order_id, orders(customer_id, customers(name, phone))")
      .eq("account_number", normAccount)
      .limit(5);
    if (!error && data && data.length > 0) {
      reasons.push("계정번호 일치");
      if (!matchedCustomerId) {
        const line = data[0];
        const cust = line.orders?.customers;
        matchedCustomerId = line.orders?.customer_id || null;
        matchedCustomerLabel = cust ? `${cust.name} (${cust.phone || "-"})` : null;
      }
    }
  }

  return {
    suspect: reasons.length > 0,
    reason: reasons.join(", "),
    matchedCustomerId,
    matchedCustomerLabel,
  };
}
