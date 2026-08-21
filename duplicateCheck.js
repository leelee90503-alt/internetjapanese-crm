import { supabase } from "./supabaseClient.js";
import { normalizeAccountNumber, normalizePhone } from "./normalize.js";

// Duplicate-suspect check: if the name + phone + account number combination
// matches existing data, the row is only flagged as a "possible duplicate."
// The system never auto-merges or auto-excludes rows — the final call is
// always made by staff on the review/confirm screen.
//
// - Name+phone matches an existing customer (customers) -> name+phone match
// - Account number matches an existing order service line (order_service_lines) -> account number match
export async function checkDuplicateSuspect({ name, phone, accountNumber }) {
  const reasons = [];
  let matchedCustomerId = null;
  let matchedCustomerLabel = null;

  const normPhone = normalizePhone(phone);
  const normAccount = normalizeAccountNumber(accountNumber);

  // These two lookups are independent of each other, so run them in
  // parallel rather than one-after-another -- with large uploads (e.g. a
  // Bundle Orders file with 70+ blocks), buildBlocks() itself already runs
  // one checkDuplicateSuspect() call per block in parallel (see
  // screen-customerUpload.js), so halving each call's own round-trip count
  // matters for how long the whole "Next: Review & Confirm" step takes.
  const [nameRes, accountRes] = await Promise.all([
    name && normPhone
      ? supabase.from("customers").select("id, name, phone").ilike("name", name.trim()).limit(5)
      : Promise.resolve({ data: null, error: null }),
    normAccount
      ? supabase
          .from("order_service_lines")
          .select("id, account_number, order_id, orders(customer_id, customers(name, phone))")
          .eq("account_number", normAccount)
          .limit(5)
      : Promise.resolve({ data: null, error: null }),
  ]);

  if (name && normPhone && !nameRes.error && nameRes.data) {
    const hit = nameRes.data.find((c) => normalizePhone(c.phone) === normPhone);
    if (hit) {
      reasons.push("Name+phone match");
      matchedCustomerId = hit.id;
      matchedCustomerLabel = `${hit.name} (${hit.phone || "-"})`;
    }
  }

  if (normAccount && !accountRes.error && accountRes.data && accountRes.data.length > 0) {
    reasons.push("Account number match");
    if (!matchedCustomerId) {
      const line = accountRes.data[0];
      const cust = line.orders?.customers;
      matchedCustomerId = line.orders?.customer_id || null;
      matchedCustomerLabel = cust ? `${cust.name} (${cust.phone || "-"})` : null;
    }
  }

  return {
    suspect: reasons.length > 0,
    reason: reasons.join(", "),
    matchedCustomerId,
    matchedCustomerLabel,
  };
}

// Exact-duplicate check used right before saving: if a service line with the
// same Account Number + Service (Product/Package) already exists under a
// customer with the same name, treat this line as already saved and skip
// inserting it again. Unlike checkDuplicateSuspect (which only warns and
// lets staff decide), this one is a hard skip — the combination the request
// asked to de-dupe on (Account Number + Customer Full Name + Product/Package)
// leaves no ambiguity for staff to resolve.
export async function findExactDuplicateLine({ accountNumber, customerName, serviceId }) {
  const normAccount = normalizeAccountNumber(accountNumber);
  const normName = (customerName || "").trim().toLowerCase();
  if (!normAccount || !serviceId || !normName) return false;

  const { data, error } = await supabase
    .from("order_service_lines")
    .select("id, account_number, service_id, orders(customers(name))")
    .eq("account_number", normAccount)
    .eq("service_id", serviceId)
    .limit(20);
  if (error || !data) return false;

  return data.some((row) => (row.orders?.customers?.name || "").trim().toLowerCase() === normName);
}
