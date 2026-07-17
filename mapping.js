// ============================================================
// mapping.js — Excel <-> database column mapping, value
// normalization and row parsing. Pure functions (no DOM),
// also loadable in Node for testing.
// ============================================================

/* Header names are matched after: trim, lowercase, collapse spaces.
   So "Quarter Wise Approval Month " matches "quarter wise approval month". */
const HEADER_MAP = {
  "ticket id":                       "ticket_id",
  "new ticket no.":                  "new_ticket_no",
  "freshdesk id":                    "freshdesk_id",
  "region":                          "region",
  "branch":                          "branch",
  "store name":                      "store_name",
  "store type":                      "store_type",
  "city":                            "city",
  "address":                         "address",
  "state":                           "state",
  "state code":                      "state_code",
  "issue raised by":                 "issue_raised_by",
  "designation":                     "designation",
  "cmkt name":                       "cmkt_name",
  "logo/non-logo":                   "logo_flag",
  "city classification":             "city_classification",
  "asset installation date":         "asset_installation_date",
  "issue raised date":               "issue_raised_date",
  "issue raised years":              "issue_raised_year",
  "quarter wise approval month":     "quarter_raised",
  "problem reported":                "problem_reported",
  "issue category":                  "issue_category",
  "issue budget category":           "budget_category",
  "status":                          "status",
  "approval date":                   "approval_date",
  "approval tat":                    "approval_tat",
  "tat as per city type":            "tat_city_type",
  "execution tentative date":        "execution_tentative_date",
  "rectification date":              "rectification_date",
  "rectification time":              "rectification_time",
  "quarter wise rectification month":"quarter_rectified",
  "final status":                    "final_status",
  "half yearly":                     "half_yearly",
  "ageing wise closer":              "ageing_closure_bucket",
  "rectified year":                  "rectified_year",
  "responsibility":                  "responsibility",
  "tat follow":                      "tat_follow",
  "material deployed":               "material_deployed",
  "qty":                             "qty",
};

// Columns that must exist in the uploaded sheet (blocking error if missing)
const REQUIRED_HEADERS = [
  "ticket id", "region", "branch", "store name", "city",
  "issue raised date", "issue category", "issue budget category",
  "status", "final status",
];

// Headers ignored entirely (helper formulas in the workbook)
const IGNORED_HEADERS = new Set([
  "sr. no.", "concatenate", "duplicate", "duplicate check",
  "issue raised month", "month", "rectification months",
  "ageing wise open tickets", "ageing",
]);

const DATE_FIELDS = new Set([
  "asset_installation_date", "issue_raised_date", "approval_date",
  "execution_tentative_date", "rectification_date",
]);
const NUM_FIELDS  = new Set(["approval_tat", "tat_city_type", "rectification_time"]);
const INT_FIELDS  = new Set(["issue_raised_year", "rectified_year"]);

// ---------- value normalization (fixes inconsistent casing in source) ----------
const CANONICAL = {
  status: {
    "rectification done":          "Rectification Done",
    "rectified by rv":             "Rectified by RV",
    "rectified by pkiosk vendor":  "Rectified by PKiosk Vendor",
    "initiated by rv":             "Initiated by RV",
    "initiated by pkiosk vendor":  "Initiated by PKiosk Vendor",
    "approved":                    "Approved",
    "approval pending":            "Approval Pending",
    "asus to align":               "Asus to Align",
    "rejected":                    "Rejected",
  },
  responsibility: { "channelplay": "Channelplay", "asus": "Asus" },
  final_status:   { "closed": "Closed", "open": "Open" },
  region:         { "north": "North", "south": "South", "east": "East", "west": "West" },
  tat_follow:     { "intat": "InTAT", "outtat": "OutTAT" },
  logo_flag:      { "logo": "Logo", "non logo": "Non Logo", "non-logo": "Non Logo" },
};

const VALID_REGIONS = ["North", "South", "East", "West"];
const VALID_FINAL   = ["Open", "Closed"];

function normHeader(h) {
  return String(h == null ? "" : h).trim().toLowerCase().replace(/\s+/g, " ");
}

function cleanStr(v) {
  if (v == null) return null;
  const s = String(v).replace(/\s+/g, " ").trim();
  return s === "" || s === "-" ? null : s;
}

// Accepts JS Date, Excel serial number, or dd-mm-yyyy / dd/mm/yyyy / yyyy-mm-dd strings
function parseDate(v) {
  if (v == null || v === "" || v === "-") return null;
  if (v instanceof Date && !isNaN(v)) {
    return v.getFullYear() + "-" + String(v.getMonth() + 1).padStart(2, "0") + "-" + String(v.getDate()).padStart(2, "0");
  }
  if (typeof v === "number" && v > 20000 && v < 60000) { // Excel serial
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return d.toISOString().slice(0, 10);
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);          // dd-mm-yyyy
  if (m) return `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
  m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);              // yyyy-mm-dd
  if (m) return `${m[1]}-${m[2].padStart(2,"0")}-${m[3].padStart(2,"0")}`;
  const d = new Date(s);
  return isNaN(d) ? undefined : d.toISOString().slice(0, 10);       // undefined = unparseable
}

function parseNum(v) {
  if (v == null || v === "" || v === "-") return null;
  const n = Number(String(v).replace(/,/g, ""));
  return isNaN(n) ? undefined : n;
}

/**
 * Parse one raw sheet row (object keyed by original headers) into
 * { row, rowWarnings, rowErrors }
 */
function parseTicketRow(raw, headerLookup) {
  const row = { extra: {} };
  const warnings = [], errors = [];

  for (const [origHeader, dbField] of headerLookup.mapped) {
    let v = raw[origHeader];
    if (DATE_FIELDS.has(dbField)) {
      const d = parseDate(v);
      if (d === undefined) { warnings.push(`unreadable date in "${origHeader}": "${v}"`); row[dbField] = null; }
      else row[dbField] = d;
    } else if (NUM_FIELDS.has(dbField)) {
      const n = parseNum(v);
      if (n === undefined) { row[dbField] = null; }
      else row[dbField] = n;
    } else if (INT_FIELDS.has(dbField)) {
      const n = parseNum(v);
      row[dbField] = (n === undefined || n === null) ? null : Math.round(n);
    } else if (dbField === "ticket_id") {
      // alphanumeric IDs allowed (e.g. "12563", "L12417"); Excel may give a number
      const s = cleanStr(typeof v === "number" ? String(Math.round(v)) : v);
      row.ticket_id = s ? s.toUpperCase().replace(/\s+/g, "") : null;
    } else {
      let s = cleanStr(v);
      if (s && CANONICAL[dbField]) {
        const canon = CANONICAL[dbField][s.toLowerCase()];
        if (canon && canon !== s) { warnings.push(`normalized ${dbField} "${s}" → "${canon}"`); s = canon; }
        else if (canon) s = canon;
      }
      row[dbField] = s;
    }
  }
  // unmapped, non-ignored columns → extra jsonb
  for (const origHeader of headerLookup.extras) {
    const s = cleanStr(raw[origHeader]);
    if (s != null) row.extra[normHeader(origHeader)] = s;
  }

  // ---- validations ----
  if (row.ticket_id == null) errors.push("missing/invalid Ticket ID");
  if (!row.region) errors.push("missing Region");
  else if (!VALID_REGIONS.includes(row.region)) warnings.push(`unknown Region "${row.region}"`);
  if (!row.issue_raised_date) errors.push("missing Issue Raised Date");
  if (!row.final_status) errors.push("missing Final Status");
  else if (!VALID_FINAL.includes(row.final_status)) warnings.push(`unknown Final Status "${row.final_status}"`);
  if (!row.store_name) warnings.push("missing Store Name");
  if (!row.issue_category) warnings.push("missing Issue Category");
  if (!row.budget_category) warnings.push("missing Budget Category");
  if (row.final_status === "Closed" && !row.rectification_date)
    warnings.push("Closed ticket without Rectification Date");
  if (row.rectification_date && row.issue_raised_date && row.rectification_date < row.issue_raised_date)
    warnings.push("Rectification Date earlier than Issue Raised Date");

  // derive year if absent
  if (row.issue_raised_year == null && row.issue_raised_date)
    row.issue_raised_year = Number(row.issue_raised_date.slice(0, 4));
  return { row, warnings, errors };
}

/** Build header lookup from the sheet's first row headers. */
function buildHeaderLookup(headers) {
  const mapped = [], extras = [], missing = [];
  const seen = new Set();
  for (const h of headers) {
    const n = normHeader(h);
    if (!n || IGNORED_HEADERS.has(n)) continue;
    if (HEADER_MAP[n] && !seen.has(HEADER_MAP[n])) { mapped.push([h, HEADER_MAP[n]]); seen.add(HEADER_MAP[n]); }
    else if (!HEADER_MAP[n]) extras.push(h);
  }
  for (const req of REQUIRED_HEADERS) {
    if (!headers.some(h => normHeader(h) === req)) missing.push(req);
  }
  return { mapped, extras, missing };
}

/** Locate the tickets sheet & store sheet by fuzzy name. */
function findSheets(sheetNames) {
  const details = sheetNames.find(n => /details/i.test(n)) || null;
  const stores  = sheetNames.find(n => /total\s*store/i.test(n)) || null;
  return { details, stores };
}

/** Extract "as on" date from a file name like "... As On 13-07-2026.xlsx" */
function asOnFromFilename(name) {
  const m = String(name).match(/as\s*on\s*(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/i);
  return m ? `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}` : null;
}

/** Pull unique CP store codes out of the "Total Store Covered" sheet rows
    (array-of-arrays). Only the FIRST code-bearing column group = overall universe. */
function parseStoreSheet(aoa) {
  if (!aoa || !aoa.length) return [];
  // find header row containing "Unique CP Store Code"
  let headerRowIdx = -1, colIdx = -1;
  for (let i = 0; i < Math.min(aoa.length, 10); i++) {
    const j = (aoa[i] || []).findIndex(c => /unique\s*cp\s*store\s*code/i.test(String(c || "")));
    if (j >= 0) { headerRowIdx = i; colIdx = j; break; }
  }
  if (headerRowIdx < 0) return [];
  const codes = new Set();
  for (let i = headerRowIdx + 1; i < aoa.length; i++) {
    const v = cleanStr((aoa[i] || [])[colIdx]);
    if (v && /^CP\d+$/i.test(v)) codes.add(v.toUpperCase());
  }
  return [...codes];
}

// Node export for testing
if (typeof module !== "undefined") {
  module.exports = { HEADER_MAP, REQUIRED_HEADERS, buildHeaderLookup, parseTicketRow,
    findSheets, asOnFromFilename, parseStoreSheet, parseDate, parseNum, normHeader };
}
