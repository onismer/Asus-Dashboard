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
  // legacy logo/budget columns (present in the 2025 historical file)
  "logo delivery status":            "logo_delivery_status",
  "logo vendor name":                "logo_vendor",
  "curiour partner":                 "logo_courier",
  "courier partner":                 "logo_courier",
  "mail received date":              "logo_dispatch_date",
  "delivery date":                   "logo_delivery_date",
  "total amount":                    "total_budget",
};
// legacy "Approval Tat" = days to approve → unified approval_days
HEADER_MAP["approval tat"] = "approval_days";

/* ---- NEW TRACKER FORMAT (Non Logo Case / Logo Case / POSM Case sheets) ---- */
const TRACKER_HEADER_MAP = {
  "ticket id":                          "ticket_id",
  "new ticket no.":                     "new_ticket_no",
  "unique cp store code":               "store_code",
  "store name":                         "store_name",
  "address":                            "address",
  "city":                               "city",
  "state":                              "state",
  "state code":                         "state_code",
  "region":                             "region",
  "branch":                             "branch",
  "store type as per asus":             "store_type",
  "issue raised by":                    "issue_raised_by",
  "designation":                        "designation",
  "cmkt name":                          "cmkt_name",
  "city classification":                "city_classification",
  "issue raised date":                  "issue_raised_date",
  "issue reported by creater":          "problem_reported",
  "issue category":                     "issue_category",
  "issue macro category":               "macro_category",
  "budget category":                    "budget_category",
  "issue type":                         "issue_type",
  "responcibility":                     "responsibility",   // (typo in source file)
  "responsibility":                     "responsibility",
  "ticket status":                      "status",
  "final ticket status":                "final_status",
  "approval date":                      "approval_date",
  "cmkt approval days":                 "approval_days",
  "approved closer tat days":           "tat_city_type",    // allowed TAT days
  "rectification date":                 "rectification_date",
  "ticket closer tat days":             "rectification_time",
  "ticket closer tat":                  "tat_follow",
  "asset installation date":            "asset_installation_date",
  "total budget":                       "total_budget",
  // logo logistics (Logo Case sheet)
  "logo delivery status":               "logo_delivery_status",
  "courier partner name":               "logo_courier",
  "logo vendor name":                   "logo_vendor",
  "dispatch details mail received date":"logo_dispatch_date",
  "delivery date":                      "logo_delivery_date",
};
const TRACKER_IGNORED = new Set([
  "sl. no.", "month", "approval month", "rectification month",
  "approval days bucket", "ageing as per creation", "ageing as per approval",
  "ageing bucket as per creation", "ageing bucket as per approval",
  "ageing bucket as per delivery date", "region wise closer formula",
  "ageing wise closer count formula", "ageing wise closer count formula 02",
]);
const TRACKER_REQUIRED = [
  "ticket id", "region", "store name", "city", "issue raised date",
  "issue category", "budget category", "ticket status", "final ticket status",
];

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
  "logo_dispatch_date", "logo_delivery_date",
]);
const NUM_FIELDS  = new Set(["approval_days", "tat_city_type", "rectification_time", "total_budget"]);
// Reference-only date columns: parsed silently (no validation warnings);
// unparseable raw values (e.g. "Q4 2023", "Docket Details Not Available")
// are preserved in `extra` so nothing is lost in the drawer/exports.
const SILENT_DATE_FIELDS = new Set(["asset_installation_date", "execution_tentative_date"]);
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

// Validate y/m/d parts; returns "yyyy-mm-dd" or undefined if not a real date
function ymd(y, mo, d) {
  y = Number(y); mo = Number(mo); d = Number(d);
  if (y < 1990 || y > 2100 || mo < 1 || mo > 12 || d < 1 || d > 31) return undefined;
  const dt = new Date(Date.UTC(y, mo - 1, d));
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== mo - 1 || dt.getUTCDate() !== d) return undefined; // e.g. 31 Feb
  return `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}

// Accepts JS Date, Excel serial number, or dd-mm-yyyy / mm/dd/yyyy / yyyy-mm-dd strings.
// Returns "yyyy-mm-dd", null (blank), or undefined (unreadable → warning, stored as null).
function parseDate(v) {
  if (v == null || v === "" || v === "-") return null;
  if (v instanceof Date && !isNaN(v)) return ymd(v.getFullYear(), v.getMonth() + 1, v.getDate());
  if (typeof v === "number") {
    if (v < 20000 || v > 60000) return undefined;                    // junk serial (e.g. 0 → "00-01-1900")
    return new Date(Math.round((v - 25569) * 86400 * 1000)).toISOString().slice(0, 10);
  }
  const s = String(v).trim();
  let m = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/);           // dd-mm-yyyy (Indian default)
  if (m) {
    let r = ymd(m[3], m[2], m[1]);
    if (r === undefined && Number(m[1]) <= 12) r = ymd(m[3], m[1], m[2]); // fall back to US mm/dd/yyyy
    return r;
  }
  m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);               // yyyy-mm-dd
  if (m) return ymd(m[1], m[2], m[3]);
  const d = new Date(s);
  if (isNaN(d)) return undefined;
  return ymd(d.getFullYear(), d.getMonth() + 1, d.getDate());
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
      if (d === undefined) {
        if (SILENT_DATE_FIELDS.has(dbField)) {
          const rawS = cleanStr(v);                       // keep raw text for reference
          if (rawS) row.extra[normHeader(origHeader)] = rawS;
        } else warnings.push(`unreadable date in "${origHeader}": "${v}"`);
        row[dbField] = null;
      }
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
  // Note: closed tickets without a Rectification Date are EXPECTED
  // (2025 historical batch + rejected tickets) — imported as-is, excluded
  // from closure-date charts with a footnote. No warning raised.
  if (row.rectification_date && row.issue_raised_date && row.rectification_date < row.issue_raised_date)
    warnings.push("Rectification Date earlier than Issue Raised Date");

  // derive year if absent
  if (row.issue_raised_year == null && row.issue_raised_date)
    row.issue_raised_year = Number(row.issue_raised_date.slice(0, 4));
  return { row, warnings, errors };
}

/** Generic header lookup builder. */
function buildLookup(headers, map, required, ignored) {
  const mapped = [], extras = [], missing = [];
  const seen = new Set();
  for (const h of headers) {
    const n = normHeader(h);
    if (!n || ignored.has(n)) continue;
    if (map[n] && !seen.has(map[n])) { mapped.push([h, map[n]]); seen.add(map[n]); }
    else if (!map[n]) extras.push(h);
  }
  for (const req of required) {
    if (!headers.some(h => normHeader(h) === req)) missing.push(req);
  }
  return { mapped, extras, missing };
}
function buildHeaderLookup(headers)        { return buildLookup(headers, HEADER_MAP, REQUIRED_HEADERS, IGNORED_HEADERS); }
function buildTrackerLookup(headers)       { return buildLookup(headers, TRACKER_HEADER_MAP, TRACKER_REQUIRED, TRACKER_IGNORED); }

/**
 * Detect workbook format.
 *  - 'legacy'  → old dashboard extract / historical 2025 file ("Details Sheet")
 *  - 'tracker' → new maintenance tracker (Non Logo / Logo / POSM Case sheets)
 */
function detectWorkbook(sheetNames) {
  const details  = sheetNames.find(n => /details/i.test(n)) || null;
  const stores   = sheetNames.find(n => /total\s*store/i.test(n)) || null;
  const cases    = sheetNames.filter(n => /^(non\s*logo|logo|posm)\s*case/i.test(n.trim()));
  const masterWod= sheetNames.find(n => /master\s*wod/i.test(n)) || null;
  if (details) return { format: "legacy", details, stores, cases: [], masterWod: null };
  if (cases.length) return { format: "tracker", details: null, stores: null, cases, masterWod };
  return { format: null, details: null, stores: null, cases: [], masterWod: null };
}

/** In an array-of-arrays sheet, find the real header row (tracker sheets have
    a meta row on top). Looks for a row containing "Ticket ID" / "Unique CP Store Code". */
function findHeaderRow(aoa, mustContain) {
  for (let i = 0; i < Math.min(aoa.length, 8); i++) {
    if ((aoa[i] || []).some(c => normHeader(c) === mustContain)) return i;
  }
  return -1;
}

/** Derive dashboard fields the tracker file doesn't carry (quarters, years,
    half-year, closure ageing bucket, logo flag). Safe to call on any row. */
function deriveFields(row, logoFlag) {
  const q = d => d ? "Q" + (Math.floor((Number(d.slice(5, 7)) - 1) / 3) + 1) + " " + d.slice(0, 4) : null;
  if (row.issue_raised_year == null && row.issue_raised_date) row.issue_raised_year = Number(row.issue_raised_date.slice(0, 4));
  if (!row.quarter_raised    && row.issue_raised_date)  row.quarter_raised    = q(row.issue_raised_date);
  if (!row.quarter_rectified && row.rectification_date) row.quarter_rectified = q(row.rectification_date);
  if (!row.half_yearly && row.issue_raised_date)
    row.half_yearly = (Number(row.issue_raised_date.slice(5, 7)) <= 6 ? "HY1" : "HY2") + " (" + row.issue_raised_date.slice(0, 4) + ")";
  if (row.rectified_year == null && row.rectification_date) row.rectified_year = Number(row.rectification_date.slice(0, 4));
  if (!row.ageing_closure_bucket && row.final_status === "Closed" && row.rectification_time != null) {
    const d = row.rectification_time;
    row.ageing_closure_bucket = d <= 30 ? "00 to 30 Days" : d <= 60 ? "30 to 60 Days" : d <= 90 ? "60 to 90 Days" : d <= 120 ? "90 to 120 Days" : "Above 120 Days";
  }
  if (!row.logo_flag && logoFlag) row.logo_flag = logoFlag;
  return row;
}

/** Parse the tracker's "Master WOD" sheet (array-of-arrays) → store master rows. */
function parseMasterWod(aoa) {
  const hr = findHeaderRow(aoa, "unique cp store code");
  if (hr < 0) return [];
  const heads = (aoa[hr] || []).map(normHeader);
  const col = name => heads.indexOf(name);
  const ix = {
    code: col("unique cp store code"), name: col("store name"),
    city: col("city name"), state: col("state name"), branch: col("branch"),
    territory: col("territory name"), district: col("district"), tm: col("tm name"),
  };
  const out = new Map();
  for (let i = hr + 1; i < aoa.length; i++) {
    const r = aoa[i] || [];
    const code = cleanStr(r[ix.code]);
    if (!code || !/^CP\d+$/i.test(code)) continue;
    out.set(code.toUpperCase(), {
      store_code: code.toUpperCase(),
      store_name: cleanStr(r[ix.name]), city: cleanStr(r[ix.city]),
      state: cleanStr(r[ix.state]), branch: cleanStr(r[ix.branch]),
      territory: cleanStr(r[ix.territory]), district: cleanStr(r[ix.district]),
      tm_name: cleanStr(r[ix.tm]),
    });
  }
  return [...out.values()];
}

/** Back-compat alias. */
function findSheets(sheetNames) {
  const d = detectWorkbook(sheetNames);
  return { details: d.details, stores: d.stores };
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
  module.exports = { HEADER_MAP, TRACKER_HEADER_MAP, REQUIRED_HEADERS, TRACKER_REQUIRED,
    buildHeaderLookup, buildTrackerLookup, parseTicketRow, detectWorkbook, findHeaderRow,
    deriveFields, parseMasterWod, findSheets, asOnFromFilename, parseStoreSheet,
    parseDate, parseNum, normHeader };
}
