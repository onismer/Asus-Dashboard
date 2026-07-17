// ============================================================
// upload.js — file drop, parse + validate, diff vs DB,
// chunked upsert, storage backup, upload history.
// ============================================================
/* global CONFIG, XLSX, buildHeaderLookup, parseTicketRow, findSheets,
   asOnFromFilename, parseStoreSheet */

const Upload = (() => {
  let pending = null; // { rows, storeCodes, fileName, asOn, warnings, errors, diff, file }

  const $ = id => document.getElementById(id);
  const A = () => window.App;

  function bind() {
    const dz = $("dropzone"), fi = $("file-input");
    $("browse-btn").onclick = () => fi.click();
    fi.onchange = () => fi.files[0] && handleFile(fi.files[0]);
    ["dragover", "dragenter"].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.add("drag"); }));
    ["dragleave", "drop"].forEach(ev => dz.addEventListener(ev, e => { e.preventDefault(); dz.classList.remove("drag"); }));
    dz.addEventListener("drop", e => e.dataTransfer.files[0] && handleFile(e.dataTransfer.files[0]));
    $("commit-btn").onclick = commit;
    $("cancel-btn").onclick = reset;
  }

  function reset() {
    pending = null;
    $("validation-card").classList.add("hidden");
    $("parse-progress").classList.add("hidden");
    $("commit-progress").classList.add("hidden");
    $("file-input").value = "";
  }

  function progress(barId, msgId, p, msg) {
    $(barId).style.width = Math.round(p * 100) + "%";
    if (msg) $(msgId).textContent = msg;
  }

  // ---------------- parse & validate ----------------
  async function handleFile(file) {
    reset();
    if (!/\.(xlsx|xls)$/i.test(file.name)) return A().toast("Please upload an Excel file (.xlsx)");
    $("parse-progress").classList.remove("hidden");
    progress("parse-bar", "parse-msg", .1, "Reading file…");
    try {
      const buf = await file.arrayBuffer();
      progress("parse-bar", "parse-msg", .35, "Parsing workbook…");
      // cellDates:false → dates arrive as Excel serial numbers; mapping.js
      // converts them exactly (avoids SheetJS timezone off-by-one issues)
      const wb = XLSX.read(buf, { cellDates: false });
      const { details, stores } = findSheets(wb.SheetNames);
      if (!details) throw new Error(`No "Details Sheet" found. Sheets in file: ${wb.SheetNames.join(", ")}`);

      const raw = XLSX.utils.sheet_to_json(wb.Sheets[details], { defval: null, raw: true });
      if (!raw.length) throw new Error("Details Sheet has no data rows.");
      const headers = Object.keys(raw[0]);
      const lookup = buildHeaderLookup(headers);
      if (lookup.missing.length)
        throw new Error("Required column(s) missing in Details Sheet: " + lookup.missing.join(", "));

      progress("parse-bar", "parse-msg", .55, "Validating rows…");
      const rows = [], errors = [], warnings = [], seenIds = new Map();
      raw.forEach((r, i) => {
        const excelRow = i + 2;
        // skip fully empty rows
        if (Object.values(r).every(v => v == null || String(v).trim() === "")) return;
        const { row, warnings: w, errors: e } = parseTicketRow(r, lookup);
        if (e.length) { errors.push(`Row ${excelRow}: ${e.join("; ")}`); return; }
        if (seenIds.has(row.ticket_id)) {
          errors.push(`Row ${excelRow}: duplicate Ticket ID ${row.ticket_id} (first seen at row ${seenIds.get(row.ticket_id)}) — later row skipped`);
          return;
        }
        seenIds.set(row.ticket_id, excelRow);
        w.forEach(msg => warnings.push(`Row ${excelRow} (T-${row.ticket_id}): ${msg}`));
        rows.push(row);
      });

      let storeCodes = [];
      if (stores) storeCodes = parseStoreSheet(XLSX.utils.sheet_to_json(wb.Sheets[stores], { header: 1, defval: null }));

      progress("parse-bar", "parse-msg", .8, "Comparing with database…");
      const diff = computeDiff(rows);
      pending = { rows, storeCodes, fileName: file.name, asOn: asOnFromFilename(file.name), warnings, errors, diff, file };
      progress("parse-bar", "parse-msg", 1, "Done");
      showValidation(lookup);
    } catch (e) {
      $("parse-progress").classList.add("hidden");
      A().toast("❌ " + e.message, 6000);
    }
  }

  function computeDiff(rows) {
    const existing = new Map(A().S.tickets.map(t => [t.ticket_id, t]));
    const fileIds = new Set(rows.map(r => r.ticket_id));
    let added = 0, updated = 0, unchanged = 0;
    const compareKeys = ["region","branch","store_name","city","issue_raised_date","issue_category",
      "budget_category","status","final_status","rectification_date","responsibility","tat_follow","rectification_time"];
    for (const r of rows) {
      const ex = existing.get(r.ticket_id);
      if (!ex) { added++; continue; }
      const changed = compareKeys.some(k => String(ex[k] ?? "") !== String(r[k] ?? ""));
      changed ? updated++ : unchanged++;
    }
    const missing = [...existing.keys()].filter(id => !fileIds.has(id));
    return { added, updated, unchanged, missing };
  }

  function showValidation(lookup) {
    const { rows, storeCodes, fileName, asOn, warnings, errors, diff } = pending;
    $("validation-card").classList.remove("hidden");
    $("val-filename").textContent = fileName + (asOn ? ` (as on ${A().fmtDate(asOn)})` : "");

    const stats = [
      ["ok",   rows.length, "valid ticket rows"],
      ["ok",   diff.added, "new tickets"],
      ["warn", diff.updated, "tickets will be updated"],
      ["",     diff.unchanged, "unchanged"],
      ["err",  errors.length, "rows with errors (skipped)"],
      ["warn", warnings.length, "warnings"],
      ["",     storeCodes.length, "store codes in master sheet"],
      ["warn", diff.missing.length, "in DB but not in file"],
    ];
    $("val-summary").innerHTML = stats.map(([cls, n, label]) =>
      `<div class="val-stat ${cls}"><b>${Number(n).toLocaleString("en-IN")}</b>${label}</div>`).join("");

    fillList("val-errors", "val-errors-box", errors);
    const extraCols = lookup.extras.length ? [`Info: ${lookup.extras.length} unmapped column(s) preserved in "extra": ${lookup.extras.slice(0, 8).join(", ")}${lookup.extras.length > 8 ? "…" : ""}`] : [];
    fillList("val-warnings", "val-warnings-box", dedupeWarnings(warnings).concat(extraCols));

    // preview
    const prevCols = ["ticket_id","region","branch","store_name","city","issue_raised_date","issue_category","budget_category","status","final_status","rectification_date","responsibility"];
    let html = "<thead><tr>" + prevCols.map(c => `<th>${c}</th>`).join("") + "</tr></thead><tbody>";
    rows.slice(0, 15).forEach(r => {
      html += "<tr>" + prevCols.map(c => `<td>${A().esc(r[c] ?? "-")}</td>`).join("") + "</tr>";
    });
    $("tbl-preview").innerHTML = html + "</tbody>";

    $("del-count").textContent = diff.missing.length;
    $("opt-delete-missing").checked = false;
    $("opt-delete-missing").disabled = diff.missing.length === 0;
    $("commit-btn").disabled = rows.length === 0;
    $("commit-btn").textContent = `Confirm & Upload (${diff.added} new, ${diff.updated} updates)`;
  }

  function dedupeWarnings(warnings) {
    // collapse identical normalization messages: keep counts
    const counts = new Map();
    const rest = [];
    for (const w of warnings) {
      const m = w.match(/normalized (.+)$/);
      if (m) counts.set(m[1], (counts.get(m[1]) || 0) + 1);
      else rest.push(w);
    }
    const out = [...counts.entries()].map(([k, n]) => `Auto-${k} — ${n} row(s)`);
    return out.concat(rest.slice(0, 400));
  }

  function fillList(ulId, boxId, items) {
    const box = $(boxId);
    if (!items.length) { box.classList.add("hidden"); return; }
    box.classList.remove("hidden");
    $(ulId).innerHTML = items.slice(0, 500).map(e => `<li>${A().esc(e)}</li>`).join("") +
      (items.length > 500 ? `<li>…and ${items.length - 500} more</li>` : "");
  }

  // ---------------- commit ----------------
  async function commit() {
    if (!pending) return;
    const sb = A().S.sb;
    const { rows, storeCodes, fileName, asOn, warnings, diff, file } = pending;
    $("commit-btn").disabled = true;
    $("commit-progress").classList.remove("hidden");
    try {
      // 1. upsert tickets in chunks
      const chunk = CONFIG.UPSERT_CHUNK || 500;
      for (let i = 0; i < rows.length; i += chunk) {
        const { error } = await sb.from("tickets").upsert(rows.slice(i, i + chunk), { onConflict: "ticket_id" });
        if (error) throw new Error("Upsert failed: " + error.message);
        progress("commit-bar", "commit-msg", (i + chunk) / (rows.length + 1) * .7, `Uploading tickets… ${Math.min(i + chunk, rows.length)}/${rows.length}`);
      }
      // 2. optional delete of missing tickets
      let deleted = 0;
      if ($("opt-delete-missing").checked && diff.missing.length) {
        for (let i = 0; i < diff.missing.length; i += 200) {
          const ids = diff.missing.slice(i, i + 200);
          const { error } = await sb.from("tickets").delete().in("ticket_id", ids);
          if (error) throw new Error("Delete failed: " + error.message);
          deleted += ids.length;
        }
      }
      // 3. refresh store master (replace-all) if sheet present
      if (storeCodes.length) {
        progress("commit-bar", "commit-msg", .78, "Refreshing store master…");
        const { error: delErr } = await sb.from("stores").delete().neq("store_code", "");
        if (delErr) throw new Error("Store master clear failed: " + delErr.message);
        for (let i = 0; i < storeCodes.length; i += 500) {
          const { error } = await sb.from("stores").insert(storeCodes.slice(i, i + 500).map(c => ({ store_code: c })));
          if (error) throw new Error("Store master insert failed: " + error.message);
        }
      }
      // 4. backup original file to storage (best effort)
      let note = "";
      if ($("opt-backup").checked) {
        progress("commit-bar", "commit-msg", .88, "Backing up original file…");
        const path = `${new Date().toISOString().replace(/[:.]/g, "-")}_${fileName}`;
        const { error } = await sb.storage.from(CONFIG.BACKUP_BUCKET).upload(path, file);
        note = error ? "backup skipped: " + error.message : "backup: " + path;
      }
      // 5. audit log
      progress("commit-bar", "commit-msg", .95, "Writing audit log…");
      await sb.from("upload_logs").insert({
        uploaded_by: A().S.session.user.email, file_name: fileName, as_on_date: asOn,
        total_rows: rows.length, inserted_rows: diff.added, updated_rows: diff.updated,
        deleted_rows: deleted, store_count: storeCodes.length || null,
        warnings: warnings.length, note,
      });
      progress("commit-bar", "commit-msg", 1, "Done — refreshing dashboard…");
      A().toast(`✅ Upload complete: ${diff.added} new, ${diff.updated} updated${deleted ? ", " + deleted + " deleted" : ""}`, 5000);
      reset();
      await A().loadData();
      loadHistory();
    } catch (e) {
      $("commit-btn").disabled = false;
      A().toast("❌ " + e.message, 8000);
      $("commit-msg").textContent = "Failed: " + e.message;
    }
  }

  // ---------------- history ----------------
  async function loadHistory() {
    const sb = A().S.sb;
    const { data, error } = await sb.from("upload_logs").select("*").order("uploaded_at", { ascending: false }).limit(30);
    if (error || !data) return;
    const cols = ["Uploaded At","By","File","As On","Rows","New","Updated","Deleted","Stores","Warnings","Note"];
    let html = "<thead><tr>" + cols.map(c => `<th>${c}</th>`).join("") + "</tr></thead><tbody>";
    for (const r of data) {
      html += `<tr><td>${new Date(r.uploaded_at).toLocaleString("en-IN")}</td><td>${A().esc(r.uploaded_by || "-")}</td><td>${A().esc(r.file_name || "-")}</td><td>${A().fmtDate(r.as_on_date)}</td><td class="num">${r.total_rows ?? "-"}</td><td class="num">${r.inserted_rows ?? "-"}</td><td class="num">${r.updated_rows ?? "-"}</td><td class="num">${r.deleted_rows ?? 0}</td><td class="num">${r.store_count ?? "-"}</td><td class="num">${r.warnings ?? 0}</td><td>${A().esc(r.note || "")}</td></tr>`;
    }
    $("tbl-uploads").innerHTML = html + "</tbody>";
  }

  return { bind, loadHistory };
})();
