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
      const det = detectWorkbook(wb.SheetNames);
      if (!det.format)
        throw new Error(`Unrecognised file. Expected either a "Details Sheet" (historical format) or "Non Logo Case / Logo Case / POSM Case" sheets (daily tracker). Sheets found: ${wb.SheetNames.join(", ")}`);

      progress("parse-bar", "parse-msg", .55, "Validating rows…");
      const rows = [], errors = [], warnings = [], seenIds = new Map();
      const extrasInfo = [];

      const takeRow = (row, w, e, excelRow, sheetLabel) => {
        const where = sheetLabel ? `${sheetLabel} row ${excelRow}` : `Row ${excelRow}`;
        if (e.length) { errors.push(`${where}: ${e.join("; ")}`); return; }
        if (seenIds.has(row.ticket_id)) {
          errors.push(`${where}: duplicate Ticket ID ${row.ticket_id} (first seen at ${seenIds.get(row.ticket_id)}) — later row skipped`);
          return;
        }
        seenIds.set(row.ticket_id, where);
        w.forEach(msg => warnings.push(`${where} (T-${row.ticket_id}): ${msg}`));
        rows.push(row);
      };

      let storeRows = [];   // objects for the stores table

      if (det.format === "legacy") {
        // ---- HISTORICAL FILE (imports as protected "Historical: 2025") ----
        const raw = XLSX.utils.sheet_to_json(wb.Sheets[det.details], { defval: null, raw: true });
        if (!raw.length) throw new Error("Details Sheet has no data rows.");
        const lookup = buildHeaderLookup(Object.keys(raw[0]));
        if (lookup.missing.length)
          throw new Error("Required column(s) missing in Details Sheet: " + lookup.missing.join(", "));
        if (lookup.extras.length) extrasInfo.push(`${lookup.extras.length} unmapped column(s) preserved: ${lookup.extras.slice(0, 8).join(", ")}${lookup.extras.length > 8 ? "…" : ""}`);
        raw.forEach((r, i) => {
          if (Object.values(r).every(v => v == null || String(v).trim() === "")) return;
          const { row, warnings: w, errors: e } = parseTicketRow(r, lookup);
          if (!e.length) { deriveFields(row, null); row.data_source = "Historical: 2025"; row.frozen = true; }
          takeRow(row, w, e, i + 2, null);
        });
        if (det.stores) storeRows = parseStoreSheet(XLSX.utils.sheet_to_json(wb.Sheets[det.stores], { header: 1, defval: null }))
          .map(c => ({ store_code: c }));
      } else {
        // ---- DAILY TRACKER (imports as "Live") ----
        for (const sheetName of det.cases) {
          const aoa = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: null, raw: true });
          const hr = findHeaderRow(aoa, "ticket id");
          if (hr < 0) { warnings.push(`Sheet "${sheetName}": could not find a header row with "Ticket ID" — sheet skipped`); continue; }
          const headers = aoa[hr].map(h => h == null ? "" : String(h));
          const lookup = buildTrackerLookup(headers);
          if (lookup.missing.length) {
            errors.push(`Sheet "${sheetName}": required column(s) missing: ${lookup.missing.join(", ")} — sheet skipped`);
            continue;
          }
          const logoFlag = /^non\s*logo/i.test(sheetName.trim()) ? "Non Logo" : /^logo/i.test(sheetName.trim()) ? "Logo" : "Non Logo";
          for (let i = hr + 1; i < aoa.length; i++) {
            const arr = aoa[i] || [];
            const obj = {};
            headers.forEach((h, j) => { if (h) obj[h] = arr[j] ?? null; });
            const { row, warnings: w, errors: e } = parseTicketRow(obj, lookup);
            // tracker sheets carry formula-filled blank rows → skip silently when
            // there is no ticket id AND no store name
            if (row.ticket_id == null && !row.store_name) continue;
            if (!e.length) { deriveFields(row, logoFlag); row.data_source = "Live"; row.frozen = false; }
            takeRow(row, w, e, i + 1, sheetName);
          }
        }
        if (!rows.length && !errors.length) throw new Error("No ticket rows found in the tracker's case sheets.");
        if (det.masterWod)
          storeRows = parseMasterWod(XLSX.utils.sheet_to_json(wb.Sheets[det.masterWod], { header: 1, defval: null, raw: true }));
      }

      progress("parse-bar", "parse-msg", .8, "Comparing with database…");
      const diff = computeDiff(rows);
      // "as on" date: from filename, else latest activity date in the file
      let asOn = asOnFromFilename(file.name);
      if (!asOn) {
        const ds = rows.flatMap(r => [r.issue_raised_date, r.rectification_date]).filter(Boolean).sort();
        asOn = ds.length ? ds[ds.length - 1] : null;
      }
      pending = { rows: diff.importRows, storeCodes: storeRows, fileName: file.name, asOn,
        warnings, errors, diff, file, format: det.format, extrasInfo };
      progress("parse-bar", "parse-msg", 1, "Done");
      showValidation();
    } catch (e) {
      $("parse-progress").classList.add("hidden");
      A().toast("Error: " + e.message, 6000);
    }
  }

  function computeDiff(rows) {
    const existing = new Map(A().S.tickets.map(t => [t.ticket_id, t]));
    const fileIds = new Set(rows.map(r => r.ticket_id));
    let added = 0, updated = 0, unchanged = 0, frozenSkipped = 0;
    const importRows = [];
    const compareKeys = ["region","branch","store_name","city","issue_raised_date","issue_category",
      "budget_category","status","final_status","rectification_date","responsibility","tat_follow",
      "rectification_time","total_budget","approval_days","logo_delivery_status","logo_delivery_date"];
    for (const r of rows) {
      const ex = existing.get(r.ticket_id);
      if (ex && ex.frozen) { frozenSkipped++; continue; }   // protected historical rows are never touched
      if (!ex) { added++; importRows.push(r); continue; }
      const changed = compareKeys.some(k => String(ex[k] ?? "") !== String(r[k] ?? ""));
      if (changed) { updated++; importRows.push(r); } else { unchanged++; importRows.push(r); }
    }
    // deletable = DB tickets absent from the file, EXCLUDING protected historical rows
    const missing = [...existing.values()].filter(t => !t.frozen && !fileIds.has(t.ticket_id)).map(t => t.ticket_id);
    return { added, updated, unchanged, missing, frozenSkipped, importRows };
  }

  function showValidation() {
    const { rows, storeCodes, fileName, asOn, warnings, errors, diff, format, extrasInfo } = pending;
    $("validation-card").classList.remove("hidden");
    const fmtLabel = format === "legacy" ? "HISTORICAL file — imports as protected “Historical: 2025”"
                                         : "DAILY TRACKER — imports as “Live”";
    $("val-filename").textContent = `${fileName}${asOn ? ` (as on ${A().fmtDate(asOn)})` : ""} — ${fmtLabel}`;

    const stats = [
      ["ok",   rows.length, "rows to import"],
      ["ok",   diff.added, "new tickets"],
      ["warn", diff.updated, "tickets will be updated"],
      ["",     diff.unchanged, "unchanged"],
      ["",     diff.frozenSkipped, "protected historical rows (skipped)"],
      ["err",  errors.length, "rows with errors (skipped)"],
      ["warn", warnings.length, "warnings"],
      ["",     storeCodes.length, "stores in master sheet"],
      ["warn", diff.missing.length, "live tickets in DB but not in file"],
    ];
    $("val-summary").innerHTML = stats.map(([cls, n, label]) =>
      `<div class="val-stat ${cls}"><b>${Number(n).toLocaleString("en-IN")}</b>${label}</div>`).join("");

    fillList("val-errors", "val-errors-box", errors);
    fillList("val-warnings", "val-warnings-box", dedupeWarnings(warnings).concat((extrasInfo || []).map(s => "Info: " + s)));

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
    let logId = null;
    try {
      const chunk = CONFIG.UPSERT_CHUNK || 500;
      const wantDelete = $("opt-delete-missing").checked && diff.missing.length;

      // 0. build before-image snapshots (for the Remove-upload rollback)
      progress("commit-bar", "commit-msg", .02, "Preparing rollback snapshots…");
      const existing = new Map(A().S.tickets.map(t => [t.ticket_id, t]));
      const snapshots = [];
      for (const r of rows) {
        const ex = existing.get(r.ticket_id);
        if (!ex) snapshots.push({ ticket_id: r.ticket_id, action: "insert", prev: null });
        else if (!ex.frozen) snapshots.push({ ticket_id: r.ticket_id, action: "update", prev: ex });
      }
      if (wantDelete) for (const id of diff.missing) {
        const ex = existing.get(id);
        if (ex && !ex.frozen) snapshots.push({ ticket_id: id, action: "delete", prev: ex });
      }

      // 1. create the audit-log entry first (snapshots reference it)
      const { data: logRow, error: logErr } = await sb.from("upload_logs").insert({
        uploaded_by: A().S.session.user.email, file_name: fileName, as_on_date: asOn,
        total_rows: rows.length, inserted_rows: diff.added, updated_rows: diff.updated,
        deleted_rows: 0, store_count: storeCodes.length || null,
        warnings: warnings.length, snapshot_rows: snapshots.length,
        note: "in progress…",
      }).select("id").single();
      if (logErr) throw new Error("Audit log failed: " + logErr.message);
      logId = logRow.id;

      // 2. store snapshots
      for (let i = 0; i < snapshots.length; i += chunk) {
        const { error } = await sb.from("upload_snapshots").insert(
          snapshots.slice(i, i + chunk).map(s => ({ ...s, upload_id: logId })));
        if (error) throw new Error("Snapshot save failed: " + error.message);
        progress("commit-bar", "commit-msg", .05 + (i / (snapshots.length + 1)) * .15, `Saving rollback snapshots… ${Math.min(i + chunk, snapshots.length)}/${snapshots.length}`);
      }

      // 3. upsert tickets in chunks
      for (let i = 0; i < rows.length; i += chunk) {
        const { error } = await sb.from("tickets").upsert(rows.slice(i, i + chunk), { onConflict: "ticket_id" });
        if (error) throw new Error("Upsert failed: " + error.message);
        progress("commit-bar", "commit-msg", .2 + (i + chunk) / (rows.length + 1) * .5, `Uploading tickets… ${Math.min(i + chunk, rows.length)}/${rows.length}`);
      }
      // 4. optional delete of missing tickets
      let deleted = 0;
      if (wantDelete) {
        for (let i = 0; i < diff.missing.length; i += 200) {
          const ids = diff.missing.slice(i, i + 200);
          const { error } = await sb.from("tickets").delete().in("ticket_id", ids);
          if (error) throw new Error("Delete failed: " + error.message);
          deleted += ids.length;
        }
      }
      // 5. refresh store master (replace-all) if sheet present
      if (storeCodes.length) {
        progress("commit-bar", "commit-msg", .78, "Refreshing store master…");
        const { error: delErr } = await sb.from("stores").delete().neq("store_code", "");
        if (delErr) throw new Error("Store master clear failed: " + delErr.message);
        for (let i = 0; i < storeCodes.length; i += 500) {
          const { error } = await sb.from("stores").insert(storeCodes.slice(i, i + 500));
          if (error) throw new Error("Store master insert failed: " + error.message);
        }
      }
      // 6. backup original file to storage (best effort)
      let note = "";
      if ($("opt-backup").checked) {
        progress("commit-bar", "commit-msg", .88, "Backing up original file…");
        const path = `${new Date().toISOString().replace(/[:.]/g, "-")}_${fileName}`;
        const { error } = await sb.storage.from(CONFIG.BACKUP_BUCKET).upload(path, file);
        note = error ? "backup skipped: " + error.message : "backup: " + path;
      }
      // 7. finalise the audit log entry
      progress("commit-bar", "commit-msg", .95, "Finalising audit log…");
      await sb.from("upload_logs").update({
        deleted_rows: deleted,
        note: [pending.format === "legacy" ? "historical import (frozen)" : "daily tracker (live)",
               diff.frozenSkipped ? diff.frozenSkipped + " frozen rows skipped" : "", note].filter(Boolean).join("; "),
      }).eq("id", logId);
      progress("commit-bar", "commit-msg", 1, "Done — refreshing dashboard…");
      A().toast(`Upload complete: ${diff.added} new, ${diff.updated} updated${deleted ? ", " + deleted + " deleted" : ""}`, 5000);
      reset();
      await A().loadData();
      loadHistory();
    } catch (e) {
      $("commit-btn").disabled = false;
      A().toast("Error: " + e.message, 8000);
      $("commit-msg").textContent = "Failed: " + e.message;
      if (logId) await sb.from("upload_logs").update({ note: "FAILED: " + e.message }).eq("id", logId);
    }
  }

  // ---------------- remove / rollback an upload ----------------
  async function removeUpload(log, allLogs) {
    const sb = A().S.sb;
    // active (not removed) later uploads with rollback data
    const newer = allLogs.filter(l => !l.removed_at && l.snapshot_rows != null &&
      new Date(l.uploaded_at) > new Date(log.uploaded_at) && !/historical import/i.test(l.note || ""));
    if (!confirm(`Remove upload "${log.file_name}" (${new Date(log.uploaded_at).toLocaleString("en-IN")})?\n\nIts inserted tickets will be deleted and overwritten tickets restored to their previous values.`)) return;
    let targets = [log];
    if (newer.length) {
      const cascade = confirm(`${newer.length} upload(s) were made AFTER this file.\n\nAlso clear those later upload(s)?\n\nOK = YES, remove them too (recommended — keeps data consistent)\nCancel = NO, remove only this upload`);
      if (cascade) targets = [...newer, log];   // newest first
    }
    targets.sort((a, b) => new Date(b.uploaded_at) - new Date(a.uploaded_at));
    A().loading(true, "Rolling back…");
    try {
      for (const t of targets) {
        // fetch this upload's snapshots (paginated)
        const snaps = [];
        for (let from = 0; ; from += 1000) {
          const { data, error } = await sb.from("upload_snapshots").select("ticket_id,action,prev")
            .eq("upload_id", t.id).order("id").range(from, from + 999);
          if (error) throw new Error("Snapshot fetch failed: " + error.message);
          snaps.push(...data);
          if (data.length < 1000) break;
        }
        // restore overwritten/deleted rows to their previous values
        const restore = snaps.filter(s => s.action !== "insert" && s.prev).map(s => s.prev);
        for (let i = 0; i < restore.length; i += 500) {
          const { error } = await sb.from("tickets").upsert(restore.slice(i, i + 500), { onConflict: "ticket_id" });
          if (error) throw new Error("Restore failed: " + error.message);
        }
        // delete rows this upload inserted
        const delIds = snaps.filter(s => s.action === "insert").map(s => s.ticket_id);
        for (let i = 0; i < delIds.length; i += 200) {
          const { error } = await sb.from("tickets").delete().in("ticket_id", delIds.slice(i, i + 200));
          if (error) throw new Error("Rollback delete failed: " + error.message);
        }
        // mark the history row as removed (kept for audit)
        const { error: mErr } = await sb.from("upload_logs").update({
          removed_by: A().S.session.user.email, removed_at: new Date().toISOString(),
        }).eq("id", t.id);
        if (mErr) throw new Error("Could not mark upload removed: " + mErr.message);
      }
      A().toast(`Rolled back ${targets.length} upload(s)`, 5000);
      await A().loadData();
      loadHistory();
    } catch (e) {
      A().toast("Rollback error: " + e.message, 8000);
    } finally { A().loading(false); }
  }

  // ---------------- history ----------------
  async function loadHistory() {
    const sb = A().S.sb;
    if (!sb || typeof sb.from !== "function") return;
    const { data, error } = await sb.from("upload_logs").select("*").order("uploaded_at", { ascending: false }).limit(50);
    if (error || !data) return;
    const isAdmin = A().S.role === "admin";
    const cols = ["Uploaded At","By","File","As On","Rows","New","Updated","Deleted","Stores","Warnings","Note","Action"];
    let html = "<thead><tr>" + cols.map(c => `<th>${c}</th>`).join("") + "</tr></thead><tbody>";
    data.forEach((r, i) => {
      const historical = /historical import/i.test(r.note || "");
      let action;
      if (r.removed_at) action = `<span class="removed-tag">Removed by: ${A().esc(r.removed_by || "?")} (${new Date(r.removed_at).toLocaleDateString("en-IN")})</span>`;
      else if (!isAdmin) action = "";
      else if (historical) action = `<span class="muted" title="Protected historical dataset — cannot be removed from the dashboard">protected</span>`;
      else if (r.snapshot_rows == null) action = `<span class="muted" title="Uploaded before rollback support — no snapshot available">no rollback data</span>`;
      else action = `<button class="btn-mini btn-remove" data-i="${i}">Remove</button>`;
      html += `<tr class="${r.removed_at ? "removed-row" : ""}"><td>${new Date(r.uploaded_at).toLocaleString("en-IN")}</td><td>${A().esc(r.uploaded_by || "-")}</td><td>${A().esc(r.file_name || "-")}</td><td>${A().fmtDate(r.as_on_date)}</td><td class="num">${r.total_rows ?? "-"}</td><td class="num">${r.inserted_rows ?? "-"}</td><td class="num">${r.updated_rows ?? "-"}</td><td class="num">${r.deleted_rows ?? 0}</td><td class="num">${r.store_count ?? "-"}</td><td class="num">${r.warnings ?? 0}</td><td>${A().esc(r.note || "")}</td><td>${action}</td></tr>`;
    });
    $("tbl-uploads").innerHTML = html + "</tbody>";
    $("tbl-uploads").querySelectorAll(".btn-remove").forEach(b =>
      b.onclick = () => removeUpload(data[Number(b.dataset.i)], data));
  }

  return { bind, loadHistory };
})();
