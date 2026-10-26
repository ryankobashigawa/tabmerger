// Stack page (/): stack every selected sheet into one sheet.
(() => {
  "use strict";

  const PREVIEW_LIMIT = 200;
  const DIFF_SHEETS_SHOWN = 3;

  /** @type {Array<Object>} each item is a core.js Sheet plus {id, included} */
  let sheets = [];
  let sheetSeq = 0;
  const fileNameCounts = {};
  // The original File objects, kept so they can be handed to Merge into tabs.
  let loadedFiles = []; // [{file, fileLabel}]

  let strategy = "union"; // 'union' | 'intersection'
  let includeSourceFile = true;
  let includeSourceSheet = true;

  // ---------- DOM refs ----------
  const sheetsPanel = document.getElementById("sheets-panel");
  const sheetGroups = document.getElementById("sheet-groups");
  const columnsPanel = document.getElementById("columns-panel");
  const headerComparisonEl = document.getElementById("header-comparison");
  const resultsPanel = document.getElementById("results-panel");
  const summaryChips = document.getElementById("summary-chips");
  const singleSheetNote = document.getElementById("single-sheet-note");
  const previewTable = document.getElementById("preview-table");
  const previewNote = document.getElementById("preview-note");
  const outputNameInput = document.getElementById("output-name");
  const switchNote = document.getElementById("switch-note");
  const setStatus = SMS.statusSetter(document.getElementById("status-line"));

  // ---------- Loading ----------
  async function ingestFiles(fileList) {
    const files = Array.from(fileList);
    if (!files.length) return;
    setStatus(`Reading ${SMS.plural(files.length, "file")}…`);
    const result = await SMS.loadFiles(files, { fileNameCounts });
    result.sheets.forEach((sheet) => {
      sheet.id = `sheet-${++sheetSeq}`;
      sheet.included = !sheet.isEmpty && sheet.hasHeaderRow;
      sheets.push(sheet);
    });
    loadedFiles = loadedFiles.concat(result.loaded);
    const summary = SMS.loadSummary(result);
    setStatus(summary.message, summary.kind);

    renderSheetList();
    recompute();

    // Count a merge once per batch of added files, not on every checkbox or
    // option change (recompute runs on those too).
    if (result.sheets.length && !resultsPanel.classList.contains("hidden")) {
      SMS.track("merge_completed");
    }
  }

  function resetSources() {
    sheets = [];
    loadedFiles = [];
    Object.keys(fileNameCounts).forEach((name) => delete fileNameCounts[name]);
    renderSheetList();
    recompute();
  }

  SMS.wireDropzone({
    dropzone: document.getElementById("dropzone"),
    fileInput: document.getElementById("file-input"),
    folderInput: document.getElementById("folder-input"),
    chooseButton: document.getElementById("choose-files-btn"),
    folderButton: document.getElementById("browse-folder"),
    onFiles: ingestFiles,
  });

  const sampleBtn = document.getElementById("sample-btn");
  sampleBtn.addEventListener("click", async () => {
    // Disabled while loading so a double-click can't interleave two loads.
    sampleBtn.disabled = true;
    try {
      resetSources();
      await ingestFiles(SMS.salesSampleFiles());
      sheetsPanel.scrollIntoView({ behavior: "smooth", block: "start" });
    } finally {
      sampleBtn.disabled = false;
    }
  });

  // ---------- Hand-off to Merge into tabs ----------
  // Saves the loaded files in this browser (IndexedDB) so the merge page can
  // pick them up. If that fails, the link still works; files just aren't carried.
  document.getElementById("switch-to-merge").addEventListener("click", async (e) => {
    if (!loadedFiles.length) return;
    e.preventDefault();
    const href = e.currentTarget.href;
    try {
      // Never let a slow or stuck IndexedDB hold up the link.
      await Promise.race([
        SMS.saveHandoff(loadedFiles.map((f) => f.file)),
        new Promise((resolve) => setTimeout(resolve, 2000)),
      ]);
    } catch (err) {
      // fall through to plain navigation
    }
    window.location.href = href;
  });

  // ---------- Sheet list ----------
  function renderSheetList() {
    sheetsPanel.classList.toggle("hidden", sheets.length === 0);
    columnsPanel.classList.toggle("hidden", sheets.length === 0);
    switchNote.hidden = sheets.length === 0;
    SMS.renderSheetChecklist(sheetGroups, sheets, {
      mode: "stack",
      onToggle: () => recompute(),
      onRemoveFile: (fileLabel) => {
        sheets = sheets.filter((s) => s.fileLabel !== fileLabel);
        loadedFiles = loadedFiles.filter((f) => f.fileLabel !== fileLabel);
        renderSheetList();
        recompute();
      },
    });
  }

  document.getElementById("clear-all-btn").addEventListener("click", () => {
    resetSources();
    setStatus("Cleared all sources.");
  });

  // ---------- Strategy + source-column controls ----------
  document.querySelectorAll('input[name="strategy"]').forEach((radio) => {
    radio.addEventListener("change", () => {
      if (radio.checked) {
        strategy = radio.value;
        recompute();
      }
    });
  });
  document.getElementById("include-source-file").addEventListener("change", (e) => {
    includeSourceFile = e.target.checked;
    recompute();
  });
  document.getElementById("include-source-sheet").addEventListener("change", (e) => {
    includeSourceSheet = e.target.checked;
    recompute();
  });

  // ---------- Header comparison rendering ----------
  function renderHeaderComparison(includedSheets) {
    headerComparisonEl.innerHTML = "";
    if (includedSheets.length === 0) {
      headerComparisonEl.textContent = "Select at least one sheet to compare its columns.";
      return;
    }
    const cmp = TM.compareHeaders(includedSheets);
    if (cmp.allMatch) {
      const p = document.createElement("p");
      p.className = "comparison-ok";
      p.textContent = `✓ All ${includedSheets.length} selected sheet${includedSheets.length !== 1 ? "s" : ""} share the same ${cmp.commonColumns.length} column${cmp.commonColumns.length !== 1 ? "s" : ""}.`;
      headerComparisonEl.appendChild(p);
      return;
    }

    const summary = document.createElement("p");
    summary.className = "comparison-summary";
    summary.textContent = `${cmp.commonColumns.length} column${cmp.commonColumns.length !== 1 ? "s" : ""} shared by all selected sheets; ${cmp.differingColumns.length} column${cmp.differingColumns.length !== 1 ? "s" : ""} only appear in some:`;
    headerComparisonEl.appendChild(summary);

    const list = document.createElement("ul");
    list.className = "diff-list";
    cmp.differingColumns.forEach((entry) => {
      const li = document.createElement("li");
      const shown = entry.sheets.slice(0, DIFF_SHEETS_SHOWN).join(", ");
      const extra = entry.sheets.length - DIFF_SHEETS_SHOWN;
      const sheetsText = extra > 0 ? `${shown}, +${extra} more` : shown;
      const strong = document.createElement("strong");
      strong.textContent = entry.label;
      const where = document.createElement("span");
      where.className = "source-meta";
      where.textContent = ` — in: ${sheetsText}`;
      li.append(strong, " ", where);
      list.appendChild(li);
    });
    headerComparisonEl.appendChild(list);
  }

  // ---------- Merge + preview ----------
  function recompute() {
    if (sheets.length === 0) {
      resultsPanel.classList.add("hidden");
      return;
    }

    const included = sheets.filter((s) => s.included);
    renderHeaderComparison(included);

    if (included.length === 0) {
      resultsPanel.classList.add("hidden");
      return;
    }

    const { columns, rows } = TM.consolidate(included, {
      strategy,
      includeSourceFile,
      includeSourceSheet,
    });

    if (columns.length === 0) {
      resultsPanel.classList.add("hidden");
      return;
    }

    resultsPanel.classList.remove("hidden");

    singleSheetNote.classList.toggle("hidden", included.length !== 1);

    summaryChips.innerHTML = "";
    const chips = [
      [`${included.length}`, "source" + (included.length !== 1 ? "s" : "")],
      [`${rows.length}`, "row" + (rows.length !== 1 ? "s" : "")],
      [`${columns.length}`, "column" + (columns.length !== 1 ? "s" : "")],
    ];
    chips.forEach(([num, label]) => {
      const chip = document.createElement("span");
      chip.className = "chip";
      chip.innerHTML = `<strong>${num}</strong> ${label}`;
      summaryChips.appendChild(chip);
    });

    previewTable.innerHTML = "";
    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    columns.forEach((c) => {
      const th = document.createElement("th");
      th.textContent = c;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    previewTable.appendChild(thead);

    const tbody = document.createElement("tbody");
    rows.slice(0, PREVIEW_LIMIT).forEach((row) => {
      const tr = document.createElement("tr");
      columns.forEach((c) => {
        const td = document.createElement("td");
        td.textContent = row[c];
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    previewTable.appendChild(tbody);

    previewNote.textContent =
      rows.length > PREVIEW_LIMIT
        ? `Showing first ${PREVIEW_LIMIT} of ${rows.length} rows. The full dataset is included in the download.`
        : `Showing all ${rows.length} row${rows.length !== 1 ? "s" : ""}.`;

    recompute._last = { columns, rows };
  }

  // ---------- Export ----------
  function getOutputBaseName() {
    return SMS.cleanOutputName(outputNameInput.value, "consolidated");
  }

  document.getElementById("export-xlsx-btn").addEventListener("click", () => {
    const last = recompute._last;
    if (!last) return;
    const wb = TM.buildWorkbook(last.columns, last.rows, "Consolidated");
    XLSX.writeFile(wb, `${getOutputBaseName()}.xlsx`);
    SMS.track("download_clicked", { tool: "stack", format: "xlsx" });
  });

  document.getElementById("export-csv-btn").addEventListener("click", () => {
    const last = recompute._last;
    if (!last) return;
    const csv = TM.sheetToCsv(last.columns, last.rows);
    SMS.downloadBlob(new Blob([csv], { type: "text/csv;charset=utf-8;" }), `${getOutputBaseName()}.csv`);
    SMS.track("download_clicked", { tool: "stack", format: "csv" });
  });
})();
