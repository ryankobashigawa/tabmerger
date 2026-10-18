(() => {
  "use strict";

  const PREVIEW_LIMIT = 200;
  const HEADER_PREVIEW_LIMIT = 8;
  const DIFF_SHEETS_SHOWN = 3;

  /** @type {Array<Object>} each item is a core.js Sheet plus {id, included} */
  let sheets = [];
  let sheetSeq = 0;
  const fileNameCounts = {};

  let strategy = "union"; // 'union' | 'intersection'
  let includeSourceFile = true;
  let includeSourceSheet = true;

  // ---------- DOM refs ----------
  const statusLine = document.getElementById("status-line");
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

  // ---------- Status helper ----------
  function setStatus(msg, kind) {
    statusLine.textContent = msg || "";
    statusLine.classList.remove("error", "good");
    if (kind) statusLine.classList.add(kind);
  }

  // ---------- File reading (browser-only glue; pure logic lives in core.js) ----------
  function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error);
      reader.readAsArrayBuffer(file);
    });
  }

  function readFileAsText(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error);
      reader.readAsText(file);
    });
  }

  async function ingestFiles(fileList) {
    const files = Array.from(fileList);
    const accepted = files.filter((f) => TM.hasAcceptedExtension(f.name));
    const rejected = files.filter((f) => !TM.hasAcceptedExtension(f.name));

    if (accepted.length === 0) {
      const suffix = rejected.length ? ` Skipped: ${rejected.map((f) => f.name).join(", ")}.` : "";
      setStatus(`No .xlsx/.xls/.csv files found in that selection.${suffix}`, "error");
      return;
    }

    setStatus(`Reading ${accepted.length} file${accepted.length > 1 ? "s" : ""}…`);
    let addedSheets = 0;
    const errors = [];

    for (const file of accepted) {
      try {
        fileNameCounts[file.name] = (fileNameCounts[file.name] || 0) + 1;
        const fileLabel = TM.buildSourceLabel(file.name, fileNameCounts[file.name]);
        const isCsv = file.name.toLowerCase().endsWith(".csv");
        const source = isCsv ? await readFileAsText(file) : await readFileAsArrayBuffer(file);
        const parsed = TM.parseSheetsFromSource(source, { isCsv, fileName: file.name, fileLabel });
        parsed.forEach((sheet) => {
          sheet.id = `sheet-${++sheetSeq}`;
          sheet.included = !sheet.isEmpty && sheet.hasHeaderRow;
          sheets.push(sheet);
          addedSheets += 1;
        });
      } catch (err) {
        errors.push(`Couldn't read "${file.name}": ${err.message || err}`);
      }
    }

    const parts = [];
    if (addedSheets) {
      parts.push(`Loaded ${addedSheets} sheet${addedSheets !== 1 ? "s" : ""} from ${accepted.length} file${accepted.length !== 1 ? "s" : ""}.`);
    }
    if (rejected.length) {
      parts.push(`Skipped unsupported file${rejected.length !== 1 ? "s" : ""}: ${rejected.map((f) => f.name).join(", ")}.`);
    }
    if (errors.length) parts.push(errors.join(" "));
    setStatus(parts.join(" "), errors.length ? "error" : "good");

    renderSheetList();
    recompute();
  }

  // ---------- Directory traversal (for dropped folders) ----------
  function readAllDirectoryEntries(directoryReader) {
    return new Promise((resolve, reject) => {
      const entries = [];
      function readBatch() {
        directoryReader.readEntries((batch) => {
          if (!batch.length) {
            resolve(entries);
          } else {
            entries.push(...batch);
            readBatch();
          }
        }, reject);
      }
      readBatch();
    });
  }

  async function collectFilesFromEntry(entry, out) {
    if (entry.isFile) {
      await new Promise((resolve) => {
        entry.file((file) => {
          if (TM.hasAcceptedExtension(file.name)) out.push(file);
          resolve();
        }, () => resolve());
      });
    } else if (entry.isDirectory) {
      const entries = await readAllDirectoryEntries(entry.createReader());
      for (const child of entries) {
        await collectFilesFromEntry(child, out);
      }
    }
  }

  async function filesFromDataTransfer(dataTransfer) {
    const items = dataTransfer.items;
    const out = [];
    if (items && items.length && items[0].webkitGetAsEntry) {
      const entries = Array.from(items)
        .map((item) => (item.webkitGetAsEntry ? item.webkitGetAsEntry() : null))
        .filter(Boolean);
      for (const entry of entries) {
        await collectFilesFromEntry(entry, out);
      }
      return out;
    }
    return Array.from(dataTransfer.files || []);
  }

  // ---------- Dropzone wiring ----------
  function wireDropzone(zoneEl, onFiles) {
    zoneEl.addEventListener("click", () => {
      const input = zoneEl.querySelector('input[type="file"]:not(#folder-input)');
      if (input) input.click();
    });
    zoneEl.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        const input = zoneEl.querySelector('input[type="file"]:not(#folder-input)');
        if (input) input.click();
      }
    });
    ["dragenter", "dragover"].forEach((evt) =>
      zoneEl.addEventListener(evt, (e) => {
        e.preventDefault();
        zoneEl.classList.add("dragover");
      })
    );
    ["dragleave", "dragend"].forEach((evt) =>
      zoneEl.addEventListener(evt, () => zoneEl.classList.remove("dragover"))
    );
    zoneEl.addEventListener("drop", async (e) => {
      e.preventDefault();
      zoneEl.classList.remove("dragover");
      const files = await filesFromDataTransfer(e.dataTransfer);
      onFiles(files);
    });
  }

  const dropzone = document.getElementById("dropzone");
  const fileInput = document.getElementById("file-input");
  const folderInput = document.getElementById("folder-input");
  wireDropzone(dropzone, ingestFiles);
  fileInput.addEventListener("change", (e) => {
    if (e.target.files.length) ingestFiles(e.target.files);
    e.target.value = "";
  });
  folderInput.addEventListener("change", (e) => {
    if (e.target.files.length) ingestFiles(e.target.files);
    e.target.value = "";
  });
  document.getElementById("browse-folder").addEventListener("click", (e) => {
    e.stopPropagation();
    folderInput.click();
  });

  // ---------- Sheet list rendering (grouped by source file) ----------
  function truncatedHeaderPreview(headers) {
    if (headers.length === 0) return "—";
    const shown = headers.slice(0, HEADER_PREVIEW_LIMIT).join(", ");
    const extra = headers.length - HEADER_PREVIEW_LIMIT;
    return extra > 0 ? `${shown}, +${extra} more` : shown;
  }

  function renderSheetList() {
    sheetsPanel.classList.toggle("hidden", sheets.length === 0);
    columnsPanel.classList.toggle("hidden", sheets.length === 0);
    sheetGroups.innerHTML = "";

    const groups = [];
    const groupByLabel = new Map();
    sheets.forEach((sheet) => {
      if (!groupByLabel.has(sheet.fileLabel)) {
        const g = { fileLabel: sheet.fileLabel, items: [] };
        groupByLabel.set(sheet.fileLabel, g);
        groups.push(g);
      }
      groupByLabel.get(sheet.fileLabel).items.push(sheet);
    });

    groups.forEach((group) => {
      const wrapper = document.createElement("li");
      wrapper.className = "file-group";

      const header = document.createElement("div");
      header.className = "file-group-header";

      const nameEl = document.createElement("span");
      nameEl.className = "file-group-name";
      nameEl.textContent = group.fileLabel;

      const metaEl = document.createElement("span");
      metaEl.className = "source-meta";
      metaEl.textContent = `${group.items.length} sheet${group.items.length !== 1 ? "s" : ""}`;

      const removeBtn = document.createElement("button");
      removeBtn.className = "remove-btn";
      removeBtn.textContent = "✕";
      removeBtn.title = "Remove this file";
      removeBtn.addEventListener("click", () => {
        sheets = sheets.filter((s) => s.fileLabel !== group.fileLabel);
        renderSheetList();
        recompute();
      });

      header.append(nameEl, metaEl, removeBtn);
      wrapper.appendChild(header);

      const list = document.createElement("ul");
      list.className = "sheet-sublist";

      group.items.forEach((sheet) => {
        const li = document.createElement("li");
        li.className = "sheet-row";

        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.checked = sheet.included;
        checkbox.addEventListener("change", () => {
          sheet.included = checkbox.checked;
          recompute();
        });

        const label = document.createElement("span");
        label.className = "sheet-label";
        const nameText = sheet.sheetName || "(CSV)";
        label.textContent = nameText;

        const badge = document.createElement("span");
        if (sheet.isEmpty) {
          badge.className = "badge badge-warn";
          badge.textContent = "empty";
        } else if (!sheet.hasHeaderRow) {
          badge.className = "badge badge-warn";
          badge.textContent = "no header row";
        }

        const meta = document.createElement("span");
        meta.className = "source-meta";
        meta.textContent = `${sheet.rows.length} row${sheet.rows.length !== 1 ? "s" : ""} · ${sheet.headers.length} col${sheet.headers.length !== 1 ? "s" : ""}`;

        const headerPreview = document.createElement("span");
        headerPreview.className = "header-preview";
        headerPreview.textContent = truncatedHeaderPreview(sheet.headers);
        headerPreview.title = sheet.headers.join(", ");

        li.append(checkbox, label);
        if (badge.textContent) li.appendChild(badge);
        li.append(meta, headerPreview);
        list.appendChild(li);
      });

      wrapper.appendChild(list);
      sheetGroups.appendChild(wrapper);
    });
  }

  document.getElementById("clear-all-btn").addEventListener("click", () => {
    sheets = [];
    renderSheetList();
    recompute();
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
      li.innerHTML = `<strong>${entry.label}</strong> <span class="source-meta">— in: ${sheetsText}</span>`;
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
  function triggerDownload(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function getOutputBaseName() {
    const raw = (outputNameInput.value || "consolidated").trim();
    return raw.replace(/\.(xlsx|csv)$/i, "") || "consolidated";
  }

  document.getElementById("export-xlsx-btn").addEventListener("click", () => {
    const last = recompute._last;
    if (!last) return;
    const wb = TM.buildWorkbook(last.columns, last.rows, "Consolidated");
    XLSX.writeFile(wb, `${getOutputBaseName()}.xlsx`);
  });

  document.getElementById("export-csv-btn").addEventListener("click", () => {
    const last = recompute._last;
    if (!last) return;
    const csv = TM.sheetToCsv(last.columns, last.rows);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    triggerDownload(blob, `${getOutputBaseName()}.csv`);
  });
})();
