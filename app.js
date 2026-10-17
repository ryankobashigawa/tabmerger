(() => {
  "use strict";

  const PREVIEW_LIMIT = 200;
  const ACCEPTED_EXT = [".xlsx", ".xls", ".csv"];

  /** @type {{id:string, label:string, fileName:string, sheetName:string, headers:string[], rows:object[], included:boolean}[]} */
  let datasets = [];
  let datasetSeq = 0;

  // ---------- DOM refs ----------
  const statusLine = document.getElementById("status-line");
  const sourcesPanel = document.getElementById("sources-panel");
  const sourceList = document.getElementById("source-list");
  const resultsPanel = document.getElementById("results-panel");
  const summaryChips = document.getElementById("summary-chips");
  const previewTable = document.getElementById("preview-table");
  const previewNote = document.getElementById("preview-note");
  const outputNameInput = document.getElementById("output-name");

  // ---------- Mode switching ----------
  const modeTabs = document.querySelectorAll(".mode-tab");
  modeTabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      modeTabs.forEach((t) => {
        t.classList.toggle("active", t === tab);
        t.setAttribute("aria-selected", t === tab ? "true" : "false");
      });
      document.getElementById("mode-workbook").classList.toggle("hidden", tab.dataset.mode !== "workbook");
      document.getElementById("mode-files").classList.toggle("hidden", tab.dataset.mode !== "files");
    });
  });

  // ---------- Status helper ----------
  function setStatus(msg, kind) {
    statusLine.textContent = msg || "";
    statusLine.classList.remove("error", "good");
    if (kind) statusLine.classList.add(kind);
  }

  function hasAcceptedExt(name) {
    const lower = name.toLowerCase();
    return ACCEPTED_EXT.some((ext) => lower.endsWith(ext));
  }

  // ---------- Parsing ----------
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

  async function parseFileIntoDatasets(file) {
    const name = file.name || "file";
    const isCsv = name.toLowerCase().endsWith(".csv");
    let workbook;
    try {
      if (isCsv) {
        const text = await readFileAsText(file);
        workbook = XLSX.read(text, { type: "string" });
      } else {
        const buf = await readFileAsArrayBuffer(file);
        workbook = XLSX.read(buf, { type: "array" });
      }
    } catch (err) {
      throw new Error(`Couldn't read "${name}": ${err.message || err}`);
    }

    const result = [];
    workbook.SheetNames.forEach((sheetName) => {
      const ws = workbook.Sheets[sheetName];
      if (!ws) return;
      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
      const headerRow = XLSX.utils.sheet_to_json(ws, { header: 1 })[0] || [];
      const headers = headerRow
        .map((h) => (h === undefined || h === null ? "" : String(h).trim()))
        .filter((h) => h !== "");
      result.push({
        id: `ds-${++datasetSeq}`,
        label: isCsv ? name : `${name} — ${sheetName}`,
        fileName: name,
        sheetName: isCsv ? "" : sheetName,
        headers,
        rows,
        included: true,
      });
    });
    return result;
  }

  async function ingestFiles(fileList) {
    const files = Array.from(fileList).filter((f) => hasAcceptedExt(f.name));
    if (files.length === 0) {
      setStatus("No .xlsx/.xls/.csv files found in that selection.", "error");
      return;
    }
    setStatus(`Reading ${files.length} file${files.length > 1 ? "s" : ""}…`);
    let addedSheets = 0;
    const errors = [];
    for (const file of files) {
      try {
        const parsed = await parseFileIntoDatasets(file);
        datasets.push(...parsed);
        addedSheets += parsed.length;
      } catch (err) {
        errors.push(err.message);
      }
    }
    if (errors.length) {
      setStatus(errors.join(" "), "error");
    } else {
      setStatus(`Loaded ${addedSheets} sheet${addedSheets !== 1 ? "s" : ""} from ${files.length} file${files.length > 1 ? "s" : ""}.`, "good");
    }
    renderSources();
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
          if (hasAcceptedExt(file.name)) out.push(file);
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
  function wireDropzone(zoneEl, { onFiles, multiple }) {
    zoneEl.addEventListener("click", () => {
      const input = zoneEl.querySelector('input[type="file"]');
      if (input) input.click();
    });
    zoneEl.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        const input = zoneEl.querySelector('input[type="file"]');
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
      if (!multiple && files.length > 1) {
        setStatus("Drop just one workbook here, or switch to the multi-file mode.", "error");
        return;
      }
      onFiles(files);
    });
  }

  const dzWorkbook = document.getElementById("dropzone-workbook");
  const fiWorkbook = document.getElementById("file-input-workbook");
  wireDropzone(dzWorkbook, { multiple: false, onFiles: ingestFiles });
  fiWorkbook.addEventListener("change", (e) => {
    if (e.target.files.length) ingestFiles(e.target.files);
    e.target.value = "";
  });

  const dzFiles = document.getElementById("dropzone-files");
  const fiFiles = document.getElementById("file-input-files");
  const fiFolder = document.getElementById("folder-input-files");
  wireDropzone(dzFiles, { multiple: true, onFiles: ingestFiles });
  fiFiles.addEventListener("change", (e) => {
    if (e.target.files.length) ingestFiles(e.target.files);
    e.target.value = "";
  });
  fiFolder.addEventListener("change", (e) => {
    if (e.target.files.length) ingestFiles(e.target.files);
    e.target.value = "";
  });
  document.getElementById("browse-folder").addEventListener("click", (e) => {
    e.stopPropagation();
    fiFolder.click();
  });

  // ---------- Sources rendering ----------
  function renderSources() {
    sourcesPanel.classList.toggle("hidden", datasets.length === 0);
    sourceList.innerHTML = "";
    datasets.forEach((ds) => {
      const li = document.createElement("li");
      li.className = "source-item";

      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.checked = ds.included;
      checkbox.addEventListener("change", () => {
        ds.included = checkbox.checked;
        recompute();
      });

      const name = document.createElement("span");
      name.className = "source-name";
      name.textContent = ds.label;
      name.title = ds.label;

      const meta = document.createElement("span");
      meta.className = "source-meta";
      meta.textContent = `${ds.rows.length} row${ds.rows.length !== 1 ? "s" : ""} · ${ds.headers.length} col${ds.headers.length !== 1 ? "s" : ""}`;

      const removeBtn = document.createElement("button");
      removeBtn.className = "remove-btn";
      removeBtn.textContent = "✕";
      removeBtn.title = "Remove this source";
      removeBtn.addEventListener("click", () => {
        datasets = datasets.filter((d) => d.id !== ds.id);
        renderSources();
        recompute();
      });

      li.append(checkbox, name, meta, removeBtn);
      sourceList.appendChild(li);
    });
  }

  document.getElementById("clear-all-btn").addEventListener("click", () => {
    datasets = [];
    renderSources();
    recompute();
    setStatus("Cleared all sources.");
  });

  // ---------- Merge engine ----------
  function computeMerge() {
    const included = datasets.filter((d) => d.included);
    const unionHeaders = [];
    const seen = new Set();
    included.forEach((ds) => {
      ds.headers.forEach((h) => {
        if (!seen.has(h)) {
          seen.add(h);
          unionHeaders.push(h);
        }
      });
    });
    const mergedRows = [];
    included.forEach((ds) => {
      ds.rows.forEach((row) => {
        const outRow = {};
        unionHeaders.forEach((h) => {
          outRow[h] = Object.prototype.hasOwnProperty.call(row, h) ? row[h] : "";
        });
        mergedRows.push(outRow);
      });
    });
    return { unionHeaders, mergedRows, includedCount: included.length };
  }

  function recompute() {
    if (datasets.length === 0) {
      resultsPanel.classList.add("hidden");
      return;
    }
    const { unionHeaders, mergedRows, includedCount } = computeMerge();
    if (includedCount === 0 || unionHeaders.length === 0) {
      resultsPanel.classList.add("hidden");
      return;
    }
    resultsPanel.classList.remove("hidden");

    summaryChips.innerHTML = "";
    const chips = [
      [`${includedCount}`, "source" + (includedCount !== 1 ? "s" : "")],
      [`${mergedRows.length}`, "row" + (mergedRows.length !== 1 ? "s" : "")],
      [`${unionHeaders.length}`, "column" + (unionHeaders.length !== 1 ? "s" : "")],
    ];
    chips.forEach(([num, label]) => {
      const chip = document.createElement("span");
      chip.className = "chip";
      chip.innerHTML = `<strong>${num}</strong> ${label}`;
      summaryChips.appendChild(chip);
    });

    // Build preview table
    previewTable.innerHTML = "";
    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    unionHeaders.forEach((h) => {
      const th = document.createElement("th");
      th.textContent = h;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    previewTable.appendChild(thead);

    const tbody = document.createElement("tbody");
    const shown = mergedRows.slice(0, PREVIEW_LIMIT);
    shown.forEach((row) => {
      const tr = document.createElement("tr");
      unionHeaders.forEach((h) => {
        const td = document.createElement("td");
        td.textContent = row[h];
        tr.appendChild(td);
      });
      tbody.appendChild(tr);
    });
    previewTable.appendChild(tbody);

    previewNote.textContent =
      mergedRows.length > PREVIEW_LIMIT
        ? `Showing first ${PREVIEW_LIMIT} of ${mergedRows.length} rows. The full dataset is included in the download.`
        : `Showing all ${mergedRows.length} rows.`;

    // stash for export
    recompute._last = { unionHeaders, mergedRows };
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
    const ws = XLSX.utils.json_to_sheet(last.mergedRows, { header: last.unionHeaders });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Consolidated");
    XLSX.writeFile(wb, `${getOutputBaseName()}.xlsx`);
  });

  document.getElementById("export-csv-btn").addEventListener("click", () => {
    const last = recompute._last;
    if (!last) return;
    const ws = XLSX.utils.json_to_sheet(last.mergedRows, { header: last.unionHeaders });
    const csv = XLSX.utils.sheet_to_csv(ws);
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    triggerDownload(blob, `${getOutputBaseName()}.csv`);
  });
})();
