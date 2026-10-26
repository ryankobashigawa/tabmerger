// Merge into tabs page (/merge-into-tabs): one workbook, one tab per sheet.
(() => {
  "use strict";

  /** @type {Array<Object>} each item is a core.js Sheet plus {id, included} */
  let sheets = [];
  let sheetSeq = 0;
  const fileNameCounts = {};
  // Tab order as sheet ids; unchecked sheets keep their place but aren't shown.
  let order = [];

  const sheetsPanel = document.getElementById("sheets-panel");
  const sheetGroups = document.getElementById("sheet-groups");
  const tabsPanel = document.getElementById("tabs-panel");
  const tabOrderEl = document.getElementById("tab-order");
  const noTabsNote = document.getElementById("no-tabs-note");
  const summaryChips = document.getElementById("summary-chips");
  const downloadBtn = document.getElementById("download-btn");
  const outputNameInput = document.getElementById("output-name");
  const setStatus = SMS.statusSetter(document.getElementById("status-line"));

  // ---------- Loading ----------
  async function ingestFiles(fileList, intro) {
    const files = Array.from(fileList);
    if (!files.length) return;
    setStatus(`Reading ${SMS.plural(files.length, "file")}…`);
    const result = await SMS.loadFiles(files, { fileNameCounts });
    result.sheets.forEach((sheet) => {
      sheet.id = `sheet-${++sheetSeq}`;
      sheet.included = sheet.hasContent;
      sheets.push(sheet);
      order.push(sheet.id);
    });
    const summary = SMS.loadSummary(result);
    setStatus((intro ? intro + " " : "") + summary.message, summary.kind);
    render();
  }

  function resetSources() {
    sheets = [];
    order = [];
    Object.keys(fileNameCounts).forEach((name) => delete fileNameCounts[name]);
    render();
  }

  SMS.wireDropzone({
    dropzone: document.getElementById("dropzone"),
    fileInput: document.getElementById("file-input"),
    folderInput: document.getElementById("folder-input"),
    chooseButton: document.getElementById("choose-files-btn"),
    folderButton: document.getElementById("browse-folder"),
    onFiles: (files) => ingestFiles(files),
  });

  const sampleBtn = document.getElementById("sample-btn");
  sampleBtn.addEventListener("click", async () => {
    sampleBtn.disabled = true;
    try {
      resetSources();
      await ingestFiles(SMS.salesSampleFiles());
      sheetsPanel.scrollIntoView({ behavior: "smooth", block: "start" });
    } finally {
      sampleBtn.disabled = false;
    }
  });

  document.getElementById("clear-all-btn").addEventListener("click", () => {
    resetSources();
    setStatus("Cleared all sources.");
  });

  // ---------- Tabs ----------
  function orderedIncluded() {
    const byId = new Map(sheets.map((s) => [s.id, s]));
    return order.map((id) => byId.get(id)).filter((s) => s && s.included);
  }

  function tabEntries() {
    const included = orderedIncluded();
    const names = TM.excelTabNames(included);
    return included.map((sheet, i) => ({ sheet, name: names[i], worksheet: sheet.worksheet }));
  }

  // Swap a sheet with its nearest visible neighbour in the given direction.
  function moveSheet(id, direction) {
    const visible = orderedIncluded().map((s) => s.id);
    const pos = visible.indexOf(id);
    const otherId = visible[pos + direction];
    if (pos === -1 || !otherId) return;
    const a = order.indexOf(id);
    const b = order.indexOf(otherId);
    [order[a], order[b]] = [order[b], order[a]];
    render();
    // Keep keyboard focus on the same control of the moved tab.
    const btn = tabOrderEl.querySelector(`[data-id="${id}"][data-dir="${direction}"]`) ||
      tabOrderEl.querySelector(`[data-id="${id}"][data-dir="${-direction}"]`);
    if (btn) btn.focus();
  }

  function moveButton(entry, direction, disabled) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "icon-btn";
    btn.textContent = direction < 0 ? "↑" : "↓";
    btn.dataset.id = entry.sheet.id;
    btn.dataset.dir = String(direction);
    btn.disabled = disabled;
    btn.setAttribute("aria-label", `Move ${entry.name} ${direction < 0 ? "up" : "down"}`);
    btn.addEventListener("click", () => moveSheet(entry.sheet.id, direction));
    return btn;
  }

  function renderTabs() {
    const entries = tabEntries();
    tabOrderEl.innerHTML = "";
    entries.forEach((entry, i) => {
      const li = document.createElement("li");
      li.className = "tab-item";

      const num = document.createElement("span");
      num.className = "tab-num";
      num.textContent = String(i + 1);

      const text = document.createElement("div");
      text.className = "tab-text";
      const name = document.createElement("span");
      name.className = "tab-name";
      name.textContent = entry.name;
      const source = document.createElement("span");
      source.className = "source-meta";
      source.textContent = entry.sheet.isCsv || entry.sheet.sheetCount <= 1
        ? `from ${entry.sheet.fileLabel}`
        : `from ${entry.sheet.fileLabel} › ${entry.sheet.sheetName}`;
      text.append(name, source);

      const controls = document.createElement("div");
      controls.className = "tab-controls";
      controls.append(moveButton(entry, -1, i === 0), moveButton(entry, 1, i === entries.length - 1));

      li.append(num, text, controls);
      tabOrderEl.appendChild(li);
    });

    noTabsNote.classList.toggle("hidden", entries.length > 0);
    downloadBtn.disabled = entries.length === 0;
    const fileCount = new Set(entries.map((e) => e.sheet.fileLabel)).size;
    summaryChips.innerHTML = "";
    [[entries.length, "tab"], [fileCount, "file"]].forEach(([n, word]) => {
      const chip = document.createElement("span");
      chip.className = "chip";
      const strong = document.createElement("strong");
      strong.textContent = String(n);
      chip.append(strong, ` ${word}${n !== 1 ? "s" : ""}`);
      summaryChips.appendChild(chip);
    });
  }

  function render() {
    sheetsPanel.classList.toggle("hidden", sheets.length === 0);
    tabsPanel.classList.toggle("hidden", sheets.length === 0);
    SMS.renderSheetChecklist(sheetGroups, sheets, {
      mode: "tabs",
      onToggle: () => renderTabs(),
      onRemoveFile: (fileLabel) => {
        const removed = new Set(sheets.filter((s) => s.fileLabel === fileLabel).map((s) => s.id));
        sheets = sheets.filter((s) => !removed.has(s.id));
        order = order.filter((id) => !removed.has(id));
        render();
      },
    });
    renderTabs();
  }

  // ---------- Download ----------
  downloadBtn.addEventListener("click", () => {
    const entries = tabEntries();
    if (!entries.length) return;
    const wb = TM.buildTabbedWorkbook(entries);
    const blob = new Blob([TM.workbookToArray(wb)], { type: SMS.XLSX_MIME });
    SMS.downloadBlob(blob, `${SMS.cleanOutputName(outputNameInput.value, "merged")}.xlsx`);
    SMS.track("download_clicked", { tool: "merge", format: "xlsx" });
  });

  // ---------- Files handed over from Stack sheets ----------
  SMS.takeHandoff().then((files) => {
    if (files.length) ingestFiles(files, `Brought over ${SMS.plural(files.length, "file")} from Stack sheets.`);
  });
})();
