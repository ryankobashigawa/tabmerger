// Split page (/split): one workbook in, one file per selected tab out.
(() => {
  "use strict";

  /** @type {Array<Object>} core.js Sheets of the current workbook, plus {id, included} */
  let sheets = [];
  let workbookName = "";
  let format = "xlsx"; // 'xlsx' | 'csv'

  const sheetsPanel = document.getElementById("sheets-panel");
  const sheetGroups = document.getElementById("sheet-groups");
  const outputPanel = document.getElementById("output-panel");
  const fileNameList = document.getElementById("file-name-list");
  const noFilesNote = document.getElementById("no-files-note");
  const summaryChips = document.getElementById("summary-chips");
  const downloadBtn = document.getElementById("download-btn");
  const setStatus = SMS.statusSetter(document.getElementById("status-line"));

  // ---------- Loading (one workbook at a time) ----------
  async function loadWorkbook(fileList) {
    const files = Array.from(fileList);
    const workbooks = files.filter((f) => TM.hasAcceptedExtension(f.name, TM.WORKBOOK_EXTENSIONS));
    if (!workbooks.length) {
      const csvNote = files.some((f) => /\.csv$/i.test(f.name)) ? " A CSV file has only one sheet, so there's nothing to split." : "";
      setStatus(`Please choose an .xlsx or .xls workbook.${csvNote}`, "error");
      return;
    }
    const file = workbooks[0];
    setStatus(`Reading ${file.name}…`);
    const result = await SMS.loadFiles([file], { extensions: TM.WORKBOOK_EXTENSIONS, fileNameCounts: {} });
    if (result.errors.length) {
      setStatus(result.errors.join(" "), "error");
      return;
    }
    sheets = result.sheets.map((sheet, i) => Object.assign(sheet, { id: `tab-${i}`, included: sheet.hasContent }));
    workbookName = TM.stripExtension(file.name);
    const extra = files.length > 1 ? ` Only one workbook can be split at a time, so ${SMS.plural(files.length - 1, "other file")} ${files.length > 2 ? "were" : "was"} ignored.` : "";
    setStatus(`Loaded ${SMS.plural(sheets.length, "tab")} from ${file.name}.${extra}`, "good");
    render();
  }

  function reset() {
    sheets = [];
    workbookName = "";
    render();
  }

  SMS.wireDropzone({
    dropzone: document.getElementById("dropzone"),
    fileInput: document.getElementById("file-input"),
    chooseButton: document.getElementById("choose-files-btn"),
    onFiles: loadWorkbook,
  });

  const sampleBtn = document.getElementById("sample-btn");
  sampleBtn.addEventListener("click", async () => {
    sampleBtn.disabled = true;
    try {
      await loadWorkbook([SMS.splitSampleFile()]);
      sheetsPanel.scrollIntoView({ behavior: "smooth", block: "start" });
    } finally {
      sampleBtn.disabled = false;
    }
  });

  document.getElementById("clear-all-btn").addEventListener("click", () => {
    reset();
    setStatus("Cleared.");
  });

  document.querySelectorAll('input[name="format"]').forEach((radio) => {
    radio.addEventListener("change", () => {
      if (radio.checked) {
        format = radio.value;
        renderOutput();
      }
    });
  });

  // ---------- Rendering ----------
  function selected() {
    return sheets.filter((s) => s.included);
  }

  function outputFileNames() {
    return TM.safeFileNames(selected().map((s) => s.sheetName), `.${format}`);
  }

  function renderOutput() {
    const names = outputFileNames();
    fileNameList.innerHTML = "";
    names.forEach((name) => {
      const li = document.createElement("li");
      li.textContent = name;
      fileNameList.appendChild(li);
    });
    noFilesNote.classList.toggle("hidden", names.length > 0);
    downloadBtn.disabled = names.length === 0;
    downloadBtn.textContent = names.length === 1 ? `Download ${names[0]}` : "Download .zip";

    summaryChips.innerHTML = "";
    const chip = document.createElement("span");
    chip.className = "chip";
    const strong = document.createElement("strong");
    strong.textContent = String(names.length);
    chip.append(strong, ` file${names.length !== 1 ? "s" : ""}`);
    summaryChips.appendChild(chip);
  }

  function render() {
    sheetsPanel.classList.toggle("hidden", sheets.length === 0);
    outputPanel.classList.toggle("hidden", sheets.length === 0);
    SMS.renderSheetChecklist(sheetGroups, sheets, {
      mode: "tabs",
      removeLabel: "Remove this workbook",
      onToggle: () => renderOutput(),
      onRemoveFile: () => {
        reset();
        setStatus("Cleared.");
      },
    });
    renderOutput();
  }

  // ---------- Download ----------
  downloadBtn.addEventListener("click", async () => {
    const entries = selected().map((s) => ({ sheetName: s.sheetName, worksheet: s.worksheet }));
    if (!entries.length) return;
    downloadBtn.disabled = true;
    try {
      const files = TM.splitSheets(entries, format);
      const type = format === "csv" ? "text/csv;charset=utf-8;" : SMS.XLSX_MIME;
      if (files.length === 1) {
        SMS.downloadBlob(new Blob([files[0].data], { type }), files[0].fileName);
      } else {
        const zip = new JSZip();
        files.forEach((f) => zip.file(f.fileName, f.data));
        const blob = await zip.generateAsync({ type: "blob", compression: "DEFLATE" });
        const zipName = SMS.cleanOutputName(`${workbookName} - ${format === "csv" ? "csv" : "tabs"}`, "split");
        SMS.downloadBlob(blob, `${zipName}.zip`);
      }
      SMS.track("download_clicked", { tool: "split", format });
    } catch (err) {
      setStatus(`Couldn't create the files: ${err.message || err}`, "error");
    } finally {
      downloadBtn.disabled = selected().length === 0;
    }
  });
})();
