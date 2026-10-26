/*
 * Browser glue shared by all three tools: file reading, drop zone + pickers,
 * folder traversal, the sheet checklist, downloads, analytics, sample files
 * and the stack -> merge hand-off. Pure spreadsheet logic lives in core.js.
 * Everything here runs locally; nothing is sent anywhere except the Umami
 * events in track(), which never include file names or contents.
 */
(function (global) {
  "use strict";

  const HEADER_PREVIEW_LIMIT = 8;

  function plural(n, word) {
    return `${n} ${word}${n !== 1 ? "s" : ""}`;
  }

  // ---------- Analytics ----------
  // Umami custom events. Never pass file names, sheet names or cell data.
  // No-ops when the Umami script is blocked or hasn't loaded yet.
  function track(eventName, data) {
    try {
      if (global.umami && typeof global.umami.track === "function") {
        global.umami.track(eventName, data);
      }
    } catch (err) {
      // Analytics must never break the app.
    }
  }

  // ---------- Status line ----------
  function statusSetter(el) {
    return function setStatus(msg, kind) {
      el.textContent = msg || "";
      el.classList.remove("error", "good");
      if (kind) el.classList.add(kind);
    };
  }

  // ---------- File reading ----------
  function readFile(file, asText) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result);
      reader.onerror = () => reject(reader.error);
      if (asText) reader.readAsText(file);
      else reader.readAsArrayBuffer(file);
    });
  }

  /**
   * Read files into core.js Sheets.
   * @param {Iterable<File>} fileList
   * @param {{extensions?: string[], fileNameCounts: Object}} opts
   * @returns {Promise<{sheets, loaded: Array<{file, fileLabel}>, rejected: File[], errors: string[]}>}
   */
  async function loadFiles(fileList, opts) {
    const extensions = opts.extensions || TM.ACCEPTED_EXTENSIONS;
    const counts = opts.fileNameCounts;
    const files = Array.from(fileList);
    const accepted = files.filter((f) => TM.hasAcceptedExtension(f.name, extensions));
    const rejected = files.filter((f) => !TM.hasAcceptedExtension(f.name, extensions));
    const sheets = [];
    const loaded = [];
    const errors = [];

    for (const file of accepted) {
      try {
        counts[file.name] = (counts[file.name] || 0) + 1;
        const fileLabel = TM.buildSourceLabel(file.name, counts[file.name]);
        const isCsv = file.name.toLowerCase().endsWith(".csv");
        const source = await readFile(file, isCsv);
        const parsed = TM.parseSheetsFromSource(source, { isCsv, fileName: file.name, fileLabel });
        parsed.forEach((sheet) => sheets.push(sheet));
        loaded.push({ file, fileLabel });
      } catch (err) {
        errors.push(`Couldn't read "${file.name}": ${err.message || err}`);
      }
    }
    return { sheets, loaded, rejected, errors };
  }

  // Status message + kind for a loadFiles() result.
  function loadSummary(result, extensionsLabel) {
    const { sheets, loaded, rejected, errors } = result;
    const parts = [];
    if (!loaded.length && !errors.length) {
      const suffix = rejected.length ? ` Skipped: ${rejected.map((f) => f.name).join(", ")}.` : "";
      return { message: `No ${extensionsLabel || ".xlsx/.xls/.csv"} files found in that selection.${suffix}`, kind: "error" };
    }
    if (sheets.length) {
      parts.push(`Loaded ${plural(sheets.length, "sheet")} from ${plural(loaded.length, "file")}.`);
    }
    if (rejected.length) {
      parts.push(`Skipped unsupported file${rejected.length !== 1 ? "s" : ""}: ${rejected.map((f) => f.name).join(", ")}.`);
    }
    if (errors.length) parts.push(errors.join(" "));
    return { message: parts.join(" "), kind: errors.length ? "error" : "good" };
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

  // ---------- Drop zone + pickers ----------
  /**
   * @param {{dropzone, fileInput, folderInput?, chooseButton, folderButton?, onFiles}} els
   */
  function wireDropzone(els) {
    const { dropzone, fileInput, folderInput, chooseButton, folderButton, onFiles } = els;

    // Clicking anywhere in the zone opens the picker; the buttons inside
    // handle their own clicks and stop propagation so it only opens once.
    dropzone.addEventListener("click", () => fileInput.click());
    dropzone.addEventListener("keydown", (e) => {
      if (e.target !== dropzone) return; // let inner buttons handle their own keys
      if (e.key === "Enter" || e.key === " ") {
        e.preventDefault();
        fileInput.click();
      }
    });
    ["dragenter", "dragover"].forEach((evt) =>
      dropzone.addEventListener(evt, (e) => {
        e.preventDefault();
        dropzone.classList.add("dragover");
      })
    );
    dropzone.addEventListener("dragleave", (e) => {
      // dragleave also fires when moving onto a child element; ignore those.
      if (!dropzone.contains(e.relatedTarget)) dropzone.classList.remove("dragover");
    });
    dropzone.addEventListener("dragend", () => dropzone.classList.remove("dragover"));
    dropzone.addEventListener("drop", async (e) => {
      e.preventDefault();
      dropzone.classList.remove("dragover");
      onFiles(await filesFromDataTransfer(e.dataTransfer));
    });

    const inputs = [fileInput, folderInput].filter(Boolean);
    inputs.forEach((input) => {
      input.addEventListener("change", (e) => {
        if (e.target.files.length) onFiles(Array.from(e.target.files));
        e.target.value = "";
      });
      // Hidden inputs sit inside the zone, so their synthetic clicks would
      // bubble back up to the zone's click handler and reopen the picker.
      input.addEventListener("click", (e) => e.stopPropagation());
    });
    chooseButton.addEventListener("click", (e) => {
      e.stopPropagation();
      fileInput.click();
    });
    if (folderButton && folderInput) {
      folderButton.addEventListener("click", (e) => {
        e.stopPropagation();
        folderInput.click();
      });
    }
  }

  // ---------- Sheet checklist (grouped by source file) ----------
  function truncatedHeaderPreview(headers) {
    if (headers.length === 0) return "—";
    const shown = headers.slice(0, HEADER_PREVIEW_LIMIT).join(", ");
    const extra = headers.length - HEADER_PREVIEW_LIMIT;
    return extra > 0 ? `${shown}, +${extra} more` : shown;
  }

  /**
   * Render sheets (each has {id, included, fileLabel, sheetName, rows, headers,
   * isEmpty, hasHeaderRow}) as a checklist grouped by file.
   * @param {HTMLElement} container - the <ul>
   * @param {{onToggle(sheet), onRemoveFile(fileLabel), mode?: 'stack'|'tabs', removeLabel?: string}} opts
   */
  function renderSheetChecklist(container, sheets, opts) {
    container.innerHTML = "";
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
      metaEl.textContent = plural(group.items.length, "sheet");

      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.className = "remove-btn";
      removeBtn.textContent = "✕";
      removeBtn.title = opts.removeLabel || "Remove this file";
      removeBtn.setAttribute("aria-label", `${opts.removeLabel || "Remove this file"}: ${group.fileLabel}`);
      removeBtn.addEventListener("click", () => opts.onRemoveFile(group.fileLabel));

      header.append(nameEl, metaEl, removeBtn);
      wrapper.appendChild(header);

      const list = document.createElement("ul");
      list.className = "sheet-sublist";

      group.items.forEach((sheet) => {
        const li = document.createElement("li");
        li.className = "sheet-row";

        const label = document.createElement("label");
        label.className = "sheet-check";
        const checkbox = document.createElement("input");
        checkbox.type = "checkbox";
        checkbox.checked = sheet.included;
        checkbox.addEventListener("change", () => {
          sheet.included = checkbox.checked;
          opts.onToggle(sheet);
        });
        const nameSpan = document.createElement("span");
        nameSpan.className = "sheet-label";
        nameSpan.textContent = sheet.sheetName || "(CSV)";
        label.append(checkbox, nameSpan);

        // Stacking needs a header row and data rows; whole-tab tools only
        // care whether the tab has anything in it.
        const wholeTabs = opts.mode === "tabs";
        let badgeText = "";
        if (wholeTabs) {
          if (!sheet.hasContent) badgeText = "empty";
        } else if (sheet.isEmpty) {
          badgeText = "empty";
        } else if (!sheet.hasHeaderRow) {
          badgeText = "no header row";
        }

        const meta = document.createElement("span");
        meta.className = "source-meta";
        meta.textContent = wholeTabs
          ? `${plural(sheet.size.rows, "row")} · ${plural(sheet.size.cols, "col")}`
          : `${plural(sheet.rows.length, "row")} · ${plural(sheet.headers.length, "col")}`;

        const headerPreview = document.createElement("span");
        headerPreview.className = "header-preview";
        headerPreview.textContent = truncatedHeaderPreview(sheet.headers);
        headerPreview.title = sheet.headers.join(", ");

        li.append(label);
        if (badgeText) {
          const badge = document.createElement("span");
          badge.className = "badge badge-warn";
          badge.textContent = badgeText;
          li.appendChild(badge);
        }
        li.append(meta, headerPreview);
        list.appendChild(li);
      });

      wrapper.appendChild(list);
      container.appendChild(wrapper);
    });
  }

  // ---------- Downloads ----------
  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

  function cleanOutputName(raw, fallback) {
    const name = String(raw || "").trim().replace(/\.(xlsx|xls|csv|zip)$/i, "");
    return TM.safeFileNames([name || fallback], "")[0];
  }

  // ---------- Sample files ----------
  // Built in code with the bundled SheetJS so the demos need no extra assets.
  function sampleWorkbookFile(fileName, tabs) {
    const wb = XLSX.utils.book_new();
    tabs.forEach(([tabName, aoa]) => {
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), tabName);
    });
    return new File([XLSX.write(wb, { type: "array", bookType: "xlsx" })], fileName, { type: XLSX_MIME });
  }

  // Includes a tab with a different column order (May) and a CSV with an
  // extra column (Notes) so the column handling is visible in the demo.
  function salesSampleFiles() {
    const header = ["Region", "Rep", "Units", "Revenue"];
    const q1 = sampleWorkbookFile("sales_q1.xlsx", [
      ["January", [header, ["North", "Avery", 42, 8400], ["South", "Jordan", 35, 7000], ["West", "Riley", 28, 5600]]],
      ["February", [header, ["North", "Avery", 38, 7600], ["South", "Jordan", 41, 8200], ["West", "Riley", 31, 6200]]],
      ["March", [header, ["North", "Avery", 47, 9400], ["South", "Jordan", 39, 7800], ["West", "Riley", 36, 7200]]],
    ]);
    const q2 = sampleWorkbookFile("sales_q2.xlsx", [
      ["April", [header, ["North", "Avery", 44, 8800], ["South", "Jordan", 37, 7400], ["West", "Riley", 33, 6600]]],
      // Same columns, different order: matched by name, not position.
      ["May", [["Revenue", "Region", "Units", "Rep"], [9000, "North", 45, "Avery"], [8000, "South", 40, "Jordan"], [7000, "West", 35, "Riley"]]],
    ]);
    const juneCsv = [
      "Region,Rep,Units,Revenue,Notes",
      "North,Avery,50,10000,Record month",
      "South,Jordan,43,8600,",
      "West,Riley,38,7600,New client signed",
    ].join("\n");
    const june = new File([juneCsv], "june_sales.csv", { type: "text/csv" });
    return [q1, q2, june];
  }

  function splitSampleFile() {
    const header = ["Item", "SKU", "In stock"];
    return sampleWorkbookFile("store_inventory.xlsx", [
      ["Downtown", [header, ["Notebook", "00412", 120], ["Pen set", "00977", 64]]],
      ["Airport", [header, ["Notebook", "00412", 35], ["Travel mug", "01203", 18]]],
      ["Online", [header, ["Notebook", "00412", 410], ["Pen set", "00977", 233], ["Travel mug", "01203", 97]]],
    ]);
  }

  // ---------- Stack -> merge hand-off ----------
  // Files are kept in this browser's IndexedDB (on this device only) just long
  // enough to open the other page, which takes them and deletes them.
  const HANDOFF_DB = "stackmysheets";
  const HANDOFF_STORE = "handoff";
  const HANDOFF_KEY = "files";
  const HANDOFF_MAX_AGE_MS = 10 * 60 * 1000;

  function openHandoffDb() {
    return new Promise((resolve, reject) => {
      const req = indexedDB.open(HANDOFF_DB, 1);
      req.onupgradeneeded = () => req.result.createObjectStore(HANDOFF_STORE);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error);
    });
  }

  function handoffRequest(mode, fn) {
    return openHandoffDb().then(
      (db) =>
        new Promise((resolve, reject) => {
          const tx = db.transaction(HANDOFF_STORE, mode);
          const result = fn(tx.objectStore(HANDOFF_STORE));
          tx.oncomplete = () => {
            db.close();
            resolve(result && "result" in result ? result.result : undefined);
          };
          tx.onerror = () => {
            db.close();
            reject(tx.error);
          };
        })
    );
  }

  function saveHandoff(files) {
    return handoffRequest("readwrite", (store) => store.put({ savedAt: Date.now(), files }, HANDOFF_KEY));
  }

  async function takeHandoff() {
    try {
      let record;
      const timeout = new Promise((_, reject) => setTimeout(() => reject(new Error("timeout")), 3000));
      await Promise.race([timeout, handoffRequest("readwrite", (store) => {
        const req = store.get(HANDOFF_KEY);
        req.onsuccess = () => {
          record = req.result;
          store.delete(HANDOFF_KEY);
        };
        return null;
      })]);
      if (!record || Date.now() - record.savedAt > HANDOFF_MAX_AGE_MS) return [];
      return record.files || [];
    } catch (err) {
      return []; // IndexedDB unavailable (e.g. some private modes): start empty
    }
  }

  global.SMS = {
    plural,
    track,
    statusSetter,
    loadFiles,
    loadSummary,
    wireDropzone,
    renderSheetChecklist,
    downloadBlob,
    XLSX_MIME,
    cleanOutputName,
    salesSampleFiles,
    splitSampleFile,
    saveHandoff,
    takeHandoff,
  };
})(window);
