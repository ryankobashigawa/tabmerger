/*
 * Pure spreadsheet engine shared by all three tools — no DOM, no FileReader.
 * Takes already-loaded file content (array buffer for xlsx/xls, text for csv)
 * and:
 *   - reads it into a normalized workbook (values only, formats kept),
 *   - flattens each sheet into headers + rows for stacking (consolidate),
 *   - builds a one-tab-per-sheet workbook (merge into tabs),
 *   - builds one file per tab (split).
 *
 * Kept dependency-free of the browser so it can run under a plain JS engine
 * for tests (see test/run-tests.sh) and would run under Node unmodified.
 * Depends only on the SheetJS `XLSX` global, loaded before this script.
 */
(function (global) {
  "use strict";

  var ACCEPTED_EXTENSIONS = [".xlsx", ".xls", ".csv"];
  var WORKBOOK_EXTENSIONS = [".xlsx", ".xls"];

  // Excel serial-number offset between the 1904 and 1900 date systems.
  var DATE1904_OFFSET = 1462;

  // Excel tab-name rules: max 31 chars, none of \ / ? * [ ] :, no leading or
  // trailing apostrophe, unique (case-insensitive), "History" is reserved.
  var MAX_TAB_NAME = 31;
  var ILLEGAL_TAB_CHARS = /[\\\/\?\*\[\]:]/g;

  // Characters Windows/macOS/Linux don't allow (or handle badly) in file names.
  var ILLEGAL_FILE_CHARS = /[\\\/:\*\?"<>\|\x00-\x1f]/g;
  var WINDOWS_RESERVED = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i;
  var MAX_FILE_BASE = 100;

  function hasAcceptedExtension(name, extensions) {
    var lower = String(name || "").toLowerCase();
    return (extensions || ACCEPTED_EXTENSIONS).some(function (ext) {
      return lower.endsWith(ext);
    });
  }

  function stripExtension(name) {
    return String(name || "").replace(/\.(xlsx|xls|csv)$/i, "");
  }

  function normalizeHeaderKey(h) {
    return String(h === undefined || h === null ? "" : h).trim().toLowerCase();
  }

  // Disambiguates repeated file names within one upload batch:
  // ("report.xlsx", 1) -> "report.xlsx", ("report.xlsx", 2) -> "report.xlsx (2)"
  function buildSourceLabel(fileName, occurrenceIndex) {
    return occurrenceIndex <= 1 ? fileName : fileName + " (" + occurrenceIndex + ")";
  }

  /**
   * A number that carries its Excel number format, e.g. a date (serial 46037,
   * "yyyy-mm-dd") or a zero-padded ID (2134, "00000"). Rows hold these instead
   * of bare numbers so the format survives into the stacked output; toString()
   * gives the text Excel would display, which the preview and CSV use.
   */
  function FormattedNumber(v, z, w) {
    this.v = v;
    this.z = z;
    this.w = w;
  }
  FormattedNumber.prototype.toString = function () {
    return this.w;
  };

  // CSV text that is safe to turn into a number: no leading zeros ("00123"
  // stays text), no thousands separators, at most 15 significant digits
  // (beyond that, doubles lose precision — long IDs stay text).
  var PLAIN_NUMBER = /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/;
  function csvTextToValue(text) {
    if (!PLAIN_NUMBER.test(text)) return null;
    if (text.replace(/[-.]/g, "").replace(/^0+/, "").length > 15) return null;
    return Number(text);
  }

  function isDateFormat(z) {
    try {
      return !!z && XLSX.SSF.is_date(z);
    } catch (e) {
      return false;
    }
  }

  function cellKeys(ws) {
    return Object.keys(ws).filter(function (k) {
      return k.charAt(0) !== "!";
    });
  }

  /**
   * Read a file into a normalized SheetJS workbook:
   *   - formulas are dropped; each cell keeps its last calculated value,
   *   - number/date formats are kept (cell.z),
   *   - CSV cells are read as text, then plain numbers become numbers, so
   *     "00123" and "2026-01-15" stay exactly as written,
   *   - 1904-date-system workbooks are shifted to the 1900 system so dates
   *     stay correct when copied into a new workbook.
   * @param {ArrayBuffer|string} source - array buffer for xlsx/xls, raw text for csv
   */
  function readWorkbook(source, opts) {
    opts = opts || {};
    var isCsv = !!opts.isCsv;
    var wb = isCsv
      ? XLSX.read(source, { type: "string", raw: true })
      : XLSX.read(source, { type: "array", cellNF: true });

    var props = (wb.Workbook && wb.Workbook.WBProps) || {};
    var date1904 = !!props.date1904;

    wb.SheetNames.forEach(function (name) {
      var ws = wb.Sheets[name];
      if (!ws) return;
      cellKeys(ws).forEach(function (addr) {
        var cell = ws[addr];
        delete cell.f;
        delete cell.F;
        if (isCsv && cell.t === "s") {
          var num = csvTextToValue(cell.v);
          if (num !== null) {
            cell.t = "n";
            cell.v = num;
            delete cell.w;
          }
        }
        if (date1904 && cell.t === "n" && isDateFormat(cell.z)) {
          cell.v += DATE1904_OFFSET;
        }
      });
    });
    if (date1904) props.date1904 = false;
    return wb;
  }

  // Cell -> row value: "" for blanks, FormattedNumber for numbers with a
  // non-General format, the raw value otherwise.
  function cellToValue(cell) {
    if (!cell || cell.t === "z") return "";
    if (cell.t === "e") return cell.w !== undefined ? cell.w : "";
    if (cell.t === "n" && cell.z && cell.z !== "General") {
      var w = cell.w !== undefined ? cell.w : XLSX.SSF.format(cell.z, cell.v);
      return new FormattedNumber(cell.v, cell.z, w);
    }
    if (cell.t === "d") return cell.w !== undefined ? cell.w : String(cell.v);
    return cell.v === undefined || cell.v === null ? "" : cell.v;
  }

  function cellToText(cell) {
    if (!cell || cell.v === undefined || cell.v === null) return "";
    if (cell.t === "s") return String(cell.v);
    return cell.w !== undefined ? String(cell.w) : String(cell.v);
  }

  // Sheet -> array of rows of cell objects (positional, from the range start).
  function sheetCellGrid(ws) {
    if (!ws || !ws["!ref"]) return [];
    var range = XLSX.utils.decode_range(ws["!ref"]);
    var grid = [];
    for (var r = range.s.r; r <= range.e.r; r++) {
      var row = [];
      for (var c = range.s.c; c <= range.e.c; c++) {
        row.push(ws[XLSX.utils.encode_cell({ r: r, c: c })]);
      }
      grid.push(row);
    }
    return grid;
  }

  // Used-range size (including the header row); 0x0 for a blank sheet.
  function sheetSize(ws) {
    if (!ws || !ws["!ref"] || !sheetHasContent(ws)) return { rows: 0, cols: 0 };
    var range = XLSX.utils.decode_range(ws["!ref"]);
    return { rows: range.e.r - range.s.r + 1, cols: range.e.c - range.s.c + 1 };
  }

  function sheetHasContent(ws) {
    return cellKeys(ws).some(function (addr) {
      return cellToText(ws[addr]).trim() !== "";
    });
  }

  /**
   * Parse every sheet out of an already-loaded workbook source.
   * @param {ArrayBuffer|string} source - array buffer for xlsx/xls, raw text for csv
   * @param {{isCsv?:boolean, fileName:string, fileLabel:string}} opts
   * @returns {Array<Sheet>} Sheet = {fileName, fileLabel, sheetName, displayName,
   *   headers, rows, isEmpty, hasHeaderRow, hasContent, size, isCsv, sheetCount, worksheet}
   */
  function parseSheetsFromSource(source, opts) {
    opts = opts || {};
    var isCsv = !!opts.isCsv;
    var fileName = opts.fileName || "file";
    var fileLabel = opts.fileLabel || fileName;

    var workbook = readWorkbook(source, { isCsv: isCsv });

    var sheets = [];
    workbook.SheetNames.forEach(function (sheetName) {
      var ws = workbook.Sheets[sheetName];
      if (!ws) return;

      // Work positionally rather than trusting SheetJS's auto object keys, so
      // header matching is based on the *trimmed* header text we display.
      var grid = sheetCellGrid(ws);
      var trimmedHeaderRow = (grid[0] || []).map(function (cell) {
        return cellToText(cell).trim();
      });
      var headers = trimmedHeaderRow.filter(function (h) {
        return h !== "";
      });
      var dataRows = grid.slice(1).filter(function (cells) {
        return cells.some(function (cell) {
          return cellToText(cell).trim() !== "";
        });
      });

      var rows = dataRows.map(function (cells) {
        var obj = {};
        trimmedHeaderRow.forEach(function (h, idx) {
          if (h === "") return; // blank-header columns aren't addressable
          obj[h] = cellToValue(cells[idx]);
        });
        return obj;
      });

      sheets.push({
        fileName: fileName,
        fileLabel: fileLabel,
        sheetName: isCsv ? "" : sheetName,
        displayName: isCsv ? fileLabel : fileLabel + " — " + sheetName,
        headers: headers,
        rows: rows,
        isEmpty: rows.length === 0,
        hasHeaderRow: headers.length > 0,
        hasContent: sheetHasContent(ws),
        size: sheetSize(ws),
        isCsv: isCsv,
        sheetCount: workbook.SheetNames.length,
        worksheet: ws,
      });
    });
    return sheets;
  }

  /**
   * Compare headers (case-insensitive, trimmed) across a set of sheets.
   * @returns {{allMatch:boolean, commonColumns:string[], differingColumns:Array<{key,label,sheets,presentInAll}>, unionOrder:string[]}}
   */
  function compareHeaders(sheets) {
    var canonicalByKey = {}; // normalizedKey -> first-seen display label
    var order = []; // normalizedKey, in first-seen order
    var countByKey = {}; // normalizedKey -> # of sheets containing it
    var sheetsWithKey = {}; // normalizedKey -> [displayName, ...]

    sheets.forEach(function (sheet) {
      var seenInThisSheet = {};
      sheet.headers.forEach(function (h) {
        var key = normalizeHeaderKey(h);
        if (!key) return;
        if (!(key in canonicalByKey)) {
          canonicalByKey[key] = h;
          order.push(key);
          countByKey[key] = 0;
          sheetsWithKey[key] = [];
        }
        if (!seenInThisSheet[key]) {
          seenInThisSheet[key] = true;
          countByKey[key] += 1;
          sheetsWithKey[key].push(sheet.displayName);
        }
      });
    });

    var total = sheets.length;
    var common = [];
    var differing = [];
    order.forEach(function (key) {
      var entry = {
        key: key,
        label: canonicalByKey[key],
        sheets: sheetsWithKey[key],
        presentInAll: total > 0 && countByKey[key] === total,
      };
      (entry.presentInAll ? common : differing).push(entry);
    });

    return {
      allMatch: differing.length === 0,
      commonColumns: common.map(function (e) {
        return e.label;
      }),
      differingColumns: differing,
      unionOrder: order.map(function (key) {
        return canonicalByKey[key];
      }),
    };
  }

  /**
   * Consolidate a Sheet[] into a flat column/row set.
   * @param {Array<Sheet>} sheets
   * @param {{strategy?: 'union'|'intersection', includeSourceFile?: boolean, includeSourceSheet?: boolean}} options
   * @returns {{columns:string[], rows:object[], headerComparison:object}}
   */
  function consolidate(sheets, options) {
    options = options || {};
    var strategy = options.strategy === "intersection" ? "intersection" : "union";
    var includeSourceFile = !!options.includeSourceFile;
    var includeSourceSheet = !!options.includeSourceSheet;

    var cmp = compareHeaders(sheets);
    var dataColumns = strategy === "intersection" ? cmp.commonColumns : cmp.unionOrder;

    var columns = [];
    if (includeSourceFile) columns.push("source_file");
    if (includeSourceSheet) columns.push("source_sheet");
    dataColumns.forEach(function (c) {
      if (columns.indexOf(c) === -1) columns.push(c);
    });

    var rows = [];
    sheets.forEach(function (sheet) {
      var ownKeyToRowKey = {};
      sheet.headers.forEach(function (h) {
        ownKeyToRowKey[normalizeHeaderKey(h)] = h;
      });

      sheet.rows.forEach(function (row) {
        var outRow = {};
        if (includeSourceFile) outRow.source_file = sheet.fileLabel;
        if (includeSourceSheet) outRow.source_sheet = sheet.sheetName || "";
        dataColumns.forEach(function (label) {
          var rowKey = ownKeyToRowKey[normalizeHeaderKey(label)];
          outRow[label] =
            rowKey !== undefined && Object.prototype.hasOwnProperty.call(row, rowKey)
              ? row[rowKey]
              : "";
        });
        rows.push(outRow);
      });
    });

    return { columns: columns, rows: rows, headerComparison: cmp };
  }

  // Rows -> worksheet. FormattedNumbers become number cells with their format.
  function rowsToWorksheet(columns, rows) {
    var formats = [];
    var aoa = [columns.slice()];
    rows.forEach(function (row, r) {
      aoa.push(
        columns.map(function (col, c) {
          var value = row[col];
          if (value instanceof FormattedNumber) {
            formats.push({ r: r + 1, c: c, z: value.z });
            return value.v;
          }
          return value === "" || value === undefined ? null : value;
        })
      );
    });
    var ws = XLSX.utils.aoa_to_sheet(aoa);
    formats.forEach(function (f) {
      var cell = ws[XLSX.utils.encode_cell({ r: f.r, c: f.c })];
      if (cell) cell.z = f.z;
    });
    return ws;
  }

  function sheetToCsv(columns, rows) {
    return XLSX.utils.sheet_to_csv(rowsToWorksheet(columns, rows));
  }

  function buildWorkbook(columns, rows, sheetName) {
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, rowsToWorksheet(columns, rows), sheetName || "Consolidated");
    return wb;
  }

  // ---------- Merge into tabs / split ----------

  // Make `name` unique (case-insensitively) against `taken` by appending
  // " (2)", " (3)", … while keeping the result within `maxLen`.
  function dedupeName(name, taken, maxLen, tidy) {
    var candidate = name;
    var n = 1;
    while (taken[candidate.toLowerCase()]) {
      n += 1;
      var suffix = " (" + n + ")";
      candidate = tidy(name.slice(0, maxLen - suffix.length)) + suffix;
    }
    taken[candidate.toLowerCase()] = true;
    return candidate;
  }

  function tidyTabName(s) {
    return s.trim().replace(/^'+|'+$/g, "").trim();
  }

  function sanitizeTabName(s) {
    return tidyTabName(String(s || "").replace(/[\x00-\x1f]/g, "").replace(ILLEGAL_TAB_CHARS, "-"));
  }

  /**
   * Excel tab names for a list of source sheets, in order. Single-tab files
   * (and CSVs) use the file name; multi-tab files use "file - sheet". When too
   * long, the file part is shortened first so the sheet name stays readable.
   * @param {Array<{fileName:string, sheetName:string, sheetCount:number, isCsv?:boolean}>} items
   */
  function excelTabNames(items) {
    var taken = {};
    return items.map(function (item) {
      var filePart = sanitizeTabName(stripExtension(item.fileName));
      var name;
      if (item.isCsv || item.sheetCount <= 1) {
        name = filePart;
      } else {
        var sheetPart = sanitizeTabName(item.sheetName);
        var room = MAX_TAB_NAME - 3 - sheetPart.length;
        if (filePart.length > room) filePart = tidyTabName(filePart.slice(0, Math.max(room, 10)));
        name = filePart ? filePart + " - " + sheetPart : sheetPart;
      }
      name = tidyTabName(name.slice(0, MAX_TAB_NAME)) || "Sheet";
      if (name.toLowerCase() === "history") name = "History-";
      return dedupeName(name, taken, MAX_TAB_NAME, tidyTabName);
    });
  }

  function tidyFileBase(s) {
    return s.trim().replace(/^\.+/, "").replace(/[. ]+$/, "").trim();
  }

  /**
   * Safe, unique file names (with extension) for a list of tab names.
   */
  function safeFileNames(names, extension) {
    var taken = {};
    return names.map(function (name) {
      var base = tidyFileBase(String(name || "").replace(ILLEGAL_FILE_CHARS, "_").slice(0, MAX_FILE_BASE)) || "Sheet";
      if (WINDOWS_RESERVED.test(base)) base += "_";
      return dedupeName(base, taken, MAX_FILE_BASE, tidyFileBase) + extension;
    });
  }

  // Copy of a worksheet's values, formats and layout (not styles or formulas).
  function cloneWorksheet(ws) {
    var out = {};
    ["!ref", "!merges", "!cols", "!rows"].forEach(function (k) {
      if (ws[k] !== undefined) out[k] = JSON.parse(JSON.stringify(ws[k]));
    });
    cellKeys(ws).forEach(function (addr) {
      var cell = ws[addr];
      var copy = { t: cell.t, v: cell.v };
      if (cell.z !== undefined) copy.z = cell.z;
      if (cell.w !== undefined) copy.w = cell.w;
      if (cell.l !== undefined) copy.l = cell.l;
      out[addr] = copy;
    });
    return out;
  }

  /**
   * One workbook, one tab per entry, in the given order.
   * @param {Array<{name:string, worksheet:object}>} entries - names already valid + unique
   */
  function buildTabbedWorkbook(entries) {
    var wb = XLSX.utils.book_new();
    entries.forEach(function (entry) {
      XLSX.utils.book_append_sheet(wb, cloneWorksheet(entry.worksheet), entry.name);
    });
    return wb;
  }

  function workbookToArray(wb) {
    return XLSX.write(wb, { type: "array", bookType: "xlsx" });
  }

  /**
   * One output file per tab.
   * @param {Array<{sheetName:string, worksheet:object}>} entries
   * @param {'xlsx'|'csv'} format
   * @returns {Array<{fileName:string, data:ArrayBuffer|string}>}
   */
  function splitSheets(entries, format) {
    var isCsv = format === "csv";
    var names = safeFileNames(
      entries.map(function (e) {
        return e.sheetName;
      }),
      isCsv ? ".csv" : ".xlsx"
    );
    return entries.map(function (entry, i) {
      var data = isCsv
        ? XLSX.utils.sheet_to_csv(entry.worksheet)
        : workbookToArray(buildTabbedWorkbook([{ name: entry.sheetName, worksheet: entry.worksheet }]));
      return { fileName: names[i], data: data };
    });
  }

  global.TM = {
    ACCEPTED_EXTENSIONS: ACCEPTED_EXTENSIONS,
    WORKBOOK_EXTENSIONS: WORKBOOK_EXTENSIONS,
    hasAcceptedExtension: hasAcceptedExtension,
    stripExtension: stripExtension,
    normalizeHeaderKey: normalizeHeaderKey,
    buildSourceLabel: buildSourceLabel,
    FormattedNumber: FormattedNumber,
    readWorkbook: readWorkbook,
    parseSheetsFromSource: parseSheetsFromSource,
    compareHeaders: compareHeaders,
    consolidate: consolidate,
    sheetToCsv: sheetToCsv,
    buildWorkbook: buildWorkbook,
    excelTabNames: excelTabNames,
    safeFileNames: safeFileNames,
    cloneWorksheet: cloneWorksheet,
    buildTabbedWorkbook: buildTabbedWorkbook,
    workbookToArray: workbookToArray,
    splitSheets: splitSheets,
  };
})(typeof window !== "undefined" ? window : this);
