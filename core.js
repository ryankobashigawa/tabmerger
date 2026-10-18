/*
 * Pure consolidation engine — no DOM, no FileReader. Takes already-loaded
 * file content (array buffer / base64 / CSV text) and produces a flat
 * Sheet[], compares headers across sheets, and consolidates them.
 *
 * Kept dependency-free of the browser so it can run under a plain JS engine
 * for tests (see test/run-tests.sh) and would run under Node unmodified.
 * Depends only on the SheetJS `XLSX` global, loaded before this script.
 */
(function (global) {
  "use strict";

  var ACCEPTED_EXTENSIONS = [".xlsx", ".xls", ".csv"];

  function hasAcceptedExtension(name) {
    var lower = String(name || "").toLowerCase();
    return ACCEPTED_EXTENSIONS.some(function (ext) {
      return lower.endsWith(ext);
    });
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
   * Parse every sheet out of an already-loaded workbook source.
   * @param {ArrayBuffer|string} source - array buffer for xlsx/xls, raw text for csv
   * @param {{isCsv?:boolean, fileName:string, fileLabel:string}} opts
   * @returns {Array<Sheet>} Sheet = {fileName, fileLabel, sheetName, displayName,
   *   headers, rows, isEmpty, hasHeaderRow}
   */
  function parseSheetsFromSource(source, opts) {
    opts = opts || {};
    var isCsv = !!opts.isCsv;
    var fileName = opts.fileName || "file";
    var fileLabel = opts.fileLabel || fileName;

    var workbook = isCsv
      ? XLSX.read(source, { type: "string" })
      : XLSX.read(source, { type: "array" });

    var sheets = [];
    workbook.SheetNames.forEach(function (sheetName) {
      var ws = workbook.Sheets[sheetName];
      if (!ws) return;

      // Work positionally (array-of-arrays) rather than trusting SheetJS's
      // auto object keys, so header matching is based on the *trimmed*
      // header text we actually display — not whatever raw key SheetJS
      // generates internally (which can differ by stray whitespace).
      var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
      var rawHeaderRow = aoa[0] || [];
      var trimmedHeaderRow = rawHeaderRow.map(function (h) {
        return h === undefined || h === null ? "" : String(h).trim();
      });
      var headers = trimmedHeaderRow.filter(function (h) {
        return h !== "";
      });
      var dataRows = aoa.slice(1).filter(function (arr) {
        return arr.some(function (cell) {
          return cell !== undefined && cell !== null && String(cell).trim() !== "";
        });
      });

      var rows = dataRows.map(function (arr) {
        var obj = {};
        trimmedHeaderRow.forEach(function (h, idx) {
          if (h === "") return; // blank-header columns aren't addressable
          obj[h] = arr[idx] !== undefined ? arr[idx] : "";
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

  function sheetToCsv(columns, rows) {
    var ws = XLSX.utils.json_to_sheet(rows, { header: columns });
    return XLSX.utils.sheet_to_csv(ws);
  }

  function buildWorkbook(columns, rows, sheetName) {
    var ws = XLSX.utils.json_to_sheet(rows, { header: columns });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName || "Consolidated");
    return wb;
  }

  global.TM = {
    ACCEPTED_EXTENSIONS: ACCEPTED_EXTENSIONS,
    hasAcceptedExtension: hasAcceptedExtension,
    normalizeHeaderKey: normalizeHeaderKey,
    buildSourceLabel: buildSourceLabel,
    parseSheetsFromSource: parseSheetsFromSource,
    compareHeaders: compareHeaders,
    consolidate: consolidate,
    sheetToCsv: sheetToCsv,
    buildWorkbook: buildWorkbook,
  };
})(typeof window !== "undefined" ? window : this);
