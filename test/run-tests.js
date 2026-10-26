// Test runner for core.js (the pure, DOM-free consolidation engine).
//
// There's no Node.js on this dev machine, so this runs under macOS's
// built-in JavaScriptCore via `osascript -l JavaScript` (see run-tests.sh).
// core.js has no DOM/File API dependency, so it would run unmodified
// under Node (or any JS engine) too — only this runner is macOS-specific.
function run() {
  ObjC.import("Foundation");
  var app = Application.currentApplication();
  app.includeStandardAdditions = true;

  var REPO = "/Users/mockasingo/Documents/tabmerger";

  // Read as UTF-8 (app.read() assumes MacRoman and mangles non-ASCII text).
  function readText(path) {
    return $.NSString.stringWithContentsOfFileEncodingError(path, $.NSUTF8StringEncoding, $()).js;
  }

  function b64OfFile(path) {
    var nsdata = $.NSData.dataWithContentsOfFile(path);
    return ObjC.unwrap(nsdata.base64EncodedStringWithOptions(0));
  }

  eval(readText(REPO + "/vendor/xlsx.full.min.js"));
  eval(readText(REPO + "/core.js"));

  var results = [];
  var pass = 0;
  var fail = 0;

  function assert(name, condition, detail) {
    if (condition) {
      pass++;
      results.push("PASS - " + name);
    } else {
      fail++;
      results.push("FAIL - " + name + (detail ? " :: " + detail : ""));
    }
  }

  function assertEqual(name, actual, expected) {
    var ok = JSON.stringify(actual) === JSON.stringify(expected);
    assert(name, ok, ok ? "" : "expected " + JSON.stringify(expected) + " got " + JSON.stringify(actual));
  }

  function workbookSheetsFromFile(path, fileName, fileLabel) {
    // core.js expects an ArrayBuffer for non-CSV sources (mirroring what
    // FileReader.readAsArrayBuffer would hand it in the browser).
    var b64 = b64OfFile(path);
    var bin = atobPolyfill(b64);
    var buf = new ArrayBuffer(bin.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i < bin.length; i++) view[i] = bin.charCodeAt(i);
    return TM.parseSheetsFromSource(buf, { isCsv: false, fileName: fileName, fileLabel: fileLabel || fileName });
  }

  // Minimal atob polyfill (JavaScriptCore has no global atob).
  function atobPolyfill(b64) {
    var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
    var str = b64.replace(/[^A-Za-z0-9\+\/]/g, "");
    var output = "";
    for (var i = 0; i < str.length; i += 4) {
      var e1 = chars.indexOf(str.charAt(i));
      var e2 = chars.indexOf(str.charAt(i + 1));
      var e3 = chars.indexOf(str.charAt(i + 2));
      var e4 = chars.indexOf(str.charAt(i + 3));
      var c1 = (e1 << 2) | (e2 >> 4);
      var c2 = ((e2 & 15) << 4) | (e3 >> 2);
      var c3 = ((e3 & 3) << 6) | e4;
      output += String.fromCharCode(c1);
      if (e3 !== -1) output += String.fromCharCode(c2);
      if (e4 !== -1) output += String.fromCharCode(c3);
    }
    return output;
  }

  function csvSheets(csvText, fileName) {
    return TM.parseSheetsFromSource(csvText, { isCsv: true, fileName: fileName, fileLabel: fileName });
  }

  // ---------------------------------------------------------------------
  // 1. One multi-tab workbook, identical headers across tabs
  // ---------------------------------------------------------------------
  var quarterly = workbookSheetsFromFile(REPO + "/test-fixtures/quarterly_report.xlsx", "quarterly_report.xlsx");
  assertEqual("quarterly_report.xlsx has 3 tabs", quarterly.length, 3);
  assertEqual("quarterly_report.xlsx tab row counts", quarterly.map(function (s) { return s.rows.length; }), [2, 2, 1]);

  var qMerged = TM.consolidate(quarterly, { strategy: "union", includeSourceFile: false, includeSourceSheet: false });
  assertEqual("single workbook union row count", qMerged.rows.length, 5);
  assertEqual("single workbook union columns", qMerged.columns, ["Client", "Invoice #", "Amount", "Status"]);
  assert("single workbook: all headers matched (no diffs)", TM.compareHeaders(quarterly).allMatch);

  // ---------------------------------------------------------------------
  // 2. Multiple CSVs (same headers)
  // ---------------------------------------------------------------------
  var csvA = csvSheets("Client,Invoice #,Amount,Status\nAcme,INV-1,10,Paid\n", "a.csv");
  var csvB = csvSheets("Client,Invoice #,Amount,Status\nBeta,INV-2,20,Pending\n", "b.csv");
  var csvMerged = TM.consolidate(csvA.concat(csvB), { strategy: "union", includeSourceFile: true, includeSourceSheet: true });
  assertEqual("multiple CSVs: row count", csvMerged.rows.length, 2);
  assertEqual("multiple CSVs: source_file values", csvMerged.rows.map(function (r) { return r.source_file; }), ["a.csv", "b.csv"]);
  assertEqual("multiple CSVs: source_sheet blank for csv", csvMerged.rows.every(function (r) { return r.source_sheet === ""; }), true);

  // ---------------------------------------------------------------------
  // 3. Multiple multi-tab workbooks + mismatched headers under both strategies
  // ---------------------------------------------------------------------
  var april = workbookSheetsFromFile(REPO + "/test-fixtures/multi_month_folder/april.xlsx", "april.xlsx");
  var may = workbookSheetsFromFile(REPO + "/test-fixtures/multi_month_folder/may.xlsx", "may.xlsx");
  assertEqual("april.xlsx headers", april[0].headers, ["Client", "Invoice #", "Amount", "Status"]);
  assertEqual("may.xlsx headers (extra Notes column)", may[0].headers, ["Client", "Invoice #", "Amount", "Status", "Notes"]);

  var combo = april.concat(may);
  var cmp = TM.compareHeaders(combo);
  assert("mismatched headers: allMatch is false", !cmp.allMatch);
  assertEqual("mismatched headers: common columns", cmp.commonColumns, ["Client", "Invoice #", "Amount", "Status"]);
  assertEqual("mismatched headers: differing columns", cmp.differingColumns.map(function (d) { return d.label; }), ["Notes"]);

  var unionResult = TM.consolidate(combo, { strategy: "union", includeSourceFile: false, includeSourceSheet: false });
  assertEqual("union strategy: columns include Notes", unionResult.columns, ["Client", "Invoice #", "Amount", "Status", "Notes"]);
  var aprilRowInUnion = unionResult.rows[0];
  assertEqual("union strategy: blank-fills missing Notes", aprilRowInUnion.Notes, "");

  var intersectionResult = TM.consolidate(combo, { strategy: "intersection", includeSourceFile: false, includeSourceSheet: false });
  assertEqual("intersection strategy: drops Notes", intersectionResult.columns, ["Client", "Invoice #", "Amount", "Status"]);
  assertEqual("intersection strategy: still has both rows", intersectionResult.rows.length, 2);

  // ---------------------------------------------------------------------
  // 4. Source column toggles
  // ---------------------------------------------------------------------
  var noSourceCols = TM.consolidate(combo, { strategy: "union", includeSourceFile: false, includeSourceSheet: false });
  assert("source columns off: no source_file key", !("source_file" in noSourceCols.rows[0]));
  assert("source columns off: no source_sheet key", !("source_sheet" in noSourceCols.rows[0]));

  var withSourceCols = TM.consolidate(combo, { strategy: "union", includeSourceFile: true, includeSourceSheet: true });
  assertEqual("source columns on: appear first", withSourceCols.columns.slice(0, 2), ["source_file", "source_sheet"]);
  assertEqual("source columns on: correct values", [withSourceCols.rows[0].source_file, withSourceCols.rows[0].source_sheet], ["april.xlsx", "Sheet1"]);

  // ---------------------------------------------------------------------
  // 5. Duplicate file name disambiguation
  // ---------------------------------------------------------------------
  assertEqual("buildSourceLabel: first occurrence", TM.buildSourceLabel("report.xlsx", 1), "report.xlsx");
  assertEqual("buildSourceLabel: second occurrence", TM.buildSourceLabel("report.xlsx", 2), "report.xlsx (2)");

  // ---------------------------------------------------------------------
  // 6. Empty sheet / no-header-row detection
  // ---------------------------------------------------------------------
  var emptySheet = csvSheets("Client,Invoice #,Amount,Status\n", "empty.csv")[0];
  assertEqual("empty sheet: isEmpty true", emptySheet.isEmpty, true);
  assertEqual("empty sheet: still has headers", emptySheet.hasHeaderRow, true);

  var noHeaderSheet = csvSheets("\n\n", "blank.csv")[0];
  assert("no-header sheet: hasHeaderRow false", !noHeaderSheet.hasHeaderRow);

  // ---------------------------------------------------------------------
  // 7. Case-insensitive, trimmed header matching
  // ---------------------------------------------------------------------
  var caseA = csvSheets("Client, Amount\nAcme,10\n", "caseA.csv");
  var caseB = csvSheets("CLIENT,amount \nBeta,20\n", "caseB.csv");
  var caseCmp = TM.compareHeaders(caseA.concat(caseB));
  assert("case-insensitive headers: treated as matching", caseCmp.allMatch, JSON.stringify(caseCmp.differingColumns));
  var caseMerged = TM.consolidate(caseA.concat(caseB), { strategy: "union", includeSourceFile: false, includeSourceSheet: false });
  assertEqual("case-insensitive headers: single canonical column set", caseMerged.columns.length, 2);
  assertEqual("case-insensitive headers: canonical label from first-seen sheet", caseMerged.columns, ["Client", "Amount"]);
  assertEqual("case-insensitive headers: second sheet values land under canonical column", caseMerged.rows[1].Amount, 20);

  // ---------------------------------------------------------------------
  // 8. Values: CSV leading zeros/dates/long IDs, formats, formulas, 1904 dates
  // ---------------------------------------------------------------------
  var T = REPO + "/test-fixtures/tools/";
  function fixtureBuffer(path) {
    var bin = atobPolyfill(b64OfFile(path));
    var buf = new ArrayBuffer(bin.length);
    var view = new Uint8Array(buf);
    for (var i = 0; i < bin.length; i++) view[i] = bin.charCodeAt(i);
    return buf;
  }
  function fixtureSheets(name) {
    if (/\.csv$/.test(name)) return csvSheets(readText(T + name), name.split("/").pop());
    var fileName = name.split("/").pop();
    return TM.parseSheetsFromSource(fixtureBuffer(T + name), { fileName: fileName, fileLabel: fileName });
  }
  function str(v) { return v === undefined || v === null ? "" : String(v); }

  var csvRow = fixtureSheets("customers export.csv")[0].rows[0];
  assertEqual("csv: leading-zero ID stays text", csvRow["Customer ID"], "00123");
  assertEqual("csv: ISO date kept as written", csvRow["Signup date"], "2026-01-15");
  assertEqual("csv: US date kept as written", csvRow["US date"], "01/02/2026");
  assertEqual("csv: plain number becomes number", csvRow.Amount, 1234.5);
  assertEqual("csv: 20-digit ID stays text (no precision loss)", csvRow["Card ref"], "12345678901234567890");
  assertEqual("csv: thousands separator stays text", csvRow.Formatted, "1,234");
  assertEqual("csv: accented text intact", csvRow.City, "Café");
  assertEqual("csv: lone 0 becomes number", fixtureSheets("customers export.csv")[0].rows[2]["Customer ID"], 0);

  var north = fixtureSheets("region_north.xlsx")[0];
  var nRow = north.rows[0];
  assertEqual("xlsx: text ID keeps zeros", nRow["Customer ID"], "00121");
  assert("xlsx: zero-padded number keeps format", nRow.Zip instanceof TM.FormattedNumber && nRow.Zip.v === 2134 && str(nRow.Zip) === "02134", JSON.stringify(nRow.Zip));
  assert("xlsx: date keeps serial + format", nRow["Order date"] instanceof TM.FormattedNumber && nRow["Order date"].v === 46037 && str(nRow["Order date"]) === "2026-01-15", JSON.stringify(nRow["Order date"]));
  assertEqual("xlsx: accented text intact", nRow.Note, "Café");

  var stacked = TM.consolidate([north].concat(fixtureSheets("region_south.xlsx")), { strategy: "union" });
  var outWs = TM.buildWorkbook(stacked.columns, stacked.rows).Sheets.Consolidated;
  assertEqual("stack output: date cell is a formatted number", [outWs.C2.t, outWs.C2.v, outWs.C2.z], ["n", 46037, "yyyy-mm-dd"]);
  assertEqual("stack output: zip keeps 00000 format", [outWs.B2.v, outWs.B2.z], [2134, "00000"]);
  assertEqual("stack output: second file's m/d/yyyy date keeps its format", [outWs.C5.v, outWs.C5.z], [46037, "m/d/yyyy"]);
  var stackedCsv = TM.sheetToCsv(stacked.columns, stacked.rows);
  assert("stack csv: shows dates and zero-padded zips as displayed", stackedCsv.indexOf("00121,02134,2026-01-15") !== -1 && stackedCsv.indexOf("1/15/2026") !== -1, stackedCsv.split("\n")[1]);

  var west = fixtureSheets("region_west.xlsx")[0];
  assertEqual("formula: cached value kept", west.rows[0].Double, 2469);
  assert("formula: formula dropped", west.worksheet.G2.f === undefined);

  var mac = fixtureSheets("mac_1904_dates.xlsx")[0];
  assert("1904 dates: shifted to 1900 system", mac.rows[0].When.v === 46037 && str(mac.rows[0].When) === "2026-01-15", JSON.stringify(mac.rows[0].When));

  // ---------------------------------------------------------------------
  // 9. Excel tab names
  // ---------------------------------------------------------------------
  function tab(fileName, sheetName, sheetCount, isCsv) {
    return { fileName: fileName, sheetName: sheetName, sheetCount: sheetCount, isCsv: !!isCsv };
  }
  var names = TM.excelTabNames([
    tab("region_north.xlsx", "Data", 1),
    tab("customers export.csv", "Sheet1", 1, true),
    tab("sales.xlsx", "January", 3),
    tab("quarterly_regional_sales_report_2026.xlsx", "January", 6),
    tab("quarterly_regional_sales_report_2026.xlsx", "Summary for leadership review", 6),
    tab("Budget [draft] v2.xlsx", "Sheet1", 1),
    tab("a:b/c\\d?e*f.xlsx", "x", 1),
    tab("report.xlsx", "Report", 1),
    tab("REPORT.xlsx", "Report", 1),
    tab("'quoted'.xlsx", "x", 1),
    tab("history.xlsx", "x", 1),
    tab(".xlsx", "x", 1),
  ]);
  assertEqual("tab names", names, [
    "region_north",
    "customers export",
    "sales - January",
    "quarterly_regional_sa - January",
    "quarterly_ - Summary for leader",
    "Budget -draft- v2",
    "a-b-c-d-e-f",
    "report",
    "REPORT (2)",
    "quoted",
    "History-",
    "Sheet",
  ]);
  assert("tab names: all within 31 chars", names.every(function (n) { return n.length <= 31; }));

  var longDupes = TM.excelTabNames([
    tab("abcdefghijklmnopqrstuvwxyz0123456789.xlsx", "x", 1),
    tab("abcdefghijklmnopqrstuvwxyz0123456789.xlsx", "x", 1),
    tab("abcdefghijklmnopqrstuvwxyz0123456789.xlsx", "x", 1),
  ]);
  assertEqual("tab names: long duplicates truncated to fit suffix", longDupes, [
    "abcdefghijklmnopqrstuvwxyz01234",
    "abcdefghijklmnopqrstuvwxyz0 (2)",
    "abcdefghijklmnopqrstuvwxyz0 (3)",
  ]);

  // ---------------------------------------------------------------------
  // 10. Split file names
  // ---------------------------------------------------------------------
  assertEqual("file names", TM.safeFileNames(['Q1 "Final" <draft>|v2', "Data", "data", "CON", "Ends with dot.", "", "Café São Paulo"], ".xlsx"), [
    "Q1 _Final_ _draft__v2.xlsx",
    "Data.xlsx",
    "data (2).xlsx",
    "CON_.xlsx",
    "Ends with dot.xlsx",
    "Sheet.xlsx",
    "Café São Paulo.xlsx",
  ]);

  // ---------------------------------------------------------------------
  // 11. Round trip: merge every fixture into tabs, split back out, compare
  // ---------------------------------------------------------------------
  function sameCells(a, b) {
    var problems = [];
    if ((a["!ref"] || "") !== (b["!ref"] || "")) problems.push("ref " + a["!ref"] + " vs " + b["!ref"]);
    Object.keys(a).concat(Object.keys(b)).forEach(function (k) {
      if (k.charAt(0) === "!") return;
      var x = a[k], y = b[k];
      if (!x || !y) { problems.push(k + " missing"); return; }
      if (x.t !== y.t || x.v !== y.v) problems.push(k + " " + JSON.stringify([x.t, x.v]) + " vs " + JSON.stringify([y.t, y.v]));
      else if (x.t === "n" && x.z && x.z !== "General" && x.z !== y.z) problems.push(k + " format " + x.z + " vs " + y.z);
    });
    return problems;
  }
  var fixtureNames = ["region_north.xlsx", "region_south.xlsx", "region_west.xlsx", "Budget [draft] v2.xlsx",
    "dupes/a/report.xlsx", "dupes/b/report.xlsx", "quarterly_regional_sales_report_2026.xlsx",
    "mac_1904_dates.xlsx", "customers export.csv"];
  var originals = [];
  fixtureNames.forEach(function (n) { originals = originals.concat(fixtureSheets(n)); });
  var tabNames = TM.excelTabNames(originals);
  var merged = TM.buildTabbedWorkbook(originals.map(function (s, i) { return { name: tabNames[i], worksheet: s.worksheet }; }));
  var reread = TM.readWorkbook(TM.workbookToArray(merged));
  assertEqual("round trip: merged tab names in order", reread.SheetNames, tabNames);
  var mergeProblems = [];
  originals.forEach(function (s, i) {
    sameCells(s.worksheet, reread.Sheets[tabNames[i]]).forEach(function (p) { mergeProblems.push(tabNames[i] + ": " + p); });
  });
  assertEqual("round trip: merged workbook matches originals cell for cell", mergeProblems, []);

  var splitFiles = TM.splitSheets(reread.SheetNames.map(function (n) { return { sheetName: n, worksheet: reread.Sheets[n] }; }), "xlsx");
  assertEqual("round trip: one file per tab", splitFiles.length, originals.length);
  var splitProblems = [];
  splitFiles.forEach(function (f, i) {
    var wb = TM.readWorkbook(f.data);
    sameCells(originals[i].worksheet, wb.Sheets[wb.SheetNames[0]]).forEach(function (p) { splitProblems.push(f.fileName + ": " + p); });
  });
  assertEqual("round trip: split files match originals cell for cell", splitProblems, []);

  var csvFiles = TM.splitSheets([{ sheetName: "Data", worksheet: north.worksheet }], "csv");
  assertEqual("split csv: values as displayed", csvFiles[0].data.split("\n")[1], "00121,02134,2026-01-15,\"1,234.50\",25%,Café");

    // ---------------------------------------------------------------------
  // Summary
  // ---------------------------------------------------------------------
  results.push("");
  results.push(pass + " passed, " + fail + " failed");
  return results.join("\n");
}
