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

  function readText(path) {
    return app.read(Path(path));
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
  // Summary
  // ---------------------------------------------------------------------
  results.push("");
  results.push(pass + " passed, " + fail + " failed");
  return results.join("\n");
}
