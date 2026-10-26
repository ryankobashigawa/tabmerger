// Generates the sample files in test-fixtures/tools/ used by the merge-into-
// tabs and split tests. Run: osascript -l JavaScript test/make-fixtures.js
// (macOS JavaScriptCore, like run-tests.js; uses the vendored SheetJS).
//
// Covers: several single-tab workbooks, a multi-tab workbook, a CSV, long and
// duplicate names, names with characters Excel or file systems don't allow,
// dates (incl. a 1904-date-system workbook), IDs with leading zeros, a formula.
function run() {
  ObjC.import("Foundation");
  var app = Application.currentApplication();
  app.includeStandardAdditions = true;

  var REPO = "/Users/mockasingo/Documents/tabmerger";
  var OUT = REPO + "/test-fixtures/tools";
  // Read as UTF-8 (app.read() assumes MacRoman and mangles SheetJS's tables).
  function readUtf8(path) {
    return $.NSString.stringWithContentsOfFileEncodingError(path, $.NSUTF8StringEncoding, $()).js;
  }
  eval(readUtf8(REPO + "/vendor/xlsx.full.min.js"));

  var fm = $.NSFileManager.defaultManager;
  function mkdirs(dir) {
    fm.createDirectoryAtPathWithIntermediateDirectoriesAttributesError(dir, true, $(), $());
  }
  function writeWorkbook(path, wb) {
    var b64 = XLSX.write(wb, { type: "base64", bookType: "xlsx" });
    var data = $.NSData.alloc.initWithBase64EncodedStringOptions(b64, 0);
    data.writeToFileAtomically(path, true);
  }
  function writeText(path, text) {
    $.NSString.stringWithString(text).writeToFileAtomicallyEncodingError(path, true, $.NSUTF8StringEncoding, $());
  }

  // Cell helpers. Dates are Excel serials (1900 system) with a date format.
  function date(serial, fmt) {
    return { t: "n", v: serial, z: fmt || "yyyy-mm-dd" };
  }
  function num(v, fmt) {
    return fmt ? { t: "n", v: v, z: fmt } : { t: "n", v: v };
  }
  function text(v) {
    return { t: "s", v: v };
  }
  function sheet(rows) {
    var ws = {};
    var maxC = 0;
    rows.forEach(function (row, r) {
      row.forEach(function (cell, c) {
        if (cell === null) return;
        if (typeof cell === "string") cell = text(cell);
        else if (typeof cell === "number") cell = num(cell);
        ws[XLSX.utils.encode_cell({ r: r, c: c })] = cell;
        if (c > maxC) maxC = c;
      });
    });
    if (rows.length) ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length - 1, c: maxC } });
    return ws;
  }
  function book(tabs, opts) {
    var wb = XLSX.utils.book_new();
    tabs.forEach(function (t) {
      XLSX.utils.book_append_sheet(wb, t[1], t[0]);
    });
    if (opts && opts.date1904) wb.Workbook = { WBProps: { date1904: true } };
    return wb;
  }

  var HEAD = ["Customer ID", "Zip", "Order date", "Amount", "Share", "Note"];
  // 46037 = 2026-01-15, 46068 = 2026-02-15, 46096 = 2026-03-15
  function regionRows(seed, dateFmt) {
    return [
      HEAD,
      [text("00" + (120 + seed)), num(2134, "00000"), date(46037, dateFmt), num(1234.5, "#,##0.00"), num(0.25, "0%"), "Café"],
      [text("0" + (4500 + seed)), num(501, "00000"), date(46068, dateFmt), num(99.99, "#,##0.00"), num(0.5, "0%"), "Leading 'quote"],
      [text("000" + seed), num(90210, "00000"), date(46096, dateFmt), num(-12, "#,##0.00"), num(1, "0%"), null],
    ];
  }

  mkdirs(OUT + "/dupes/a");
  mkdirs(OUT + "/dupes/b");

  // Several single-tab workbooks
  writeWorkbook(OUT + "/region_north.xlsx", book([["Data", sheet(regionRows(1))]]));
  writeWorkbook(OUT + "/region_south.xlsx", book([["Data", sheet(regionRows(2, "m/d/yyyy"))]]));

  // A formula: the cached value (2469) should be kept, the formula dropped.
  var west = sheet(regionRows(3));
  west["G1"] = text("Double");
  west["G2"] = { t: "n", v: 2469, f: "D2*2" };
  west["!ref"] = "A1:G4";
  writeWorkbook(OUT + "/region_west.xlsx", book([["Data", west]]));

  // Illegal tab-name characters in the file name ([ ] aren't allowed in tab names)
  writeWorkbook(OUT + "/Budget [draft] v2.xlsx", book([["Sheet1", sheet([["Line", "Cost"], ["Rent", num(1500, "#,##0")], ["Power", num(220, "#,##0")]])]]));

  // Duplicate file names in different folders
  writeWorkbook(OUT + "/dupes/a/report.xlsx", book([["Report", sheet([["Team", "Score"], ["Red", 10]])]]));
  writeWorkbook(OUT + "/dupes/b/report.xlsx", book([["Report", sheet([["Team", "Score"], ["Blue", 12]])]]));

  // Multi-tab workbook with a long file name, long tab names, characters that
  // are fine in tab names but not in file names, accents, and an empty tab.
  writeWorkbook(
    OUT + "/quarterly_regional_sales_report_2026.xlsx",
    book([
      ["January", sheet([["Rep", "Start date", "ID"], ["Avery", date(46023), text("007")], ["Jordan", date(46024), text("012")]])],
      ["February", sheet([["Rep", "Start date", "ID"], ["Riley", date(46054), text("019")]])],
      ["Summary for leadership review", sheet([["Metric", "Value"], ["Growth", num(0.125, "0.0%")], ["Revenue", num(1250000, "$#,##0")]])],
      ['Q1 "Final" <draft>|v2', sheet([["Status"], ["Locked"]])],
      ["Café São Paulo", sheet([["Loja", "Vendas"], ["Centro", num(3200.75, "#,##0.00")]])],
      ["Empty", {}],
    ])
  );

  // A workbook saved with the 1904 date system (old Mac Excel default).
  // 44575 in the 1904 system is 2026-01-15.
  writeWorkbook(OUT + "/mac_1904_dates.xlsx", book([["Dates", sheet([["Event", "When"], ["Launch", date(44575)]])]], { date1904: true }));

  // A CSV: leading zeros, dates as text, numbers, a long ID, thousands separators.
  writeText(
    OUT + "/customers export.csv",
    [
      "Customer ID,Signup date,US date,Amount,Card ref,Formatted,City",
      "00123,2026-01-15,01/02/2026,1234.50,12345678901234567890,\"1,234\",Café",
      "04567,2026-02-01,02/14/2026,-7,98765432109876543210,\"9,999\",São Paulo",
      "0,2026-03-31,03/31/2026,0.5,11111111111111111111,\"10,000\",Zürich",
    ].join("\n") + "\n"
  );

  return "Wrote fixtures to " + OUT;
}
