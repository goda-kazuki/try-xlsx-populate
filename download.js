const XlsxPopulate = require('xlsx-populate');

// Load an existing workbook
XlsxPopulate.fromFileAsync("./template.xlsx")
  .then(workbook => {
    // Modify the workbook.
    workbook.sheet("シート1").cell("A3").value("赤いセルを上書き");
    workbook.sheet("シート1").cell("A7").value("結合セル上書き");
    workbook.sheet("シート1").cell("A12").value("フォント異なるセル上書き");

    workbook.sheet("シート1").cell("D2").value("列幅異なるセル上書き");
    workbook.sheet("シート1").cell("E4").value("行幅異なるセル上書き");

    workbook.sheet("シート2").cell("C1").value("C1上書き");

    // Write to file.
    return workbook.toFileAsync("./out.xlsx");
  });
