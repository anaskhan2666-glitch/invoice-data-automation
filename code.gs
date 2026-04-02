function extractJaredInvoiceData() {
  const mainSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = mainSheet.getLastRow();

  const clientData = mainSheet.getRange("A2:A" + lastRow).getValues();
  const invoiceNames = mainSheet.getRange("B2:B" + lastRow).getValues();

  let poNumbers = [];
  let shippingNumbers = [];
  let dates = [];
  let invoiceNumbers = [];

  // 🔒 Replace with your own file IDs
  const fileIds = [
    "FILE_ID_1",
    "FILE_ID_2",
    "FILE_ID_3",
    "FILE_ID_4"
  ];

  // Load all sheets from provided files
  let allSheets = [];
  fileIds.forEach(id => {
    const ss = SpreadsheetApp.openById(id);
    allSheets = allSheets.concat(ss.getSheets());
  });

  for (let i = 0; i < clientData.length; i++) {
    const client = (clientData[i][0] + "").trim();
    const invoiceKey = (invoiceNames[i][0] + "").trim();

    // Process only Jared client
    if (client !== "Jared" || !invoiceKey) {
      poNumbers.push([""]);
      shippingNumbers.push([""]);
      dates.push([""]);
      invoiceNumbers.push([""]);
      continue;
    }

    let invoiceSheet = null;

    // Search matching sheet
    for (let s of allSheets) {
      if (s.getName().trim() === invoiceKey) {
        invoiceSheet = s;
        break;
      }
    }

    if (!invoiceSheet) {
      Logger.log(`Sheet not found: ${invoiceKey}`);
      poNumbers.push([""]);
      shippingNumbers.push([""]);
      dates.push([""]);
      invoiceNumbers.push([""]);
      continue;
    }

    const data = invoiceSheet.getDataRange().getValues();

    let poVal = "", shippingVal = "", dateVal = "", invoiceVal = "";

    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length - 1; c++) {
        const cell = (data[r][c] + "")
          .replace(/\u00A0/g, " ")
          .trim()
          .toLowerCase();

        // PO value (below the "PO" cell)
        if (!poVal && cell === "po") {
          poVal = (data[r + 1] && data[r + 1][c])
            ? (data[r + 1][c] + "").trim()
            : "";
        }

        // Shipping number (next cell)
        if (!shippingVal && cell.includes("shipping")) {
          shippingVal = (data[r][c + 1] + "").trim();
        }

        // Date (formatted)
        if (!dateVal && cell === "date") {
          let rawDate = data[r][c + 1];
          if (rawDate instanceof Date) {
            dateVal = Utilities.formatDate(
              rawDate,
              SpreadsheetApp.getActive().getSpreadsheetTimeZone(),
              "MM/dd/yy"
            );
          } else {
            dateVal = (rawDate + "").trim();
          }
        }

        // Invoice number (next cell)
        if (!invoiceVal && cell.includes("invoice")) {
          invoiceVal = (data[r][c + 1] + "").trim();
        }
      }
    }

    poNumbers.push([poVal]);
    shippingNumbers.push([shippingVal]);
    dates.push([dateVal]);
    invoiceNumbers.push([invoiceVal]);
  }

  // Write results to sheet
  mainSheet.getRange(2, 3, poNumbers.length, 1).setValues(poNumbers);
  mainSheet.getRange(2, 4, shippingNumbers.length, 1).setValues(shippingNumbers);
  mainSheet.getRange(2, 5, dates.length, 1).setValues(dates);
  mainSheet.getRange(2, 6, invoiceNumbers.length, 1).setValues(invoiceNumbers);

  Logger.log("Jared invoice extraction completed");
}
