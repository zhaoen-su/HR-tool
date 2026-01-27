const SHEET_CONFIG = {
  sheetNamePrefix: "HR_Report_",
  headers: [
    { name: "ID", width: 50 },
    { name: "Employee Name", width: 150 },
    { name: "Department", width: 120 },
    { name: "Join Date", width: 100 }
  ],
  styles: {
    header: {
      backgroundColor: "#4a86e8",
      fontColor: "#ffffff",
      fontWeight: "bold",
      horizontalAlignment: "center"
    },
    frozenRows: 1
  }
};

function createNewSpreadsheet() {
  try {
    const timestamp = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm");
    const newSS = SpreadsheetApp.create(SHEET_CONFIG.sheetNamePrefix + timestamp);
    const sheet = newSS.getSheets()[0];

    // 1. Set Headers
    const headerValues = SHEET_CONFIG.headers.map(h => h.name);
    const headerRange = sheet.getRange(1, 1, 1, headerValues.length);
    headerRange.setValues([headerValues]);

    // 2. Apply Styles
    const styles = SHEET_CONFIG.styles.header;
    headerRange.setBackground(styles.backgroundColor)
      .setFontColor(styles.fontColor)
      .setFontWeight(styles.fontWeight)
      .setHorizontalAlignment(styles.horizontalAlignment);

    // 3. Set Column Widths
    SHEET_CONFIG.headers.forEach((header, index) => {
      sheet.setColumnWidth(index + 1, header.width);
    });

    // 4. Freeze Rows
    if (SHEET_CONFIG.styles.frozenRows > 0) {
      sheet.setFrozenRows(SHEET_CONFIG.styles.frozenRows);
    }

    console.log("建立成功: " + newSS.getUrl());
    return newSS.getUrl();

  } catch (e) {
    console.error("建立失敗: " + e.toString());
  }
}
