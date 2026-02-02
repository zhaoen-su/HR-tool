function exportDataToNewSheet() {
    const TARGET_FOLDER_ID = "1yEf9uy3K0_7BBEOWdpBOZAr80Kz7R3gx";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getActiveSheet();
    const employeeName = sheet.getRange("B2").getValue();
    const startDate = Utilities.formatDate(sheet.getRange("B6").getValue(), "GMT+8", "yyyy-MM-dd");

    // 1. 建立一個全新的試算表檔案
    const newFileName = employeeName + "-" + startDate;
    const newFile = SpreadsheetApp.create(newFileName);
    const newSheet = newFile.getSheets()[0];

    // 2. 決定要複製的範圍 (例如 A1:E10)
    const sourceRange = sheet.getRange("A1:F10");
    const values = sourceRange.getValues();
    const formats = sourceRange.getBackgrounds();
    const fontWeights = sourceRange.getFontWeights();

    // 3. 將資料和格式寫入新表單
    const targetRange = newSheet.getRange(1, 1, values.length, values[0].length);
    targetRange.setValues(values);
    targetRange.setBackgrounds(formats);
    targetRange.setFontWeights(fontWeights);

    // 5. 提示使用者新檔案的連結
    const url = newFile.getUrl();
    const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
    DriveApp.getFileById(newFile.getId()).moveTo(folder);

    const htmlOutput = HtmlService.createHtmlOutput(
        `<div style="font-family: sans-serif; text-align: center; padding: 20px;">
        <p style="font-size: 16px;">新試算表建立成功！</p>
        <br>
        <a href="${url}" target="_blank" 
            style="background-color: #4285f4; color: white; padding: 10px 20px; 
                text-decoration: none; border-radius: 4px; font-weight: bold;">
            前往新試算表
        </a>
        <br><br>
        <p style="color: gray; font-size: 12px;">點擊按鈕將於新分頁開啟</p>
    </div>`
    )
        .setWidth(350)
        .setHeight(220);

    // 彈出視窗
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, "系統訊息");
}