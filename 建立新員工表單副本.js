function createNewSpreadsheet() {
    const TEMPLATE_ID = "1wuXFmv1sEi2nea6-zrah_AjRiwsp3eTQ-6wwbZDqYmo";
    const TARGET_FOLDER_ID = "1yEf9uy3K0_7BBEOWdpBOZAr80Kz7R3gx";

    try {
        // 2. 準備新檔名
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName("各項時程");
        const employeeChName = sheet.getRange("B1").getValue();
        const employeeEnName = sheet.getRange("D1").getValue();
        const startDate = Utilities.formatDate(sheet.getRange("C3").getValue(), "GMT+8", "yyyy-MM-dd");
        const newFileName = `${startDate}_${employeeChName} ${employeeEnName}_新人歷程檔案（閱覽權限GL以上主管+HR）`;

        // 3. 取得範本檔案並建立副本
        const templateFile = DriveApp.getFileById(TEMPLATE_ID);
        const newFile = templateFile.makeCopy(newFileName);

        const newFileId = newFile.getId();
        const newUrl = newFile.getUrl();
        const newSS = SpreadsheetApp.openById(newFileId);

        const sheetsToDeleteList = ["索引", "行事曆紀錄"];
        const allSheetsList = newSS.getSheets();

        // 從原始試算表讀取正確的值，寫入副本（避免副本公式未計算的問題）
        allSheetsList.forEach(newSheet => {
            const sheetName = newSheet.getName();
            if (sheetsToDeleteList.includes(sheetName)) {
                try {
                    newSS.deleteSheet(newSheet);
                } catch (e) {
                    console.log("無法刪除分頁: " + sheetName);
                }
            } else {
                const originalSheet = ss.getSheetByName(sheetName);
                if (originalSheet) {
                    const values = originalSheet.getDataRange().getValues();
                    newSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
                }
            }
        });
        // 4. 放到特定資料夾
        const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
        newFile.moveTo(folder);

        const htmlOutput = HtmlService.createHtmlOutput(
            `<div style="font-family: sans-serif; text-align: center; padding: 20px;">
                <p style="font-size: 16px;">${employeeChName}&nbsp;${employeeEnName}&nbsp;的檔案已建立成功！</p>
                <br>
                <a href="${newUrl}" target="_blank" 
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

    } catch (e) {
        console.error("建立失敗: " + e.toString());
        SpreadsheetApp.getUi().alert("建立失敗: " + e.toString()); // 顯示錯誤給使用者
    }
}