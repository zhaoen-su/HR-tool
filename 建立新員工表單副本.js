// 請在此填入您的範本試算表 ID
const TEMPLATE_ID = "1-I3pxesO-aXYxTjzXOSXVoaeINE36RfNjzOr7MKYzyc";
const TARGET_FOLDER_ID = "1yEf9uy3K0_7BBEOWdpBOZAr80Kz7R3gx";
function createNewSpreadsheet() {
    try {
        // 2. 準備新檔名
        const ss = SpreadsheetApp.getActiveSpreadsheet();   
        const sheet = ss.getActiveSheet();
        const employeeName = sheet.getRange("B2").getValue();
        const startDate = Utilities.formatDate(sheet.getRange("B6").getValue(), "GMT+8", "yyyy-MM-dd");
        const newFileName = employeeName + "-" + startDate;

        // 3. 取得範本檔案並建立副本
        const templateFile = DriveApp.getFileById(TEMPLATE_ID);
        const newFile = templateFile.makeCopy(newFileName);

        const newFileId = newFile.getId();
        const newSS = SpreadsheetApp.openById(newFileId);
        const allSheets = newSS.getSheets();
        const keepSheetName = "資料";

        allSheets.forEach(s => {
            if (s.getName() !== keepSheetName) {
                // 如果分頁名稱不是要保留的，就刪除
                // 注意：試算表至少必須保留一個分頁，否則會報錯
                if (newSS.getSheets().length > 1) {
                    newSS.deleteSheet(s);
                }
            }
        });
        // 4. 放到特定資料夾
        const folder = DriveApp.getFolderById(TARGET_FOLDER_ID);
        newFile.moveTo(folder);

        const htmlOutput = HtmlService.createHtmlOutput(
            `<div style="font-family: sans-serif; text-align: center; padding: 20px;">
                <p style="font-size: 16px;">${newFileName}&nbsp;建立成功！</p>
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