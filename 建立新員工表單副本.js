// 請在此填入您的範本試算表 ID
const TEMPLATE_ID = "1-I3mGTnQs2jrwWqnNCaI9BiZs0b6SDkmFtPHYOdf9v5x";

function createNewSpreadsheet() {
    try {
        // 2. 準備新檔名
        const ss = SpreadsheetApp.getActiveSpreadsheet();   
        const sheet = ss.getActiveSheet();
        const employeeName = sheet.getRange("B2").getValue(); // 假設姓名在 B2 儲存格
        const startDate = Utilities.formatDate(sheet.getRange("B6").getValue(), "GMT+8", "yyyy-MM-dd");
        const newFileName = employeeName + "_" + startDate;

        // 3. 取得範本檔案並建立副本
        const templateFile = DriveApp.getFileById(TEMPLATE_ID);
        const newFile = templateFile.makeCopy(newFileName);

        // 4. (選用) 如果需要將新檔案放到特定資料夾，可以在這裡實作
        // const folder = DriveApp.getFolderById("TARGET_FOLDER_ID");
        // newFile.moveTo(folder);

        console.log("建立副本成功: " + newFile.getUrl());
        return newFile.getUrl();

    } catch (e) {
        console.error("建立失敗: " + e.toString());
        SpreadsheetApp.getUi().alert("建立失敗: " + e.toString()); // 顯示錯誤給使用者
    }
}