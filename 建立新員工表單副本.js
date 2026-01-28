// 請在此填入您的範本試算表 ID
const TEMPLATE_ID = "1-I3pxesO-aXYxTjzXOSXVoaeINE36RfNjzOr7MKYzyc";

function createNewSpreadsheet() {
    try {
        // 2. 準備新檔名
        const ss = SpreadsheetApp.getActiveSpreadsheet();   
        const sheet = ss.getActiveSheet();
        const employeeName = sheet.getRange("B2").getValue(); // 假設姓名在 B2 儲存格
        const startDate = Utilities.formatDate(sheet.getRange("B6").getValue(), "GMT+8", "yyyy-MM-dd");
        const newFileName = employeeName + "-" + startDate;

        // 3. 取得範本檔案並建立副本
        const templateFile = DriveApp.getFileById(TEMPLATE_ID);
        const newFile = templateFile.makeCopy(newFileName);

        const newFileId = newFile.getId();
        const newSS = SpreadsheetApp.openById(newFileId); // 打開新複製出來的那份表
        const allSheets = newSS.getSheets();
        const keepSheetName = "資料"; // 👈 請修改這裡

        allSheets.forEach(s => {
            if (s.getName() !== keepSheetName) {
                // 如果分頁名稱不是要保留的，就刪除
                // 注意：試算表至少必須保留一個分頁，否則會報錯
                if (newSS.getSheets().length > 1) {
                    newSS.deleteSheet(s);
                }
            }
        });
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