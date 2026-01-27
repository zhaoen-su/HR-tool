# Google Clasp Setup Guide

Clasp (Command Line Apps Script Projects) 讓你在本地終端機開發和管理 Google Apps Script 專案。

## 前置需求

- **Node.js**: 已安裝 (版本 v22.17.0)
- **npm**: 已安裝

## 安裝

我已經為你全域安裝了 `clasp`：

```bash
npm install -g @google/clasp
```

## 登入 (Login)

偵測到你已經登入！ (發現 `~/.clasprc.json` 檔案)。
如果需要重新登入，請執行：

```bash
clasp login
```

這會開啟瀏覽器進行 Google 帳號驗證。

## 啟用 Google Apps Script API

請確認你的 Google 帳號設定中已啟用 [Google Apps Script API](https://script.google.com/home/usersettings)。

## 建立新專案 (Create New Project)

在當前目錄建立新專案：

```bash
clasp create --type standalone --title "My Project"
# 類型可選: standalone, docs, sheets, slides, forms, webapp, api
```

這會生成 [.clasp.json](file:///Users/xxzrainy/Documents/web/HR-tool/.clasp.json) (設定檔) 和 [appsscript.json](file:///Users/xxzrainy/Documents/web/HR-tool/appsscript.json) (專案清單)。

## 複製顯有專案 (Clone Existing Project)

**[已完成]** 專案已成功複製到本地！
Script ID: `1Cl76Q9iEOPD__I3mGTnQs2jrwWqnNCaI9BiZs0b6SDkmFtPHYOdf9v5x`

### 檔案列表：
- [appsscript.json](file:///Users/xxzrainy/Documents/web/HR-tool/appsscript.json): 專案設定檔
- [code.js](file:///Users/xxzrainy/Documents/web/HR-tool/code.js): 主要程式碼
- `工具列.js`
- `檢查工作日.js`
- `測試建立新表單.js`

> **注意**：Google Apps Script 的 `.gs` 檔案在本地會被儲存為 [.js](file:///Users/xxzrainy/Documents/web/HR-tool/code.js) 檔案。當你執行 `clasp push` 時，它們會自動轉換回 `.gs`。

## 基本指令

- `clasp pull`: 從 Google 拉取最新程式碼
- `clasp push`: 上傳本地修改到 Google
- `clasp open-script`: 在瀏覽器開啟專案
- `clasp deploy`: 部署腳本版本
