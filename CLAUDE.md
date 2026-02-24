# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Google Apps Script (GAS) HR onboarding automation tool. Manages employee onboarding schedules via Google Sheets and Google Calendar. All source files use Chinese naming. The `.js` files are GAS scripts that get converted to `.gs` on push.

## Commands

- **Push to Google:** `clasp push`
- **Pull from Google:** `clasp pull`
- **Open in browser:** `clasp open-script`
- **Run tests (watch):** `npm test`
- **Run tests (single run):** `npm run test:run`
- **Run a single test file:** `npx vitest run <filename>`

## Architecture

### Runtime Environment

Google Apps Script V8 runtime (not Node.js). Global objects like `SpreadsheetApp`, `CalendarApp`, `DriveApp`, `LockService`, `Utilities` are provided by GAS — they don't exist locally. Tests must mock these globals.

### Module Structure

| File | Purpose |
|------|---------|
| `工具列.js` | `onOpen()` — creates the custom menu in Google Sheets |
| `加入日期.js` | `generateDate()` — calculates onboarding milestone dates, skipping weekends and Taiwan holidays |
| `建立行事曆.js` | `createCalendar()` — creates Google Calendar events from spreadsheet data |
| `刪除行事曆.js` | `deleteCalendar()` — deletes calendar events with confirmation dialog |
| `建立新員工表單副本.js` | `createNewSpreadsheet()` — copies template spreadsheet for new employees |
| `匯出乾淨的試算表.js` | Exports clean employee progress spreadsheet to Drive folder |
| `onEdit自動更新.js` | `onEdit()` trigger — syncs spreadsheet date changes to calendar events, uses `LockService` for concurrency |

### Key Spreadsheet Sheets

- **「各項時程」** — schedule data with employee/manager/HR info (read by calendar creation)
- **「行事曆紀錄」** — stores created calendar event IDs for tracking/deletion
- **「行事曆控制表」** — triggers `onEdit` auto-update when dates in column C change

### Data Flow

Spreadsheet "各項時程" → read employee/event data → determine attendees by title keywords (主管/Mentor/HR) → create all-day Calendar events → write event IDs to "行事曆紀錄"

### Date Calculation Logic

Base date + intervals `[14, 35, 42, 56, 90, 90, 150, 180, 365]` days. Holidays are fetched from the Taiwan holiday Google Calendar and stored in a `Set` for O(1) lookup. Dates falling on weekends/holidays are shifted forward.

## Testing

Uses Vitest with `globals: true`. Test files: `*.test.js` / `*.spec.js`. The `setup.js` file mocks `SpreadsheetApp` globally. When testing GAS functions, all Google APIs (`CalendarApp`, `DriveApp`, etc.) must be mocked with `vi.fn()`.

## Deployment

CLASP-based manual deployment. The `.claspignore` excludes `node_modules`, test files, config files, and docs from being pushed to Google. Only `.js` source files and `appsscript.json` are deployed.

## Conventions

- **Commit messages:** Semantic prefixes with Chinese descriptions — `feat:`, `fix:`, `docs:`, `refactor:`, `chore:`
- **Language:** Code comments and file names are in Traditional Chinese (zh-TW)
- **Timezone:** Asia/Taipei (configured in `appsscript.json`)
