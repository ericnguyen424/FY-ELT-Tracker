# FY-ELT-Tracker

A Google Apps Script automation for the weekly leadership tracker at [Odyssey House](https://www.odysseyhouse.org), a Utah nonprofit treatment provider. I built it during my Summer 2025 software internship, and the team used it in production after launch in July 2025.

## The problem

Leadership tracked weekly numbers for every program (such as census and capacity) in one large Google Sheet. Each week someone had to add a new block of rows by hand: one row per program, the right date, formatting and borders, carried-over values like capacity, and a totals row with fresh `SUM` formulas. Monthly occupancy averages were also calculated manually. This was slow and easy to get wrong, and a single bad copy-paste could break the totals.

## What it does

- **Adds each new week automatically.** A time-driven trigger runs every Sunday and appends a block of rows for the next week, one per program plus a totals row. It advances the date by 7 days, copies formatting, carries forward selected columns, and writes `SUM` formulas for the totals.
- **Rolls back on failure.** It snapshots the sheet before writing, and if anything goes wrong it deletes the partially inserted rows and restores the original data, so the tracker is never left half-updated.
- **Calculates monthly averages.** A Stats sheet shows each program's average census as a percentage of full capacity for the selected month, recalculated when the sheet is edited. Programs with missing data show "N/A" or "Not found" rather than wrong numbers.
- **Is driven by configuration, not code.** The program list, which columns to total, which columns to carry forward, and the name of the active data sheet (for example, a new fiscal-year tab) all live in config sheets. Non-developers can add a program or start a new fiscal year without touching the script.
- **Guards against mistakes.** It restricts recalculation to authorized users, adds a click cooldown on buttons, clears Google's Smart Fill from new rows, and shows clear toast messages for success and errors.

## How it works

```
_sheetconfig ─┐                ┌─▶ Data sheet (e.g. "FY26 Tracker")
_programlist ─┼─▶ config.js ───┤     ▲  addNewWeek.js   (Sunday trigger)
              │                │     │
              │                └─▶ Stats sheet
              │                      ▲  averages.js     (onEdit)
              └── helper_functions.js: column lookup by header name,
                  sheet access, permissions, cooldowns
```

Columns are looked up by **header name rather than position**, so the team can reorder or add columns in the sheet without breaking the script.

| File | Purpose |
| --- | --- |
| `addNewWeek.js` | Inserts the next week's rows, copies formatting and values, writes totals, rolls back on error |
| `averages.js` | Groups rows by program for the selected month and computes average census / full % |
| `config.js` | Reads the program list, column rules and active sheet name from config sheets |
| `helper_functions.js` | Shared utilities: column lookup, sheet access, permission checks, button cooldown |
| `testing.js` | Test suite that runs against a disposable copy of the live sheet |
| `appsscript.json` | Apps Script manifest (V8 runtime, America/Denver timezone) |

## Testing

Apps Script has no built-in test framework, so I wrote a small test runner. `runTests()` copies the live data sheet into a temporary `TESTING COPY` tab, runs `addNewWeek` against it, checks the result, logs a pass/fail summary, and deletes the copy. Production data is never touched. The tests check that:

- every configured program, plus a totals row, appears in the new week exactly once
- the new week's date is exactly 7 days after the prior week
- no data that shouldn't carry over was copied into the new rows
- the new block has the expected number of rows and columns
- every totals formula references the correct row range

## Setup

1. Open the spreadsheet, then **Extensions → Apps Script**, and add the `.js` files (as `.gs`) and `appsscript.json`.
2. Create the config sheets:
   - `_sheetconfig`: the name of the active data sheet in cell A2
   - `_programlist`: programs in column A, plus header rows that mark which columns to total and which to carry forward
3. Add a **time-driven trigger** that runs `newWeek` weekly on Sunday.
4. Run `runTests` once to confirm everything works against your sheet.

## Tech

Google Apps Script (JavaScript, V8 runtime), Google Sheets API (`SpreadsheetApp`), `PropertiesService`, time-driven and `onEdit` triggers.

## Context

This was part of a broader effort at Odyssey House to move operations off Google Tables after it was discontinued. Alongside this tracker, I evaluated replacement platforms (AppSheet and Monday.com), rebuilt existing Google Tables workflows in Monday.com, and designed Figma prototypes for a community resource guide. After the internship, I was offered a contract to continue the work.
