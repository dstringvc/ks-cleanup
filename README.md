# ks-cleanup

A PowerShell script that reconciles a Kansas legislation tracking spreadsheet (`KSUpdates.xlsx`) against a SQL Server database, updates matching database rows, and kicks off a downstream Node.js downloader for the bills it finds.

## What it does

For each row in `KSUpdates.xlsx`, the script (`ks-cleanup.ps1`):

1. Looks up a matching record in the `RawFileID` table using the row's `State`, `Session`, `BillType`, and `BillNumber` values.
2. Writes the result back into the spreadsheet's `RawFileID` column — either the matching ID, or `"NotFound"`.
3. For matched rows, updates `RawFileID.LegislationPath` (from the row's `BillTextURL`) and sets `RawFileID.LockURL = 1`.
4. Attempts to update `Legis.EffectiveDate` for any row in the `Legis` table with a matching `RawDataID`. It's expected that many `RawFileID`s have no corresponding `Legis` row — that's not an error. When a `Legis` row *is* updated, the spreadsheet's `EffectiveDateUpdated` column is set to `True`.
5. After all rows are processed, saves the spreadsheet, writes every matched `RawFileID` to `download-rawfileid-ks.txt` (one per line), and runs the Node.js downloader against that file.

Errors on an individual row are logged and skipped — the script continues with the remaining rows. A fatal error (e.g. the database is unreachable, or the spreadsheet is missing a required column) stops the run entirely and the Node downloader is not launched.

## Requirements

- Windows PowerShell 5.1 (uses `System.Data.SqlClient`, which ships with the .NET Framework — no extra install needed).
- The [`ImportExcel`](https://github.com/dfinke/ImportExcel) PowerShell module. The script installs it automatically from the PowerShell Gallery on first run if it isn't already present, which requires outbound internet access.
- A `node` installation, with the downloader app available at `C:\apps\llarts\downloader`.
- An environment variable named `DBConnectionStringProd` containing the SQL Server connection string.

## File layout

`KSUpdates.xlsx` must sit in the same directory as `ks-cleanup.ps1`. The script writes its output there too, regardless of the caller's current working directory:

- `download-rawfileid-ks.txt` — matched `RawFileID` values, one per line, passed to the downloader.
- `ks-cleanup.log` — timestamped log of dry-run previews, per-row errors, and fatal errors.

The expected spreadsheet columns are: `State`, `Session`, `BillType`, `BillNumber`, `RawFileID`, `BillTextURL`, `EffectiveDate`, `EffectiveDateUpdated`.

## Usage

By default the script runs in **dry-run mode**: it performs the lookups and populates the spreadsheet, but does not write to the database and does not run the Node downloader. Instead, `ks-cleanup.log` records what each database update would have been.

```powershell
# Dry run — no database writes, no downloader launch
.\ks-cleanup.ps1

# Live run — writes to the database and launches the Node downloader
.\ks-cleanup.ps1 -EnableDatabaseUpdates
```

Use the dry run to review `ks-cleanup.log` and the updated spreadsheet before pointing this at production data.

## Notes

- All SQL statements use parameterized queries.
- The downloader is invoked as: `node C:\apps\llarts\downloader custom --state=TN --rawfilelist=<path-to-download-rawfileid-ks.txt> -d`.
