# PassRe

*For the power user - This app only runs on Windows*
*Reversible worksheet protection for people who own their spreadsheets.*

Worksheet protection in Excel - "Protect Sheet", "Protect Workbook Structure" - isn't security. It's a convenience lock that stops accidental edits. But it gets in the way when you have every right to the file: a workbook you protected yourself months ago, one a colleague handed over without the password, or a template you need to adjust. PassRe lifts the lock so you can work, then restores it afterwards exactly as it was - same options, same password hash - so you hand the file back with its protection intact, not weakened or stripped.

It works the documented way. An `.xlsx` is a ZIP of XML, and worksheet protection is a single plaintext element inside it; PassRe just does the unzip-edit-rezip faster than doing it by hand. It does **not** touch real security: a password required to *open* a file is genuine AES encryption, and PassRe will not and cannot remove it.

**Responsible use.** Use PassRe on files you own or are authorised to change. Worksheet protection is often there to prevent accidental edits, so restoring it when you're done is the point of the tool, not an afterthought.

---

Remove Excel **protection**, edit the file, then restore it **exactly as it was** - same password hash, same locked options.

No install, no dependencies: pure PowerShell using the .NET ZIP/XML support and WPF that already ship with Windows. **Three small files:**

| File | What it is |
|------|------------|
| `PassRe.bat` | The launcher you run / double-click / drag files onto. |
| `PassRe.ps1` | The whole app (engine + CLI + GUI). |
| `PassRe.Manifest.json` | Version + update source, read by the self-updater. |

## Run it

- **Double-click `PassRe.bat`** → opens the window.
- **Drag an `.xlsx` onto `PassRe.bat`** → opens the window with that workbook loaded.
- **Command line:**

```bat
PassRe.bat unprotect budget.xlsx
```
...edit in Excel, save...
```bat
PassRe.bat reprotect budget.xlsx
```

Also `PassRe.bat status budget.xlsx` and `PassRe.bat toggle budget.xlsx` (reprotect if a stash is waiting to be restored, otherwise unprotect).

The launcher runs PowerShell with `-STA -ExecutionPolicy Bypass`, so WPF works and no machine policy needs changing. If `.bat` is blocked on a locked-down machine you can still call the app directly: `powershell -STA -ExecutionPolicy Bypass -File PassRe.ps1 …`

## Updating

PassRe can update itself from its GitHub repo.

- **GUI:** the window footer shows the current version and a **Check for updates** button. It doesn't check by itself when it opens; click the button to look for a newer version, and if there is one it offers to install it.
- **Command line:** `PassRe.bat update` installs the latest after asking (`-Force` skips the question); `PassRe.bat update -CheckOnly` just reports whether one is available.

How it works: `PassRe.Manifest.json` holds the version and the GitHub coordinates (`repoOwner` / `repoName` / `branch`). The updater fetches the remote manifest from `raw.githubusercontent.com`, compares versions, and - only if the remote is newer - downloads the listed files to a temp folder, **backs up** the current ones to `backup/<timestamp>/`, then replaces them (and **rolls back** from the backup if anything fails). It never touches file-open encryption; it only swaps its own program files. Point it at a fork or private mirror by editing the manifest's repo fields (updates then download every file from there and keep those fields), or by setting `PASSRE_UPDATE_BASEURL`.

## What it handles

An `.xlsx` is a ZIP of XML parts. PassRe removes these protection elements and keeps each one verbatim in a sidecar stash (`<file>.passre`); restoring splices the identical element back into its correct schema position:

| Protection | Element | Lives in |
|---|---|---|
| Worksheet protection ("Protect Sheet") | `<sheetProtection>` | `xl/worksheets/sheetN.xml` |
| Chart-sheet protection | `<sheetProtection>` | `xl/chartsheets/sheetN.xml` |
| Workbook structure ("Protect Workbook Structure") | `<workbookProtection>` | `xl/workbook.xml` |
| File-level "password to modify" | `<fileSharing>` | `xl/workbook.xml` |

❌ **Not** handled: a password required to **open** the file. That's real AES encryption of the whole package, not a plaintext XML element, and can't be removed without the password.

`unprotect` also keeps a one-time `<file>.bak` of the original as a safety net: it is made the first time protection is actually removed and is never overwritten (delete it to take a fresh one; turn off with `-NoBackup`). While a stash is waiting to be restored, `unprotect` refuses to run again - reprotect first.

### CLI options

| Flag | Applies to | Effect |
|------|------------|--------|
| `-NoBackup` | unprotect | Skip the one-time `<file>.bak` copy. |
| `-KeepSidecar` | reprotect | Keep the `.passre` stash, marked as restored, instead of deleting it; the next unprotect that removes protection replaces it. |
| `-Sidecar <path>` | all | Use a custom stash location. |
| `-CheckOnly` | update | Only report whether a newer version exists. |
| `-Force` | update | Install without asking first. |
| `-Quiet` | all | Suppress output. |

(Advanced: `. .\PassRe.ps1 -AsLibrary` dot-sources the engine functions - `Invoke-PassReUnprotect`, `Invoke-PassReReprotect`, `Get-PassReStatus` - for use from your own scripts.)

## How restore stays faithful

The removed element is stored *verbatim* - exact `algorithmName`, `hashValue`, `saltValue`, `spinCount` (or a legacy 16-bit hash). On restore it goes back at the schema-correct spot (worksheet protection right after `</sheetData>`; chart-sheet protection after `<sheetViews>`; `<workbookProtection>` before `<bookViews>`; `<fileSharing>` right after `<fileVersion>`), so Excel opens the file without complaint and the protection comes back exactly as the author set it.

Each sheet's element goes back to the same sheet, matched by its internal sheet id and name rather than by the part file name Excel renumbers when sheets move. If a sheet was deleted or renamed in between, reprotect restores the rest, keeps the stash and names the sheet it could not find: rename it back and run reprotect again, or - if the sheet is gone for good - delete the `.passre` stash once you are sure.
