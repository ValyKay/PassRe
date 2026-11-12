# PassRe

*For the power user.*

Your colleagues and partners keep protecting their worksheets. They mean well - but to you it was never protection, only a quiet tax on your productivity. PassRe lifts that tax. It politely removes every worksheet lock so you can get on with your work, and when you're done it graciously puts each one back, exactly as it was, so nobody is any the wiser.

No hacking, no brute force. PassRe never opens an encrypted file and never recovers a password. It only does what you, the power user, already know how to do by hand - crack the document open with a zip tool, edit the XML, zip it back up - it just does it faster and more quietly, so you can keep your mind on the work and not the ceremony.

A word of etiquette: it's rude to strip the locks off other people's worksheets. Do it only when you genuinely must, and always put the protection back when you're done.

---

Strip Excel **protection**, edit the file, then put it back **exactly as it was** - same password hash, same locked options.

No install, no dependencies: pure PowerShell using the .NET ZIP/XML support and WPF that already ship with Windows. **Two files, that's it:**

| File | What it is |
|------|------------|
| `PassRe.bat` | The launcher you run / double-click / drag files onto. |
| `PassRe.ps1` | The whole app (engine + CLI + GUI). |

## Run it

- **Double-click `PassRe.bat`** → opens the window.
- **Drag an `.xlsx` onto `PassRe.bat`** → opens the window with that workbook loaded.
- **Command line:**

```bat
PassRe.bat unprotect budget.xlsx
```
…edit in Excel, save…
```bat
PassRe.bat reprotect budget.xlsx
```

Also `PassRe.bat status budget.xlsx` and `PassRe.bat toggle budget.xlsx` (unprotect if protected, reprotect if a stash exists).

The launcher runs PowerShell with `-STA -ExecutionPolicy Bypass`, so WPF works and no machine policy needs changing. If `.bat` is blocked on a locked-down machine you can still call the app directly: `powershell -STA -ExecutionPolicy Bypass -File PassRe.ps1 …`

## What it handles

An `.xlsx` is a ZIP of XML parts. PassRe removes these protection elements and keeps each one verbatim in a sidecar stash (`<file>.passre`); restoring splices the identical element back into its correct schema position:

| Protection | Element | Lives in |
|---|---|---|
| Worksheet protection ("Protect Sheet") | `<sheetProtection>` | `xl/worksheets/sheetN.xml` |
| Chart-sheet protection | `<sheetProtection>` | `xl/chartsheets/sheetN.xml` |
| Workbook structure ("Protect Workbook Structure") | `<workbookProtection>` | `xl/workbook.xml` |
| File-level "password to modify" | `<fileSharing>` | `xl/workbook.xml` |

❌ **Not** handled: a password required to **open** the file. That's real AES encryption of the whole package, not a plaintext XML element, and can't be removed without the password.

`unprotect` also drops a one-time `<file>.bak` of the original as a safety net (turn off with `-NoBackup`).

### CLI options

| Flag | Applies to | Effect |
|------|------------|--------|
| `-NoBackup` | unprotect | Skip the one-time `<file>.bak` copy. |
| `-Force` | unprotect | Overwrite an existing stash instead of refusing. |
| `-KeepSidecar` | reprotect | Keep the `.passre` stash instead of deleting it. |
| `-Sidecar <path>` | all | Use a custom stash location. |
| `-Quiet` | all | Suppress output. |

(Advanced: `. .\PassRe.ps1 -AsLibrary` dot-sources the engine functions -
`Invoke-PassReUnprotect`, `Invoke-PassReReprotect`, `Get-PassReStatus` - for use from your own scripts.)

## How restore stays faithful

The removed element is stored *verbatim* - exact `algorithmName`, `hashValue`, `saltValue`, `spinCount` (or a legacy 16-bit hash). On restore it goes back at the schema-correct spot (worksheet protection right after `</sheetData>`; chart-sheet protection after `<sheetViews>`; `<workbookProtection>` before `<bookViews>`; `<fileSharing>` right after `<fileVersion>`), so Excel opens the file without complaint and the protection is indistinguishable from the original.
