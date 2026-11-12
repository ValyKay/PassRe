<#
.SYNOPSIS
  PassRe - strip Excel sheet/chart/workbook protection, then put it back exactly.

.DESCRIPTION
  One file, four ways to use it:
    * No arguments .......... opens the GUI.
    * A file path ........... opens the GUI with that workbook loaded (drag-drop).
    * A verb + file ......... command line: unprotect | reprotect | toggle | status
    * update [-CheckOnly] [-Force] ... pull the latest version from GitHub (see PassRe.Manifest.json);
                            -CheckOnly only reports, -Force skips the confirmation prompt

  An .xlsx is a ZIP of XML. PassRe removes the protection elements - <sheetProtection>
  on worksheets and chart sheets, <workbookProtection> and the "password to modify"
  <fileSharing> in the workbook - stashing each one verbatim in a sidecar
  (<file>.passre). "reprotect" splices them back at the correct schema position, so
  the file looks untouched.

  It does NOT handle a password required to OPEN the file: that's AES encryption,
  not a plaintext XML element, and can't be removed without the password.

  Normally you run it through PassRe.bat. Dot-source with -AsLibrary to reuse the
  engine functions from your own scripts.

.EXAMPLE
  .\PassRe.ps1 unprotect budget.xlsx      # edit in Excel, then:
  .\PassRe.ps1 reprotect budget.xlsx

.EXAMPLE
  .\PassRe.ps1                            # just the GUI
  .\PassRe.ps1 status budget.xlsx
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0)][string]$Action,   # verb | file path | (empty => GUI)
    [Parameter(Position = 1)][string]$Path,
    [string]$Sidecar,
    [switch]$NoBackup,
    [switch]$KeepSidecar,
    [switch]$Force,
    [switch]$Quiet,
    [switch]$Gui,
    [switch]$SelfTest,     # smoke check: build the GUI headless, print SELFTEST_OK, exit
    [switch]$AsLibrary,    # dot-source target: define functions, then return
    [switch]$CheckOnly     # update: report whether a newer version exists, don't install
)

$ErrorActionPreference = 'Stop'
try {
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
} catch { }

# ---- self-update configuration ----------------------------------------------
$script:AppRoot               = $PSScriptRoot
$script:ManifestPath          = Join-Path $script:AppRoot 'PassRe.Manifest.json'
$script:UpdateBaseUrlOverride = $env:PASSRE_UPDATE_BASEURL   # private mirror / tests; else GitHub raw
$script:CheckUpdatesOnLaunch  = $false                       # GUI: quiet check when the window opens (off: use the footer button)

#region ============================ ENGINE ====================================

$script:PassReLogSink = $null
function Set-PassReLogSink { param([scriptblock]$Sink) $script:PassReLogSink = $Sink }

function Write-Log {
    param([string]$Message, [ValidateSet('info', 'ok', 'warn', 'err', 'muted', 'head')][string]$Level = 'info')
    if ($script:PassReLogSink) { & $script:PassReLogSink $Message $Level; return }
    $color = switch ($Level) { 'ok' { 'Green' } 'warn' { 'Yellow' } 'err' { 'Red' } 'muted' { 'DarkGray' } 'head' { 'Cyan' } default { 'Gray' } }
    Write-Host $Message -ForegroundColor $color
}

$script:Rx = @{
    Sheet       = '<sheetProtection\b[^>]*?/>'
    Workbook    = '<workbookProtection\b[^>]*?/>'
    FileSharing = '<fileSharing\b[^>]*?/>'
    # Only a <fileSharing> carrying a password is "password to modify"; Excel also writes
    # <fileSharing readOnlyRecommended="1"/> on its own, which is a preference, not protection.
    FileSharingPw = '<fileSharing\b(?=[^>]*\b(?:reservationPassword|hashValue)=)[^>]*?/>'
}

function Resolve-Xlsx {
    param([string]$Path, [switch]$Quiet)
    if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
    $item = Get-Item -LiteralPath $Path
    if ($item.PSIsContainer) { throw "Path is a directory, not a file: $Path" }
    if ($item.Extension -notin '.xlsx', '.xlsm', '.xltx', '.xltm') {
        if (-not $Quiet) { Write-Log "Warning: '$($item.Name)' doesn't look like an Excel package (.xlsx/.xlsm)." 'warn' }
    }
    return $item.FullName
}

# A relative -Sidecar is resolved against PowerShell's location ($PWD) here, because the stash is also read and
# written with [IO.File], which would resolve it against the process's current directory instead.
function Get-SidecarPath {
    param([string]$Full, [string]$Sidecar)
    if ($Sidecar) { return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Sidecar) }
    return "$Full.passre"
}

function Read-EntryString {
    param($Zip, [string]$Name)
    $entry = $Zip.GetEntry($Name); if ($null -eq $entry) { return $null }
    $ms = New-Object System.IO.MemoryStream; $s = $entry.Open()
    try { $s.CopyTo($ms); $bytes = $ms.ToArray() } finally { $s.Dispose(); $ms.Dispose() }
    $enc = New-Object System.Text.UTF8Encoding($false)
    if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) { return $enc.GetString($bytes, 3, $bytes.Length - 3) }
    return $enc.GetString($bytes)
}

function Write-EntryString {
    param($Zip, [string]$Name, [string]$Content)
    $entry = $Zip.GetEntry($Name); if ($null -eq $entry) { throw "Entry not found while writing: $Name" }
    $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes($Content)
    $s = $entry.Open()
    try { $s.SetLength(0); $s.Position = 0; $s.Write($bytes, 0, $bytes.Length); $s.Flush() } finally { $s.Dispose() }
}

function Get-WorksheetEntryNames { param($Zip) $Zip.Entries | Where-Object { $_.FullName -match '^xl/worksheets/sheet[^/]+\.xml$' } | ForEach-Object { $_.FullName } }
function Get-ChartsheetEntryNames { param($Zip) $Zip.Entries | Where-Object { $_.FullName -match '^xl/chartsheets/sheet[^/]+\.xml$' } | ForEach-Object { $_.FullName } }

# Open the package, with clear messages for the usual failures: not a ZIP (typically encrypted with a password
# to OPEN), open in Excel, read-only. -Shared reads even while another program (Excel) has the file open for
# writing; only status uses it - unprotect and reprotect must not work on a file Excel can save over.
function Open-Xlsx {
    param([string]$Full, [string]$Mode, [switch]$Shared)
    $leaf = Split-Path -Leaf $Full
    $fs = $null
    # One untyped catch testing each exception in the chain, outermost first: typed catch clauses picked the wrong
    # clause here in PS 5.1 (see CLAUDE.md, "Traps"), and a 0-byte file's InvalidDataException wraps an IOException.
    try {
        if (-not $Shared) { return [System.IO.Compression.ZipFile]::Open($Full, [System.IO.Compression.ZipArchiveMode]$Mode) }
        $fs = New-Object System.IO.FileStream($Full, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read, $false)
        $fs = $null   # the archive owns the stream now
        return $zip
    } catch {
        if ($fs) { $fs.Dispose() }
        for ($e = $_.Exception; $e; $e = $e.InnerException) {
            if ($e -is [System.IO.InvalidDataException]) { throw "'$leaf' is not a readable Excel package. If it asks for a password to OPEN, it is encrypted, which PassRe can't strip." }
            if ($e -is [System.UnauthorizedAccessException]) {
                if ($Mode -eq 'Update') { throw "'$leaf' can't be changed: it is read-only (file Properties) or in a folder you can't write to." }
                throw "Access to '$leaf' was denied."
            }
            if ($e -is [System.IO.IOException] -and (($e.HResult -band 0xFFFF) -in 32, 33)) { throw "'$leaf' is open in another program (Excel?) - close it there and try again." }   # sharing / lock violation
        }
        throw
    }
}

# Every <sheet> in workbook.xml with its current part (resolved through xl/_rels/workbook.xml.rels).
# Excel renumbers sheetN.xml and the r:ids to follow tab order when it saves; sheetId and name stay with the sheet.
function Get-SheetParts {
    param($Zip)
    $wb = Read-EntryString $Zip 'xl/workbook.xml'; $rels = Read-EntryString $Zip 'xl/_rels/workbook.xml.rels'
    if (-not $wb -or -not $rels) { return }
    $target = @{}
    foreach ($r in ([xml]$rels).Relationships.Relationship) {
        $t = $r.GetAttribute('Target')
        $target[$r.GetAttribute('Id')] = if ($t.StartsWith('/')) { $t.Substring(1) } else { 'xl/' + $t }
    }
    $rNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships', 'http://purl.oclc.org/ooxml/officeDocument/relationships'   # Transitional, Strict
    foreach ($s in ([xml]$wb).workbook.sheets.sheet) {
        $rid = $s.GetAttribute('id', $rNs[0]); if (-not $rid) { $rid = $s.GetAttribute('id', $rNs[1]) }
        [pscustomobject]@{ sheetId = $s.GetAttribute('sheetId'); name = $s.GetAttribute('name'); part = $target[$rid] }
    }
}

# Current part for a stash record, or $null when its sheet can't be matched. sheetId and name must
# both match: Excel can reuse a deleted sheet's id after a reopen (a new sheet that also takes the old
# name can't be told apart), and the part name changes whenever sheets move.
function Resolve-StashRecord {
    param($Rec, $Sheets)
    if ($Rec.kind -eq 'fileSharing' -or $Rec.kind -eq 'workbook') { return 'xl/workbook.xml' }
    if (-not $Rec.sheetId) { return [string]$Rec.entry }   # no identity recorded (stash from PassRe <= v2): part name only
    $hit = @($Sheets | Where-Object { $_.sheetId -eq [string]$Rec.sheetId -and $_.name -eq [string]$Rec.sheet })
    $folder = if ($Rec.kind -eq 'chartsheet') { 'xl/chartsheets/' } else { 'xl/worksheets/' }
    if ($hit.Count -eq 1 -and $hit[0].part -and $hit[0].part.StartsWith($folder)) { return $hit[0].part }
    return $null
}

function Save-PassReStash {
    param($Stash, [string]$Side)
    $tmp = "$Side.tmp"
    try {
        $Stash | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $tmp -Encoding UTF8 -ErrorAction Stop   # keep the BOM: PS 5.1 Get-Content reads BOM-less files as ANSI
        Move-Item -LiteralPath $tmp -Destination $Side -Force -ErrorAction Stop
    } catch { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue; throw }
}

# A stash is active until reprotect has put it back; a kept stash is marked 'restored'.
function Test-PassReStashActive {
    param([string]$Side)
    if (-not (Test-Path -LiteralPath $Side)) { return $false }
    return (-not (Get-Content -LiteralPath $Side -Raw | ConvertFrom-Json).restored)
}

# Worksheet: after </sheetData> (skip optional <sheetCalcPr/>).
function Add-SheetProtection {
    param([string]$Xml, [string]$Element)
    $Xml = [regex]::Replace($Xml, $script:Rx.Sheet, '')
    $idx = $null
    $m = [regex]::Match($Xml, '</sheetData>')
    if ($m.Success) { $idx = $m.Index + $m.Length }
    else { $m2 = [regex]::Match($Xml, '<sheetData\b[^>]*?/>'); if ($m2.Success) { $idx = $m2.Index + $m2.Length } }
    if ($null -eq $idx) { throw "No <sheetData> anchor; cannot place sheet protection." }
    $mc = [regex]::Match($Xml.Substring($idx), '^\s*<sheetCalcPr\b[^>]*?/>'); if ($mc.Success) { $idx += $mc.Length }
    return $Xml.Substring(0, $idx) + $Element + $Xml.Substring($idx)
}

# Chart sheet (no <sheetData>): after sheetViews, else sheetPr, else root open tag.
function Add-ChartsheetProtection {
    param([string]$Xml, [string]$Element)
    $Xml = [regex]::Replace($Xml, $script:Rx.Sheet, '')
    $idx = $null
    foreach ($pat in '</sheetViews>', '<sheetViews\b[^>]*?/>', '</sheetPr>', '<sheetPr\b[^>]*?/>') {
        $m = [regex]::Match($Xml, $pat); if ($m.Success) { $idx = $m.Index + $m.Length; break }
    }
    if ($null -eq $idx) {
        $m = [regex]::Match($Xml, '<chartsheet\b[^>]*?>'); if (-not $m.Success) { throw "No <chartsheet> root; cannot place chart-sheet protection." }
        $idx = $m.Index + $m.Length
    }
    return $Xml.Substring(0, $idx) + $Element + $Xml.Substring($idx)
}

# Workbook structure: before <bookViews> (else <sheets>).
function Add-WorkbookProtection {
    param([string]$Xml, [string]$Element)
    $Xml = [regex]::Replace($Xml, $script:Rx.Workbook, '')
    $m = [regex]::Match($Xml, '<bookViews\b'); if (-not $m.Success) { $m = [regex]::Match($Xml, '<sheets\b') }
    if (-not $m.Success) { throw "No <bookViews>/<sheets> anchor in workbook.xml." }
    return $Xml.Substring(0, $m.Index) + $Element + $Xml.Substring($m.Index)
}

# fileSharing: before the first of workbookPr / workbookProtection / bookViews / sheets.
function Add-FileSharing {
    param([string]$Xml, [string]$Element)
    $Xml = [regex]::Replace($Xml, $script:Rx.FileSharing, '')
    $best = [int]::MaxValue
    foreach ($pat in '<workbookPr\b', '<workbookProtection\b', '<bookViews\b', '<sheets\b') {
        $m = [regex]::Match($Xml, $pat); if ($m.Success -and $m.Index -lt $best) { $best = $m.Index }
    }
    if ($best -eq [int]::MaxValue) { throw "No workbookPr/bookViews/sheets anchor in workbook.xml for fileSharing." }
    return $Xml.Substring(0, $best) + $Element + $Xml.Substring($best)
}

function Get-PassReStatus {
    param([string]$Full, [string]$Sidecar)
    $side = Get-SidecarPath -Full $Full -Sidecar $Sidecar
    $sheetCount = 0; $chartCount = 0; $wbProtected = $false; $fileShare = $false
    $zip = Open-Xlsx $Full 'Read' -Shared
    try {
        $wb = Read-EntryString $zip 'xl/workbook.xml'
        if ($wb) {
            if ([regex]::IsMatch($wb, $script:Rx.Workbook)) { $wbProtected = $true }
            if ([regex]::IsMatch($wb, $script:Rx.FileSharingPw)) { $fileShare = $true }
        }
        foreach ($n in Get-WorksheetEntryNames $zip) { $x = Read-EntryString $zip $n; if ($x -and [regex]::IsMatch($x, $script:Rx.Sheet)) { $sheetCount++ } }
        foreach ($n in Get-ChartsheetEntryNames $zip) { $x = Read-EntryString $zip $n; if ($x -and [regex]::IsMatch($x, $script:Rx.Sheet)) { $chartCount++ } }
    } finally { $zip.Dispose() }

    # A stash that can't be parsed counts as active (fail-safe): it blocks unprotect, and reprotect reports the parse error.
    $stashCount = 0; $stashCreated = $null; $stashRestored = $null; $hasStash = Test-Path -LiteralPath $side
    if ($hasStash) { try { $d = Get-Content -LiteralPath $side -Raw | ConvertFrom-Json; $stashCount = @($d.entries).Count; $stashCreated = $d.created; $stashRestored = $d.restored } catch { } }
    $stashActive = ($hasStash -and -not $stashRestored)
    $anyProt = ($sheetCount -gt 0 -or $chartCount -gt 0 -or $wbProtected -or $fileShare)
    $next = if ($stashActive) { 'reprotect' } elseif ($anyProt) { 'unprotect' } else { 'none' }

    [pscustomobject]@{
        File = $Full; Sidecar = $side
        ProtectedSheets = $sheetCount; ProtectedChartSheets = $chartCount
        WorkbookProtected = $wbProtected; PasswordToModify = $fileShare
        HasStash = $hasStash; StashActive = $stashActive; StashCount = $stashCount; StashCreated = $stashCreated; StashRestored = $stashRestored
        AnyProtection = $anyProt; Next = $next
    }
}

function Invoke-PassReUnprotect {
    param([string]$Full, [hashtable]$Opt = @{})
    $side = Get-SidecarPath -Full $Full -Sidecar $Opt.Sidecar
    if (Test-PassReStashActive $side) { throw "A stash from an earlier unprotect has not been restored yet: $side`nRun reprotect first. If that stash is obsolete, delete it by hand." }

    # 1. Scan read-only and work out the edits; the workbook itself isn't touched until step 4.
    $records = New-Object System.Collections.Generic.List[object]
    $found = New-Object System.Collections.Generic.List[string]
    $edits = [ordered]@{}; $orig = @{}
    $zip = Open-Xlsx $Full 'Read'
    try {
        $sheets = @(Get-SheetParts $zip)
        $wb = Read-EntryString $zip 'xl/workbook.xml'
        if ($wb) {
            $wb0 = $wb; $changed = $false
            $m = [regex]::Match($wb, $script:Rx.FileSharingPw)
            if ($m.Success) { $records.Add([pscustomobject]@{ entry = 'xl/workbook.xml'; kind = 'fileSharing'; element = $m.Value }); $wb = $wb.Remove($m.Index, $m.Length); $changed = $true; $found.Add("Removed file-level 'password to modify'.") }
            $m = [regex]::Match($wb, $script:Rx.Workbook)
            if ($m.Success) { $records.Add([pscustomobject]@{ entry = 'xl/workbook.xml'; kind = 'workbook'; element = $m.Value }); $wb = $wb.Remove($m.Index, $m.Length); $changed = $true; $found.Add("Removed workbook structure protection.") }
            if ($changed) { $edits['xl/workbook.xml'] = $wb; $orig['xl/workbook.xml'] = $wb0 }
        }
        foreach ($kind in 'sheet', 'chartsheet') {
            $names = if ($kind -eq 'sheet') { Get-WorksheetEntryNames $zip } else { Get-ChartsheetEntryNames $zip }
            foreach ($name in $names) {
                $x = Read-EntryString $zip $name; if (-not $x) { continue }
                $m = [regex]::Match($x, $script:Rx.Sheet); if (-not $m.Success) { continue }
                $s = $sheets | Where-Object { $_.part -eq $name } | Select-Object -First 1
                $records.Add([pscustomobject]@{ entry = $name; kind = $kind; element = $m.Value; sheetId = $s.sheetId; sheet = $s.name })
                $edits[$name] = $x.Remove($m.Index, $m.Length); $orig[$name] = $x
                $label = if ($s) { "'$($s.name)' ($name)" } else { $name }
                $found.Add($(if ($kind -eq 'sheet') { "Removed sheet protection -> $label" } else { "Removed chart-sheet protection -> $label" }))
            }
        }
    } finally { $zip.Dispose() }

    if ($records.Count -eq 0) {
        Write-Log "No sheet/chart/workbook protection found - nothing to remove." 'warn'
        Write-Log "(A password required to OPEN the file is encryption, which PassRe can't strip.)" 'muted'
        return [pscustomobject]@{ Removed = 0; Stash = $null; Backup = $null }
    }

    # 2. One-time backup: an existing .bak is the oldest original and is never overwritten.
    $backup = $null; $madeBackup = $false
    if (-not $Opt.NoBackup) {
        $backup = "$Full.bak"
        if (Test-Path -LiteralPath $backup) { Write-Log "Backup kept from an earlier run ($((Get-Item -LiteralPath $backup).LastWriteTime.ToString('yyyy-MM-dd HH:mm'))) -> $backup" 'muted' }
        else { Copy-Item -LiteralPath $Full -Destination $backup -ErrorAction Stop; $madeBackup = $true; Write-Log "Backup of the original  -> $backup" 'muted' }
    }

    # 3. Save the stash before touching the workbook, so a stripped workbook always has one.
    # A kept (restored) stash gets replaced here; remember it so an abort below can put it back.
    $prevStash = if (Test-Path -LiteralPath $side) { [IO.File]::ReadAllBytes($side) } else { $null }
    $stash = [pscustomobject]@{ tool = 'PassRe'; version = 3; created = (Get-Date).ToString('o'); originalFile = (Split-Path -Leaf $Full); entries = $records.ToArray() }
    try { Save-PassReStash $stash $side }
    catch {
        $msg = $_.Exception.Message
        if ($madeBackup) { Remove-Item -LiteralPath $backup -Force -ErrorAction SilentlyContinue }
        if ($prevStash -and -not (Test-Path -LiteralPath $side)) { try { [IO.File]::WriteAllBytes($side, $prevStash) } catch { } }
        throw "Could not save the stash to ${side}: $msg`nThe workbook was not changed."
    }

    # 4. Commit the edits - but only onto the parts as they were scanned; another program may have saved meanwhile.
    # When nothing gets written, undo this run's .bak and stash (putting back a kept stash it replaced) so they
    # don't block the next run.
    $undo = {
        if ($madeBackup) { Remove-Item -LiteralPath $backup -Force -ErrorAction SilentlyContinue }
        if ($prevStash) {
            try { [IO.File]::WriteAllBytes($side, $prevStash); '' }
            catch { "`nThe earlier kept stash at $side could not be put back; it now holds this run's stash - delete it by hand before running unprotect again." }
        } else {
            Remove-Item -LiteralPath $side -Force -ErrorAction SilentlyContinue
            if (Test-Path -LiteralPath $side) { "`nThe stash at $side could not be removed - delete it by hand before running unprotect again." } else { '' }
        }
    }
    $stale = $null; $opened = $false
    try {
        $zip = Open-Xlsx $Full 'Update'; $opened = $true
        try {
            foreach ($k in $edits.Keys) { if ((Read-EntryString $zip $k) -cne $orig[$k]) { $stale = $k; break } }
            if (-not $stale) { foreach ($k in $edits.Keys) { Write-EntryString $zip $k $edits[$k] } }
        } finally { $zip.Dispose() }
    } catch {
        $msg = $_.Exception.Message
        if (-not $opened) { throw ("Could not open the workbook for writing: $msg`nNothing was removed." + (& $undo)) }
        throw "Could not write the workbook: $msg`nThe stash is kept at $side - run reprotect to make sure the protection is back."
    }
    if ($stale) { throw ("'$stale' changed while PassRe was working (saved by another program?). Nothing was removed - run unprotect again." + (& $undo)) }

    foreach ($line in $found) { Write-Log $line 'ok' }
    Write-Log "Unprotected. Stashed $($records.Count) element(s) -> $side" 'head'
    return [pscustomobject]@{ Removed = $records.Count; Stash = $side; Backup = $backup }
}

function Invoke-PassReReprotect {
    param([string]$Full, [hashtable]$Opt = @{})
    $side = Get-SidecarPath -Full $Full -Sidecar $Opt.Sidecar
    if (-not (Test-Path -LiteralPath $side)) { throw "No stash at: $side`nCan't restore the original protection without it. (Run unprotect first.)" }
    $stash = Get-Content -LiteralPath $side -Raw | ConvertFrom-Json
    $recs = @($stash.entries | Where-Object { $_ })
    if ($recs.Count -eq 0) { throw "Stash is empty: $side" }
    if ([int]$stash.version -lt 3) { Write-Log "This stash is from an older PassRe and names sheets by part only; if sheets were moved, added or deleted since unprotect, check the result." 'warn' }
    if ($stash.restored) { Write-Log "This stash was already restored ($($stash.restored)); reprotecting from it replaces the sheets' current protection with the stashed one." 'warn' }

    $restored = 0; $missed = New-Object System.Collections.Generic.List[string]
    $zip = Open-Xlsx $Full 'Update'
    try {
        $sheets = @(Get-SheetParts $zip)
        foreach ($rec in $recs) {
            $part = Resolve-StashRecord $rec $sheets
            $x = if ($part) { Read-EntryString $zip $part } else { $null }
            $label = if ($rec.sheet) { "'$($rec.sheet)'" } else { [string]$rec.entry }
            if ($null -eq $x) { $missed.Add($label); Write-Log "Not restored - no matching sheet: $label" 'warn'; continue }
            switch ($rec.kind) {
                'fileSharing' { Write-EntryString $zip $part (Add-FileSharing         $x $rec.element); Write-Log "Restored 'password to modify'." 'ok' }
                'workbook'    { Write-EntryString $zip $part (Add-WorkbookProtection   $x $rec.element); Write-Log "Restored workbook structure protection." 'ok' }
                'chartsheet'  { Write-EntryString $zip $part (Add-ChartsheetProtection $x $rec.element); Write-Log "Restored chart-sheet protection -> $label" 'ok' }
                default       { Write-EntryString $zip $part (Add-SheetProtection      $x $rec.element); Write-Log "Restored sheet protection -> $label" 'ok' }
            }
            $restored++
        }
    } finally { $zip.Dispose() }

    # Keep the stash whenever something could not be put back, so nothing is lost. A kept stash that was marked
    # restored is active again: it still holds protection the workbook lacks, so unprotect must not replace it.
    if ($missed.Count -gt 0) {
        if ($stash.restored) {
            $stash.PSObject.Properties.Remove('restored')
            try { Save-PassReStash $stash $side } catch { Write-Log "Could not mark the stash as waiting to be restored: $($_.Exception.Message)" 'warn' }
        }
        throw "Restored $restored of $($recs.Count) element(s); no matching sheet for: $($missed -join ', ') (deleted or renamed since unprotect?).`nThe stash is kept: $side`nRename the sheet back and run reprotect again, or delete the stash once you are sure."
    }
    if (-not $Opt.KeepSidecar) { Remove-Item -LiteralPath $side -Force; Write-Log "Stash cleared." 'muted' }
    else {
        $stash | Add-Member -NotePropertyName restored -NotePropertyValue (Get-Date).ToString('o') -Force
        Save-PassReStash $stash $side
        Write-Log "Stash kept, marked restored (the next unprotect replaces it): $side" 'muted'
    }
    Write-Log "Reprotected: $restored element(s) put back exactly as they were." 'head'
    return [pscustomobject]@{ Restored = $restored }
}

# ---- self-update (manifest + GitHub raw, with backup / rollback) -------------

function Enable-Tls12 {
    try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
}

function Get-LocalManifest {
    if (-not (Test-Path -LiteralPath $script:ManifestPath)) { return $null }
    try { return (Get-Content -LiteralPath $script:ManifestPath -Raw | ConvertFrom-Json) } catch { return $null }
}

function Get-PassReVersion {
    $m = Get-LocalManifest
    if ($m -and ($m.PSObject.Properties.Name -contains 'version')) { return [string]$m.version }
    return $null
}

function Assert-ManifestShape {
    param($Manifest, [string]$Label)
    foreach ($p in 'appName', 'version', 'repoOwner', 'repoName', 'branch', 'files') {
        if (-not ($Manifest.PSObject.Properties.Name -contains $p)) { throw ("{0} is missing '{1}'." -f $Label, $p) }
    }
    if ([string]::IsNullOrWhiteSpace([string]$Manifest.version)) { throw ("{0} 'version' is empty." -f $Label) }
    if (@($Manifest.files).Count -lt 1) { throw ("{0} 'files' is empty." -f $Label) }
    foreach ($core in 'PassRe.ps1', 'PassRe.bat', 'PassRe.Manifest.json') {
        if (@($Manifest.files) -notcontains $core) { throw ("{0} 'files' is missing '{1}'." -f $Label, $core) }
    }
}

function Get-RawBaseUrl {
    param($Manifest)
    if ($script:UpdateBaseUrlOverride) { return ([string]$script:UpdateBaseUrlOverride).TrimEnd('/') }
    return ('https://raw.githubusercontent.com/{0}/{1}/{2}' -f $Manifest.repoOwner, $Manifest.repoName, $Manifest.branch)
}

function Compare-PassReVersion {
    param([string]$Left, [string]$Right)
    try { return ([version]$Left).CompareTo([version]$Right) }
    catch { return [string]::Compare($Left, $Right, [System.StringComparison]::OrdinalIgnoreCase) }
}

function Get-RemoteManifest {
    param([string]$BaseUrl)
    Enable-Tls12
    $resp = Invoke-WebRequest -Uri ('{0}/PassRe.Manifest.json' -f $BaseUrl) -UseBasicParsing -TimeoutSec 20 -ErrorAction Stop
    # Decode the raw bytes as UTF-8 whatever content type / charset the server sent (mirrors may send
    # octet-stream or no charset), then tolerate a UTF-8 BOM.
    $txt = (New-Object System.Text.UTF8Encoding($false)).GetString($resp.RawContentStream.ToArray()).TrimStart([char]0xFEFF)
    return ($txt | ConvertFrom-Json -ErrorAction Stop)
}

function Get-FriendlyUpdateError {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return 'Update failed.' }
    if ($Text -match '(?i)was empty|empty response|zero bytes') { return 'A downloaded file was empty - usually transient. Retry, or check your network/proxy.' }
    if ($Text -match '\(404\)') { return 'The update source returned 404 (not found). Check repoOwner/repoName/branch in PassRe.Manifest.json, and that every file it lists is published.' }
    if ($Text -match '(?i)could not be resolved|name resolution|no such host|could not resolve') { return 'Could not reach the update host. Check your network connection.' }
    if ($Text -match '(?i)ssl|tls|secure channel|trust relationship') { return 'TLS negotiation failed. A TLS-inspecting proxy may be the cause.' }
    if ($Text -match '(?i)\(40[13]\)|unauthori|forbidden') { return 'The update source returned an auth error. Check the repo/branch in PassRe.Manifest.json.' }
    return "Update failed: $Text"
}

# Staged files must be what the manifest promised: none may be an error, sign-in or block page; PassRe.ps1 must
# parse AND end with its '# PassRe:EOF' line (plain text such as "Access Denied", or a script cut off between two
# functions, still parses); PassRe.bat must launch PassRe.ps1; the manifest must be well-formed and the same version.
function Assert-StagedUpdate {
    param([string]$StageRoot, $Files, $Remote)
    foreach ($f in $Files) {
        $txt = [IO.File]::ReadAllText((Join-Path $StageRoot $f))
        if ($txt -match '(?is)^\s*(<\?xml[^>]*\?>\s*)?(<!--.*?-->\s*)*<(!doctype\s+html|html[\s>])') { throw "Downloaded file '$f' is a web page, not part of PassRe (a proxy or sign-in page?)." }
    }
    $ps1 = Join-Path $StageRoot 'PassRe.ps1'
    $tokens = $null; $perr = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($ps1, [ref]$tokens, [ref]$perr)
    if (@($perr).Count -gt 0) { throw "Downloaded PassRe.ps1 does not parse (cut off?): $($perr[0].Message)" }
    if (-not [IO.File]::ReadAllText($ps1).TrimEnd().EndsWith('# PassRe:EOF', [StringComparison]::Ordinal)) { throw "Downloaded PassRe.ps1 is incomplete or not PassRe (its end marker is missing) - a cut-off download or a proxy error?" }
    if ($Files -contains 'PassRe.bat' -and [IO.File]::ReadAllText((Join-Path $StageRoot 'PassRe.bat')) -notmatch 'PassRe\.ps1') { throw "Downloaded PassRe.bat is not the PassRe launcher (a proxy error?)." }
    $m = [IO.File]::ReadAllText((Join-Path $StageRoot 'PassRe.Manifest.json')).TrimStart([char]0xFEFF) | ConvertFrom-Json
    Assert-ManifestShape -Manifest $m -Label 'Downloaded manifest'
    if ([string]$m.version -ne [string]$Remote.version) { throw "Downloaded manifest is v$($m.version), expected v$($Remote.version) - the source changed mid-update. Try again in a few minutes." }
}

function Invoke-PassReSelfUpdate {
    param([switch]$CheckOnly, [scriptblock]$Confirm, [switch]$Force)

    $local = Get-LocalManifest
    if (-not $local) { throw "PassRe.Manifest.json beside the app is missing or not valid JSON - can't check for updates." }
    Assert-ManifestShape -Manifest $local -Label 'Local manifest'
    $base = Get-RawBaseUrl -Manifest $local   # one source for the whole update: the manifest and every file
    $remote = Get-RemoteManifest -BaseUrl $base
    Assert-ManifestShape -Manifest $remote -Label 'Remote manifest'

    if ((Compare-PassReVersion $remote.version $local.version) -le 0) {
        Write-Log "You're on the latest version (v$($local.version))." 'ok'
        return [pscustomobject]@{ Updated = $false; UpToDate = $true; Current = $local.version; Remote = $remote.version }
    }
    Write-Log "Update available: v$($local.version) -> v$($remote.version)" 'head'
    if ($CheckOnly) { return [pscustomobject]@{ Updated = $false; UpToDate = $false; Current = $local.version; Remote = $remote.version } }
    if (-not $Force -and $Confirm -and -not (& $Confirm $local.version $remote.version)) {
        Write-Log 'Update cancelled.' 'muted'
        return [pscustomobject]@{ Updated = $false; UpToDate = $false; Cancelled = $true; Current = $local.version; Remote = $remote.version }
    }

    $files = @($remote.files)
    $ts = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backupRoot = Join-Path $script:AppRoot "backup\$ts"
    $stageRoot = Join-Path ([System.IO.Path]::GetTempPath()) "PassRe\update-$ts"
    New-Item -ItemType Directory -Path $stageRoot -Force | Out-Null

    try {
        # Phase 1: download everything to staging (each file must be non-empty, then pass Assert-StagedUpdate), before touching the install.
        Enable-Tls12
        foreach ($f in $files) {
            $stage = Join-Path $stageRoot $f
            New-Item -ItemType Directory -Path (Split-Path -Parent $stage) -Force | Out-Null
            Invoke-WebRequest -Uri ('{0}/{1}' -f $base, $f) -OutFile $stage -UseBasicParsing -TimeoutSec 60 -ErrorAction Stop
            if ((Get-Item -LiteralPath $stage).Length -le 0) { throw "Downloaded file '$f' was empty." }
        }
        Assert-StagedUpdate -StageRoot $stageRoot -Files $files -Remote $remote

        # Keep this install's update source: the new manifest must not repoint it (forks, mirrors).
        $mPath = Join-Path $stageRoot 'PassRe.Manifest.json'
        $sm = [IO.File]::ReadAllText($mPath).TrimStart([char]0xFEFF) | ConvertFrom-Json
        $moved = $false
        foreach ($k in 'repoOwner', 'repoName', 'branch') { if ([string]$sm.$k -ne [string]$local.$k) { $sm.$k = $local.$k; $moved = $true } }
        if ($moved) { [IO.File]::WriteAllText($mPath, ($sm | ConvertTo-Json -Depth 5), (New-Object System.Text.UTF8Encoding($false))) }

        # Phase 2: back up current files (backup\<ts> is created here, only once downloads succeeded), then replace; roll back on any failure.
        $replaced = New-Object System.Collections.Generic.List[string]
        try {
            foreach ($f in $files) {
                $cur = Join-Path $script:AppRoot $f
                if (Test-Path -LiteralPath $cur) {
                    $bf = Join-Path $backupRoot $f
                    New-Item -ItemType Directory -Path (Split-Path -Parent $bf) -Force | Out-Null
                    Copy-Item -LiteralPath $cur -Destination $bf -Force
                }
            }
            $order = @($files | Where-Object { $_ -ne 'PassRe.ps1' }) + @($files | Where-Object { $_ -eq 'PassRe.ps1' })
            foreach ($f in $order) {
                $replaced.Add($f)   # before the copy: a copy that fails partway has already overwritten the file
                Copy-Item -LiteralPath (Join-Path $stageRoot $f) -Destination (Join-Path $script:AppRoot $f) -Force
                Write-Log "Updated $f" 'ok'
            }
        } catch {
            # Restore every file, even if one restore fails (antivirus briefly locking a fresh script), and keep the
            # original error.
            $err = $_
            $stuck = New-Object System.Collections.Generic.List[string]
            foreach ($f in $replaced) {
                $bf = Join-Path $backupRoot $f
                $cur = Join-Path $script:AppRoot $f
                try {
                    if (Test-Path -LiteralPath $bf) { Copy-Item -LiteralPath $bf -Destination $cur -Force -ErrorAction Stop }
                    elseif (Test-Path -LiteralPath $cur) { Remove-Item -LiteralPath $cur -Force -ErrorAction Stop }
                } catch { $stuck.Add($f) }
            }
            if ($stuck.Count -gt 0) { throw ("{0}`nCould not roll back {1}; copy the previous version back by hand from {2}" -f $err.Exception.Message, ($stuck -join ', '), $backupRoot) }
            throw $err
        }
    } finally {
        Remove-Item -LiteralPath $stageRoot -Recurse -Force -ErrorAction SilentlyContinue   # temp staging only
    }

    Write-Log "Updated to v$($remote.version). Restart PassRe to use it." 'head'
    Write-Log "Backup of the previous version: $backupRoot" 'muted'
    return [pscustomobject]@{ Updated = $true; Current = $local.version; Remote = $remote.version; Backup = $backupRoot }
}

#endregion

#region ============================= GUI ======================================

$script:UI = @{}

function Set-GuiVal { param($Tb, [string]$Text, [bool]$Warn) $Tb.Text = $Text; $Tb.Foreground = if ($Warn) { '#C05621' } else { '#2F855A' } }

function Clear-GuiStatusValues {
    $g = $script:UI
    foreach ($v in $g.WbVal, $g.FsVal, $g.ShVal, $g.ChVal, $g.StVal) { $v.Text = [string][char]0x2014; $v.Foreground = '#718096' }
    $g.UnprotectBtn.IsEnabled = $false; $g.ReprotectBtn.IsEnabled = $false; $g.ToggleBtn.IsEnabled = $false
}

function Update-GuiStatus {
    $g = $script:UI
    if (-not $g.File) {
        $g.FileText.Text = 'Drop an .xlsx here'; $g.PathText.Text = 'or click Browse to choose a file'
        Clear-GuiStatusValues
        return
    }
    $g.FileText.Text = (Split-Path -Leaf $g.File); $g.PathText.Text = $g.File
    # On a read failure (e.g. the workbook is open in Excel) don't leave the previous file's values and buttons showing.
    try { $s = Get-PassReStatus -Full $g.File } catch { $msg = $_.Exception.Message; Clear-GuiStatusValues; Write-Log "ERROR: $msg" 'err'; return }
    Set-GuiVal $g.WbVal $(if ($s.WorkbookProtected) { 'PROTECTED' } else { 'none' }) $s.WorkbookProtected
    Set-GuiVal $g.FsVal $(if ($s.PasswordToModify) { 'SET' } else { 'none' }) $s.PasswordToModify
    Set-GuiVal $g.ShVal "$($s.ProtectedSheets)" ($s.ProtectedSheets -gt 0)
    Set-GuiVal $g.ChVal "$($s.ProtectedChartSheets)" ($s.ProtectedChartSheets -gt 0)
    if ($s.StashActive) { Set-GuiVal $g.StVal "present ($($s.StashCount))" $false }
    elseif ($s.HasStash) { $g.StVal.Text = "kept, restored ($($s.StashCount))"; $g.StVal.Foreground = '#718096' }
    else { $g.StVal.Text = 'none'; $g.StVal.Foreground = '#718096' }
    $g.UnprotectBtn.IsEnabled = ($s.AnyProtection -and -not $s.StashActive)
    $g.ReprotectBtn.IsEnabled = $s.HasStash
    $g.ToggleBtn.IsEnabled = ($s.AnyProtection -or $s.StashActive)
}

function Set-GuiFile {
    param([string]$Path)
    if (-not $Path) { return }
    if (-not (Test-Path -LiteralPath $Path)) { Write-Log "Not found: $Path" 'err'; return }
    if (Test-Path -LiteralPath $Path -PathType Container) { Write-Log "That's a folder - drop a workbook file: $Path" 'err'; return }
    $script:UI.File = (Get-Item -LiteralPath $Path).FullName
    Write-Log "Loaded: $(Split-Path -Leaf $script:UI.File)" 'head'
    Update-GuiStatus
}

function Get-GuiOpt { @{ NoBackup = (-not $script:UI.BackupChk.IsChecked); KeepSidecar = [bool]$script:UI.KeepStashChk.IsChecked } }

# Non-blocking update check: fetch the remote version on a background runspace,
# poll for completion with a DispatcherTimer so the window stays responsive during
# the check. (The install itself, Invoke-GuiUpdateInstall, runs on the UI thread.)
function Start-GuiUpdateCheck {
    param([switch]$Silent)
    if ($script:UI.UpdateBusy) { return }
    $local = Get-LocalManifest
    $why = if (-not $local) { 'PassRe.Manifest.json beside the app is missing or not valid JSON.' }
           else { try { Assert-ManifestShape -Manifest $local -Label 'PassRe.Manifest.json'; $null } catch { $_.Exception.Message } }
    if ($why) {
        if ($Silent) { Write-Log "Update check skipped: $why" 'muted' }
        else { [System.Windows.MessageBox]::Show("$why`n`nUpdates cannot be checked.", 'Updates', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null }
        return
    }
    $url = '{0}/PassRe.Manifest.json' -f (Get-RawBaseUrl -Manifest $local)
    $ps = [powershell]::Create()
    [void]$ps.AddScript({
            param($u)
            $ErrorActionPreference = 'Stop'
            try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
            try {
                $r = Invoke-WebRequest -Uri $u -UseBasicParsing -TimeoutSec 20 -ErrorAction Stop
                $m = ((New-Object System.Text.UTF8Encoding($false)).GetString($r.RawContentStream.ToArray()).TrimStart([char]0xFEFF)) | ConvertFrom-Json -ErrorAction Stop   # same decoding as Get-RemoteManifest
                $v = [string]$m.version
                if ([string]::IsNullOrWhiteSpace($v)) { [pscustomobject]@{ Ok = $false; Error = 'Remote manifest has no version.' } }
                else { [pscustomobject]@{ Ok = $true; Version = $v } }
            } catch { [pscustomobject]@{ Ok = $false; Error = $_.Exception.Message } }
        }).AddArgument($url)
    $script:UI.UpdateBusy = $true
    $script:UI.UpdateSilent = [bool]$Silent
    $script:UI.UpdateLocalVer = [string]$local.version
    $script:UI.UpdatePS = $ps
    $script:UI.UpdateHandle = $ps.BeginInvoke()
    $t = New-Object System.Windows.Threading.DispatcherTimer
    $t.Interval = [TimeSpan]::FromMilliseconds(150)
    $t.Add_Tick({ Complete-GuiUpdateCheck })
    $script:UI.UpdateTimer = $t
    if (-not $Silent) { Write-Log 'Checking for updates...' 'muted'; $script:UI.UpdateBtn.IsEnabled = $false }
    $t.Start()
}

function Complete-GuiUpdateCheck {
    if (-not $script:UI.UpdateHandle.IsCompleted) { return }
    $script:UI.UpdateTimer.Stop()
    $res = $null; $err = $null
    try { $res = @($script:UI.UpdatePS.EndInvoke($script:UI.UpdateHandle))[0] } catch { $err = $_.Exception.Message }
    try { $script:UI.UpdatePS.Dispose() } catch { }
    $silent = $script:UI.UpdateSilent
    $script:UI.UpdateBusy = $false
    $script:UI.UpdateBtn.IsEnabled = $true

    if ($err -or -not $res -or -not $res.Ok) {
        $detail = if ($err) { $err } elseif ($res) { [string]$res.Error } else { 'Unknown error.' }
        $msg = Get-FriendlyUpdateError $detail
        if ($silent) { Write-Log "Update check skipped: $msg" 'muted' }
        else { Write-Log "Update check failed: $msg" 'warn'; [System.Windows.MessageBox]::Show($msg, 'Update check failed', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning) | Out-Null }
        return
    }

    $remote = [string]$res.Version
    $localV = [string]$script:UI.UpdateLocalVer
    if ((Compare-PassReVersion $remote $localV) -le 0) {
        $script:UI.PendingVersion = $null
        $script:UI.UpdateBtn.Content = 'Check for updates'
        if ($silent) { Write-Log "Up to date (v$localV)." 'muted' }
        else { Write-Log "Up to date (v$localV)." 'ok'; [System.Windows.MessageBox]::Show("You're on the latest version (v$localV).", 'No updates', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null }
        return
    }

    $script:UI.PendingVersion = $remote
    $script:UI.UpdateBtn.Content = "Update to v$remote"
    Write-Log "Update available: v$localV -> v$remote" 'head'
    if (-not $silent) { Invoke-GuiUpdateInstall }
}

function Invoke-GuiUpdateInstall {
    if (-not $script:UI.PendingVersion) { Start-GuiUpdateCheck; return }
    $remote = [string]$script:UI.PendingVersion
    $localV = [string]$script:UI.UpdateLocalVer
    $ans = [System.Windows.MessageBox]::Show("Update available: v$localV -> v$remote.`r`n`r`nInstall now?", 'Update available', [System.Windows.MessageBoxButton]::YesNo, [System.Windows.MessageBoxImage]::Question)
    if ($ans -ne [System.Windows.MessageBoxResult]::Yes) { return }
    try {
        $script:UI.Win.Cursor = [System.Windows.Input.Cursors]::Wait
        $r = Invoke-PassReSelfUpdate -Force
        if ($r.Updated) {
            $script:UI.PendingVersion = $null
            $script:UI.UpdateBtn.Content = 'Check for updates'
            [System.Windows.MessageBox]::Show("Updated to v$($r.Remote). Restart PassRe to use the new version.", 'Update complete', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information) | Out-Null
        }
    } catch {
        Write-Log "Update failed: $($_.Exception.Message)" 'err'
        [System.Windows.MessageBox]::Show((Get-FriendlyUpdateError $_.Exception.Message), 'Update failed', [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error) | Out-Null
    } finally { $script:UI.Win.Cursor = $null }
}

function Build-PassReGui {
    Add-Type -AssemblyName PresentationFramework
    $xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PassRe" Width="560" Height="640" MinWidth="460" MinHeight="560"
        WindowStartupLocation="CenterScreen" Background="#F4F5F7" FontFamily="Segoe UI" FontSize="13">
  <Window.Resources>
    <Style x:Key="Primary" TargetType="Button">
      <Setter Property="Height" Value="42"/><Setter Property="Foreground" Value="White"/>
      <Setter Property="Background" Value="#2F855A"/><Setter Property="FontSize" Value="14"/>
      <Setter Property="FontWeight" Value="SemiBold"/><Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border x:Name="bd" Background="{TemplateBinding Background}" CornerRadius="7">
            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#276749"/></Trigger>
            <Trigger Property="IsEnabled" Value="False"><Setter TargetName="bd" Property="Background" Value="#CBD5D0"/><Setter Property="Foreground" Value="#EDF2F0"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value></Setter>
    </Style>
    <Style x:Key="Ghost" TargetType="Button">
      <Setter Property="Height" Value="42"/><Setter Property="Foreground" Value="#2D3748"/>
      <Setter Property="Background" Value="White"/><Setter Property="FontSize" Value="14"/><Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Template"><Setter.Value>
        <ControlTemplate TargetType="Button">
          <Border x:Name="bd" Background="{TemplateBinding Background}" CornerRadius="7" BorderBrush="#CBD5E0" BorderThickness="1">
            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
          </Border>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="bd" Property="Background" Value="#EDF2F7"/></Trigger>
            <Trigger Property="IsEnabled" Value="False"><Setter Property="Foreground" Value="#A0AEC0"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value></Setter>
    </Style>
    <Style x:Key="Lbl" TargetType="TextBlock"><Setter Property="Foreground" Value="#718096"/><Setter Property="Margin" Value="0,3,0,3"/></Style>
    <Style x:Key="Val" TargetType="TextBlock"><Setter Property="FontWeight" Value="SemiBold"/><Setter Property="HorizontalAlignment" Value="Right"/><Setter Property="Margin" Value="0,3,0,3"/></Style>
  </Window.Resources>

  <Grid Margin="18">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/>
      <RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <StackPanel Grid.Row="0" Margin="0,0,0,14">
      <TextBlock Text="PassRe" FontSize="24" FontWeight="Bold" Foreground="#1A202C"/>
      <TextBlock Text="Remove Excel protection, edit, then restore it exactly." Foreground="#718096" Margin="0,2,0,0"/>
    </StackPanel>

    <Border Grid.Row="1" x:Name="DropZone" Background="White" CornerRadius="10" BorderBrush="#CBD5E0" BorderThickness="1" Padding="16" Margin="0,0,0,14" AllowDrop="True">
      <StackPanel>
        <TextBlock x:Name="FileText" Text="Drop an .xlsx here" FontSize="15" FontWeight="SemiBold" Foreground="#2D3748" TextTrimming="CharacterEllipsis"/>
        <TextBlock x:Name="PathText" Text="or click Browse to choose a file" Foreground="#718096" Margin="0,2,0,10" TextTrimming="CharacterEllipsis"/>
        <Button x:Name="BrowseBtn" Content="Browse&#8230;" Style="{StaticResource Ghost}" Width="110" Height="34" HorizontalAlignment="Left"/>
      </StackPanel>
    </Border>

    <Border Grid.Row="2" Background="White" CornerRadius="10" BorderBrush="#E2E8F0" BorderThickness="1" Padding="14" Margin="0,0,0,14">
      <Grid>
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
        <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/></Grid.RowDefinitions>
        <TextBlock Grid.Row="0" Grid.Column="0" Text="Workbook structure lock" Style="{StaticResource Lbl}"/>
        <TextBlock Grid.Row="0" Grid.Column="1" x:Name="WbVal" Text="&#8212;" Style="{StaticResource Val}"/>
        <TextBlock Grid.Row="1" Grid.Column="0" Text="Password to modify" Style="{StaticResource Lbl}"/>
        <TextBlock Grid.Row="1" Grid.Column="1" x:Name="FsVal" Text="&#8212;" Style="{StaticResource Val}"/>
        <TextBlock Grid.Row="2" Grid.Column="0" Text="Protected worksheets" Style="{StaticResource Lbl}"/>
        <TextBlock Grid.Row="2" Grid.Column="1" x:Name="ShVal" Text="&#8212;" Style="{StaticResource Val}"/>
        <TextBlock Grid.Row="3" Grid.Column="0" Text="Protected chart sheets" Style="{StaticResource Lbl}"/>
        <TextBlock Grid.Row="3" Grid.Column="1" x:Name="ChVal" Text="&#8212;" Style="{StaticResource Val}"/>
        <TextBlock Grid.Row="4" Grid.Column="0" Text="Stash" Style="{StaticResource Lbl}"/>
        <TextBlock Grid.Row="4" Grid.Column="1" x:Name="StVal" Text="&#8212;" Style="{StaticResource Val}"/>
      </Grid>
    </Border>

    <Grid Grid.Row="3" Margin="0,0,0,12">
      <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="10"/><ColumnDefinition Width="*"/><ColumnDefinition Width="10"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
      <Button x:Name="UnprotectBtn" Grid.Column="0" Content="Unprotect" Style="{StaticResource Primary}" IsEnabled="False"/>
      <Button x:Name="ReprotectBtn" Grid.Column="2" Content="Reprotect" Style="{StaticResource Ghost}" IsEnabled="False"/>
      <Button x:Name="ToggleBtn"    Grid.Column="4" Content="Toggle"    Style="{StaticResource Ghost}" IsEnabled="False"/>
    </Grid>

    <StackPanel Grid.Row="4" Orientation="Horizontal" Margin="0,0,0,10">
      <CheckBox x:Name="BackupChk" Content="Keep .bak backup" IsChecked="True" Margin="0,0,18,0"/>
      <CheckBox x:Name="KeepStashChk" Content="Keep stash after reprotect"/>
    </StackPanel>

    <Border Grid.Row="5" Background="#1A202C" CornerRadius="8" Padding="8">
      <TextBox x:Name="LogBox" IsReadOnly="True" Background="Transparent" Foreground="#E2E8F0" BorderThickness="0" FontFamily="Consolas" FontSize="12" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto"/>
    </Border>

    <Grid Grid.Row="6" Margin="0,10,0,0">
      <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/></Grid.ColumnDefinitions>
      <TextBlock x:Name="VersionText" Grid.Column="0" Foreground="#A0AEC0" VerticalAlignment="Center"/>
      <Button x:Name="UpdateBtn" Grid.Column="1" Content="Check for updates" Style="{StaticResource Ghost}" Width="170" Height="30" FontSize="12"/>
    </Grid>
  </Grid>
</Window>
'@
    $win = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$xaml)))
    $script:UI = @{ Win = $win; File = $null }
    foreach ($n in 'DropZone', 'FileText', 'PathText', 'BrowseBtn', 'WbVal', 'FsVal', 'ShVal', 'ChVal', 'StVal', 'UnprotectBtn', 'ReprotectBtn', 'ToggleBtn', 'BackupChk', 'KeepStashChk', 'LogBox', 'VersionText', 'UpdateBtn') { $script:UI[$n] = $win.FindName($n) }

    Set-PassReLogSink {
        param($m, $l)
        $tag = switch ($l) { 'err' { ' !  ' } 'ok' { ' +  ' } 'warn' { ' ~  ' } 'head' { '==> ' } 'muted' { '    ' } default { ' .  ' } }
        $script:UI.LogBox.AppendText("$tag$m`r`n"); $script:UI.LogBox.ScrollToEnd()
    }

    $script:UI.BrowseBtn.Add_Click({ $dlg = New-Object Microsoft.Win32.OpenFileDialog; $dlg.Filter = 'Excel files|*.xlsx;*.xlsm;*.xltx;*.xltm|All files|*.*'; if ($dlg.ShowDialog()) { Set-GuiFile $dlg.FileName } })
    $script:UI.DropZone.Add_DragOver({ param($s, $e) $e.Effects = [System.Windows.DragDropEffects]::Copy; $e.Handled = $true })
    $script:UI.DropZone.Add_Drop({ param($s, $e) $f = $e.Data.GetData([System.Windows.DataFormats]::FileDrop); if ($f -and $f.Length -ge 1) { Set-GuiFile $f[0] } })
    $script:UI.UnprotectBtn.Add_Click({ try { [void](Invoke-PassReUnprotect -Full $script:UI.File -Opt (Get-GuiOpt)) } catch { Write-Log "ERROR: $($_.Exception.Message)" 'err' } Update-GuiStatus })
    $script:UI.ReprotectBtn.Add_Click({ try { [void](Invoke-PassReReprotect -Full $script:UI.File -Opt (Get-GuiOpt)) } catch { Write-Log "ERROR: $($_.Exception.Message)" 'err' } Update-GuiStatus })
    $script:UI.ToggleBtn.Add_Click({
            try {
                $side = Get-SidecarPath -Full $script:UI.File
                if (Test-PassReStashActive $side) { [void](Invoke-PassReReprotect -Full $script:UI.File -Opt (Get-GuiOpt)) } else { [void](Invoke-PassReUnprotect -Full $script:UI.File -Opt (Get-GuiOpt)) }
            } catch { Write-Log "ERROR: $($_.Exception.Message)" 'err' }
            Update-GuiStatus
        })

    $ver = Get-PassReVersion
    $script:UI.VersionText.Text = if ($ver) { "PassRe v$ver" } else { 'PassRe (version unknown)' }
    $script:UI.PendingVersion = $null
    $script:UI.UpdateBusy = $false
    $script:UI.UpdateBtn.Add_Click({ if ($script:UI.PendingVersion) { Invoke-GuiUpdateInstall } else { Start-GuiUpdateCheck } })

    Write-Log 'Ready. Drop a workbook or click Browse.' 'muted'
    return $win
}

function Show-PassReGui {
    param([string]$Preload, [switch]$BuildOnly)
    $win = Build-PassReGui
    Update-GuiStatus   # empty state first, so a preload that fails (not found, a folder) still shows grey dashes
    if ($Preload) { Set-GuiFile $Preload }
    if ($BuildOnly) { return $win }
    if ($script:CheckUpdatesOnLaunch) { Start-GuiUpdateCheck -Silent }
    [void]$win.ShowDialog()
}

function Start-PassReGui {
    param([string]$Preload)
    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -ne 'STA') {
        $exe = Join-Path $PSHOME 'powershell.exe'
        $a = @('-STA', '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$PSCommandPath`"")
        if ($Preload) { $a += "`"$Preload`"" }
        Start-Process -FilePath $exe -ArgumentList $a
        return
    }
    Show-PassReGui -Preload $Preload
}

#endregion

if ($AsLibrary) { return }   # dot-source stops here: engine + GUI functions are now defined

#region ============================ MAIN ======================================

if ($Quiet) { Set-PassReLogSink { param($m, $l) } }

if ($SelfTest) { [void](Show-PassReGui -BuildOnly); Write-Output 'SELFTEST_OK'; return }

$verbs = 'unprotect', 'reprotect', 'toggle', 'status'

# -Gui: the launcher's file branch passes this, so a path that doesn't exist (e.g. one cmd cut at '&') still
# opens the window, which reports "Not found" instead of failing invisibly in a hidden console.
if ($Gui) { $pre = if ($Path) { $Path } else { $Action }; Start-PassReGui -Preload $pre; return }
if (-not $Action) { Start-PassReGui -Preload ''; return }

$verb = $Action.ToLowerInvariant()

if ($verb -eq 'update') {
    $confirm = {
        param($from, $to)
        $ans = Read-Host "Install update v$from -> v$to? [y/N]"
        return ($ans -match '^(?i)y(es)?$')
    }
    try { [void](Invoke-PassReSelfUpdate -CheckOnly:$CheckOnly -Confirm $confirm -Force:$Force) }
    catch { Write-Log (Get-FriendlyUpdateError $_.Exception.Message) 'err'; exit 1 }
    return
}

if ($verbs -contains $verb) {
    if (-not $Path) { throw "Usage: PassRe $verb <file.xlsx>" }
    $opt = @{ Sidecar = $Sidecar; NoBackup = [bool]$NoBackup; KeepSidecar = [bool]$KeepSidecar }
    $full = Resolve-Xlsx -Path $Path -Quiet:$Quiet
    switch ($verb) {
        'unprotect' { [void](Invoke-PassReUnprotect -Full $full -Opt $opt) }
        'reprotect' { [void](Invoke-PassReReprotect -Full $full -Opt $opt) }
        'toggle' {
            $side = Get-SidecarPath -Full $full -Sidecar $Sidecar
            if (Test-PassReStashActive $side) { Write-Log "Stash waiting to be restored -> reprotecting." 'muted'; [void](Invoke-PassReReprotect -Full $full -Opt $opt) }
            else { Write-Log "No stash waiting -> unprotecting." 'muted'; [void](Invoke-PassReUnprotect -Full $full -Opt $opt) }
        }
        'status' {
            $s = Get-PassReStatus -Full $full -Sidecar $Sidecar
            Write-Log "File              : $($s.File)"
            Write-Log ("Workbook lock     : " + $(if ($s.WorkbookProtected) { 'PROTECTED' } else { 'none' })) $(if ($s.WorkbookProtected) { 'warn' } else { 'ok' })
            Write-Log ("Password to modify: " + $(if ($s.PasswordToModify) { 'SET' } else { 'none' })) $(if ($s.PasswordToModify) { 'warn' } else { 'ok' })
            Write-Log ("Protected sheets  : $($s.ProtectedSheets)") $(if ($s.ProtectedSheets) { 'warn' } else { 'ok' })
            Write-Log ("Protected charts  : $($s.ProtectedChartSheets)") $(if ($s.ProtectedChartSheets) { 'warn' } else { 'ok' })
            if ($s.StashActive) { Write-Log "Stash             : present ($($s.StashCount) element(s), made $($s.StashCreated))" 'head' }
            elseif ($s.HasStash) { Write-Log "Stash             : kept, already restored ($($s.StashCount) element(s), restored $($s.StashRestored))" 'muted' }
            else { Write-Log "Stash             : none" 'muted' }
            Write-Log "Next              : $($s.Next)" 'muted'
        }
    }
    return
}

# Not a verb: a dropped/opened file -> GUI preloaded.
if (Test-Path -LiteralPath $Action) { Start-PassReGui -Preload $Action; return }

throw "Unknown action '$Action'. Use: unprotect|reprotect|toggle|status <file.xlsx>, update [-CheckOnly] [-Force], a file to open the GUI on it, or no arguments for the GUI."

#endregion

# Keep this the last line: the self-updater refuses a downloaded PassRe.ps1 that doesn't end with it.
# PassRe:EOF
