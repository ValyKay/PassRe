<#
.SYNOPSIS
  PassRe - strip Excel sheet/chart/workbook protection, then put it back exactly.

.DESCRIPTION
  One file, three ways to use it:
    * No arguments .......... opens the GUI.
    * A file path ........... opens the GUI with that workbook loaded (drag-drop).
    * A verb + file ......... command line: unprotect | reprotect | toggle | status

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
    [switch]$SelfTest,     # build the GUI headless and exit (used by tests)
    [switch]$AsLibrary     # dot-source target: define functions, then return
)

$ErrorActionPreference = 'Stop'
try {
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
} catch { }

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

function Get-SidecarPath { param([string]$Full, [string]$Sidecar) if ($Sidecar) { return $Sidecar } return "$Full.passre" }

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
    $zip = [System.IO.Compression.ZipFile]::Open($Full, [System.IO.Compression.ZipArchiveMode]::Read)
    try {
        $wb = Read-EntryString $zip 'xl/workbook.xml'
        if ($wb) {
            if ([regex]::IsMatch($wb, $script:Rx.Workbook)) { $wbProtected = $true }
            if ([regex]::IsMatch($wb, $script:Rx.FileSharing)) { $fileShare = $true }
        }
        foreach ($n in Get-WorksheetEntryNames $zip) { $x = Read-EntryString $zip $n; if ($x -and [regex]::IsMatch($x, $script:Rx.Sheet)) { $sheetCount++ } }
        foreach ($n in Get-ChartsheetEntryNames $zip) { $x = Read-EntryString $zip $n; if ($x -and [regex]::IsMatch($x, $script:Rx.Sheet)) { $chartCount++ } }
    } finally { $zip.Dispose() }

    $stashCount = 0; $stashCreated = $null; $hasStash = Test-Path -LiteralPath $side
    if ($hasStash) { try { $d = Get-Content -LiteralPath $side -Raw | ConvertFrom-Json; $stashCount = @($d.entries).Count; $stashCreated = $d.created } catch { } }
    $anyProt = ($sheetCount -gt 0 -or $chartCount -gt 0 -or $wbProtected -or $fileShare)
    $next = if ($hasStash) { 'reprotect' } elseif ($anyProt) { 'unprotect' } else { 'none' }

    [pscustomobject]@{
        File = $Full; Sidecar = $side
        ProtectedSheets = $sheetCount; ProtectedChartSheets = $chartCount
        WorkbookProtected = $wbProtected; PasswordToModify = $fileShare
        HasStash = $hasStash; StashCount = $stashCount; StashCreated = $stashCreated
        AnyProtection = $anyProt; Next = $next
    }
}

function Invoke-PassReUnprotect {
    param([string]$Full, [hashtable]$Opt = @{})
    $side = Get-SidecarPath -Full $Full -Sidecar $Opt.Sidecar
    if ((Test-Path -LiteralPath $side) -and -not $Opt.Force) { throw "A stash already exists: $side`nThe file looks already unprotected. Use reprotect, or pass Force to re-stash." }

    $backup = $null
    if (-not $Opt.NoBackup) { $backup = "$Full.bak"; Copy-Item -LiteralPath $Full -Destination $backup -Force; Write-Log "Backup of the original  -> $backup" 'muted' }

    $records = New-Object System.Collections.Generic.List[object]
    $zip = [System.IO.Compression.ZipFile]::Open($Full, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $wb = Read-EntryString $zip 'xl/workbook.xml'
        if ($wb) {
            $changed = $false
            $m = [regex]::Match($wb, $script:Rx.FileSharing)
            if ($m.Success) { $records.Add([pscustomobject]@{ entry = 'xl/workbook.xml'; kind = 'fileSharing'; element = $m.Value }); $wb = $wb.Remove($m.Index, $m.Length); $changed = $true; Write-Log "Removed file-level 'password to modify'." 'ok' }
            $m = [regex]::Match($wb, $script:Rx.Workbook)
            if ($m.Success) { $records.Add([pscustomobject]@{ entry = 'xl/workbook.xml'; kind = 'workbook'; element = $m.Value }); $wb = $wb.Remove($m.Index, $m.Length); $changed = $true; Write-Log "Removed workbook structure protection." 'ok' }
            if ($changed) { Write-EntryString $zip 'xl/workbook.xml' $wb }
        }
        foreach ($name in Get-WorksheetEntryNames $zip) {
            $x = Read-EntryString $zip $name; if (-not $x) { continue }
            $m = [regex]::Match($x, $script:Rx.Sheet)
            if ($m.Success) { $records.Add([pscustomobject]@{ entry = $name; kind = 'sheet'; element = $m.Value }); Write-EntryString $zip $name ($x.Remove($m.Index, $m.Length)); Write-Log "Removed sheet protection -> $name" 'ok' }
        }
        foreach ($name in Get-ChartsheetEntryNames $zip) {
            $x = Read-EntryString $zip $name; if (-not $x) { continue }
            $m = [regex]::Match($x, $script:Rx.Sheet)
            if ($m.Success) { $records.Add([pscustomobject]@{ entry = $name; kind = 'chartsheet'; element = $m.Value }); Write-EntryString $zip $name ($x.Remove($m.Index, $m.Length)); Write-Log "Removed chart-sheet protection -> $name" 'ok' }
        }
    } finally { $zip.Dispose() }

    if ($records.Count -eq 0) {
        Write-Log "No sheet/chart/workbook protection found - nothing to remove." 'warn'
        Write-Log "(A password required to OPEN the file is encryption, which PassRe can't strip.)" 'muted'
        if ($backup -and (Test-Path -LiteralPath $backup)) { Remove-Item -LiteralPath $backup -Force }
        return [pscustomobject]@{ Removed = 0; Stash = $null; Backup = $null }
    }

    $stash = [pscustomobject]@{ tool = 'PassRe'; version = 2; created = (Get-Date).ToString('o'); originalFile = (Split-Path -Leaf $Full); entries = $records.ToArray() }
    $stash | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $side -Encoding UTF8
    Write-Log "Unprotected. Stashed $($records.Count) element(s) -> $side" 'head'
    return [pscustomobject]@{ Removed = $records.Count; Stash = $side; Backup = $backup }
}

function Invoke-PassReReprotect {
    param([string]$Full, [hashtable]$Opt = @{})
    $side = Get-SidecarPath -Full $Full -Sidecar $Opt.Sidecar
    if (-not (Test-Path -LiteralPath $side)) { throw "No stash at: $side`nCan't restore the original protection without it. (Run unprotect first.)" }
    $recs = @((Get-Content -LiteralPath $side -Raw | ConvertFrom-Json).entries)
    if ($recs.Count -eq 0) { throw "Stash is empty: $side" }

    $restored = 0
    $zip = [System.IO.Compression.ZipFile]::Open($Full, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        foreach ($rec in $recs) {
            $x = Read-EntryString $zip $rec.entry
            if ($null -eq $x) { Write-Log "Skipped (part gone): $($rec.entry)" 'warn'; continue }
            switch ($rec.kind) {
                'fileSharing' { Write-EntryString $zip $rec.entry (Add-FileSharing         $x $rec.element); Write-Log "Restored 'password to modify'." 'ok' }
                'workbook'    { Write-EntryString $zip $rec.entry (Add-WorkbookProtection   $x $rec.element); Write-Log "Restored workbook structure protection." 'ok' }
                'chartsheet'  { Write-EntryString $zip $rec.entry (Add-ChartsheetProtection $x $rec.element); Write-Log "Restored chart-sheet protection -> $($rec.entry)" 'ok' }
                default       { Write-EntryString $zip $rec.entry (Add-SheetProtection      $x $rec.element); Write-Log "Restored sheet protection -> $($rec.entry)" 'ok' }
            }
            $restored++
        }
    } finally { $zip.Dispose() }

    if (-not $Opt.KeepSidecar) { Remove-Item -LiteralPath $side -Force; Write-Log "Stash cleared." 'muted' } else { Write-Log "Stash kept: $side" 'muted' }
    Write-Log "Reprotected: $restored element(s) put back exactly as they were." 'head'
    return [pscustomobject]@{ Restored = $restored }
}

#endregion

#region ============================= GUI ======================================

$script:UI = @{}

function Set-GuiVal { param($Tb, [string]$Text, [bool]$Warn) $Tb.Text = $Text; $Tb.Foreground = if ($Warn) { '#C05621' } else { '#2F855A' } }

function Update-GuiStatus {
    $g = $script:UI
    if (-not $g.File) {
        $g.FileText.Text = 'Drop an .xlsx here'; $g.PathText.Text = 'or click Browse to choose a file'
        foreach ($v in $g.WbVal, $g.FsVal, $g.ShVal, $g.ChVal, $g.StVal) { $v.Text = [string][char]0x2014; $v.Foreground = '#718096' }
        $g.UnprotectBtn.IsEnabled = $false; $g.ReprotectBtn.IsEnabled = $false; $g.ToggleBtn.IsEnabled = $false
        return
    }
    try { $s = Get-PassReStatus -Full $g.File } catch { Write-Log "ERROR: $($_.Exception.Message)" 'err'; return }
    $g.FileText.Text = (Split-Path -Leaf $g.File); $g.PathText.Text = $g.File
    Set-GuiVal $g.WbVal $(if ($s.WorkbookProtected) { 'PROTECTED' } else { 'none' }) $s.WorkbookProtected
    Set-GuiVal $g.FsVal $(if ($s.PasswordToModify) { 'SET' } else { 'none' }) $s.PasswordToModify
    Set-GuiVal $g.ShVal "$($s.ProtectedSheets)" ($s.ProtectedSheets -gt 0)
    Set-GuiVal $g.ChVal "$($s.ProtectedChartSheets)" ($s.ProtectedChartSheets -gt 0)
    if ($s.HasStash) { Set-GuiVal $g.StVal "present ($($s.StashCount))" $false } else { $g.StVal.Text = 'none'; $g.StVal.Foreground = '#718096' }
    $g.UnprotectBtn.IsEnabled = ($s.AnyProtection -and -not $s.HasStash)
    $g.ReprotectBtn.IsEnabled = $s.HasStash
    $g.ToggleBtn.IsEnabled = ($s.AnyProtection -or $s.HasStash)
}

function Set-GuiFile {
    param([string]$Path)
    if (-not $Path) { return }
    if (-not (Test-Path -LiteralPath $Path)) { Write-Log "Not found: $Path" 'err'; return }
    $script:UI.File = (Get-Item -LiteralPath $Path).FullName
    Write-Log "Loaded: $(Split-Path -Leaf $script:UI.File)" 'head'
    Update-GuiStatus
}

function Get-GuiOpt { @{ NoBackup = (-not $script:UI.BackupChk.IsChecked); KeepSidecar = [bool]$script:UI.KeepStashChk.IsChecked } }

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
  </Grid>
</Window>
'@
    $win = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader ([xml]$xaml)))
    $script:UI = @{ Win = $win; File = $null }
    foreach ($n in 'DropZone', 'FileText', 'PathText', 'BrowseBtn', 'WbVal', 'FsVal', 'ShVal', 'ChVal', 'StVal', 'UnprotectBtn', 'ReprotectBtn', 'ToggleBtn', 'BackupChk', 'KeepStashChk', 'LogBox') { $script:UI[$n] = $win.FindName($n) }

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
                if (Test-Path -LiteralPath $side) { [void](Invoke-PassReReprotect -Full $script:UI.File -Opt (Get-GuiOpt)) } else { [void](Invoke-PassReUnprotect -Full $script:UI.File -Opt (Get-GuiOpt)) }
            } catch { Write-Log "ERROR: $($_.Exception.Message)" 'err' }
            Update-GuiStatus
        })

    Write-Log 'Ready. Drop a workbook or click Browse.' 'muted'
    return $win
}

function Show-PassReGui {
    param([string]$Preload, [switch]$BuildOnly)
    $win = Build-PassReGui
    if ($Preload) { Set-GuiFile $Preload } else { Update-GuiStatus }
    if ($BuildOnly) { return $win }
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

if ($Gui) { $pre = if ($Action -and (Test-Path -LiteralPath $Action)) { $Action } elseif ($Path) { $Path } else { '' }; Start-PassReGui -Preload $pre; return }
if (-not $Action) { Start-PassReGui -Preload ''; return }

$verb = $Action.ToLowerInvariant()
if ($verbs -contains $verb) {
    if (-not $Path) { throw "Usage: PassRe $verb <file.xlsx>" }
    $opt = @{ Sidecar = $Sidecar; NoBackup = [bool]$NoBackup; KeepSidecar = [bool]$KeepSidecar; Force = [bool]$Force }
    $full = Resolve-Xlsx -Path $Path -Quiet:$Quiet
    switch ($verb) {
        'unprotect' { [void](Invoke-PassReUnprotect -Full $full -Opt $opt) }
        'reprotect' { [void](Invoke-PassReReprotect -Full $full -Opt $opt) }
        'toggle' {
            $side = Get-SidecarPath -Full $full -Sidecar $Sidecar
            if (Test-Path -LiteralPath $side) { Write-Log "Stash present -> reprotecting." 'muted'; [void](Invoke-PassReReprotect -Full $full -Opt $opt) }
            else { Write-Log "No stash -> unprotecting." 'muted'; [void](Invoke-PassReUnprotect -Full $full -Opt $opt) }
        }
        'status' {
            $s = Get-PassReStatus -Full $full -Sidecar $Sidecar
            Write-Log "File              : $($s.File)"
            Write-Log ("Workbook lock     : " + $(if ($s.WorkbookProtected) { 'PROTECTED' } else { 'none' })) $(if ($s.WorkbookProtected) { 'warn' } else { 'ok' })
            Write-Log ("Password to modify: " + $(if ($s.PasswordToModify) { 'SET' } else { 'none' })) $(if ($s.PasswordToModify) { 'warn' } else { 'ok' })
            Write-Log ("Protected sheets  : $($s.ProtectedSheets)") $(if ($s.ProtectedSheets) { 'warn' } else { 'ok' })
            Write-Log ("Protected charts  : $($s.ProtectedChartSheets)") $(if ($s.ProtectedChartSheets) { 'warn' } else { 'ok' })
            if ($s.HasStash) { Write-Log "Stash             : present ($($s.StashCount) element(s), made $($s.StashCreated))" 'head' } else { Write-Log "Stash             : none" 'muted' }
            Write-Log "Next              : $($s.Next)" 'muted'
        }
    }
    return
}

# Not a verb: a dropped/opened file -> GUI preloaded.
if (Test-Path -LiteralPath $Action) { Start-PassReGui -Preload $Action; return }

throw "Unknown action '$Action'. Use: unprotect|reprotect|toggle|status <file.xlsx>, give a file to open the GUI on it, or run with no arguments for the GUI."

#endregion
