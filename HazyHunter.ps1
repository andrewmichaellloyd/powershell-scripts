#requires -version 5.1


# --- TLS hardening for Windows PowerShell 5.1 (common fix for modern HTTPS sites)
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}

# Suppress built-in console progress (e.g. "Reading web response...")
$global:ProgressPreference = 'SilentlyContinue'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ------------------------
# GUI
# ------------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Andy's IPA Hunter"
$form.Size = New-Object System.Drawing.Size(1050, 980)
$form.StartPosition = "CenterScreen"

$lbl = New-Object System.Windows.Forms.Label
$lbl.Text = "Find latest IPAs across selected shops."
$lbl.AutoSize = $true
$lbl.Location = New-Object System.Drawing.Point(12, 12)

# Elapsed time (NEW)
$lblElapsed = New-Object System.Windows.Forms.Label
$lblElapsed.Text = "Elapsed: 00:00"
$lblElapsed.AutoSize = $true
$lblElapsed.Location = New-Object System.Drawing.Point(890, 12)

$grp = New-Object System.Windows.Forms.GroupBox
$grp.Text = "Sources"
$grp.Location = New-Object System.Drawing.Point(12, 40)
$grp.Size = New-Object System.Drawing.Size(1018, 125)

# DPI-safe, clean 3-row layout using TableLayoutPanel
$tlpSources = New-Object System.Windows.Forms.TableLayoutPanel
$tlpSources.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpSources.Padding = New-Object System.Windows.Forms.Padding(10,18,10,10)
$tlpSources.ColumnCount = 4
$tlpSources.RowCount = 4
$tlpSources.GrowStyle = [System.Windows.Forms.TableLayoutPanelGrowStyle]::FixedSize

$tlpSources.ColumnStyles.Clear()
for ($i=0; $i -lt 4; $i++) {
  $tlpSources.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 25)))
}
$tlpSources.RowStyles.Clear()
for ($i=0; $i -lt 4; $i++) {
  $tlpSources.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 25)))
}

function New-SourceCheckBox([string]$text) {
  $cb = New-Object System.Windows.Forms.CheckBox
  $cb.Text = $text
  $cb.Checked = $true
  $cb.AutoSize = $true
  $cb.Dock = [System.Windows.Forms.DockStyle]::Fill
  $cb.Margin = New-Object System.Windows.Forms.Padding(6,2,6,2)
  return $cb
}

$chkVerdant          = New-SourceCheckBox "Verdant (FAST)"
$chkCloudwater       = New-SourceCheckBox "Cloudwater (FAST)"
$chkPollys           = New-SourceCheckBox "Polly's"
$chkLeftField        = New-SourceCheckBox "Left Field Beer"

$chkGhostWhale       = New-SourceCheckBox "Ghost Whale (FAST)"
$chkHopBurnsBlack    = New-SourceCheckBox "Hop Burns & Black (FAST)"
$chkRadBeer          = New-SourceCheckBox "RAD Beer (FAST)"
$chkTremblingMadness = New-SourceCheckBox "Trembling Madness"

$chkBeerMerchants    = New-SourceCheckBox "Beer Merchants"
$chkBeerRitz         = New-SourceCheckBox "Beer Ritz"
$chkShinyBrewery     = New-SourceCheckBox "Shiny Brewery (FAST)"
$chkBeerGuerrilla    = New-SourceCheckBox "Beer Guerrilla (FAST)"
$chkMakemake         = New-SourceCheckBox "Makemake (FAST)"
$chkVaultCity        = New-SourceCheckBox "Vault City Brewing (FAST)"


# Row 1
$tlpSources.Controls.Add($chkVerdant,          0, 0)
$tlpSources.Controls.Add($chkCloudwater,       1, 0)
$tlpSources.Controls.Add($chkPollys,           2, 0)
$tlpSources.Controls.Add($chkLeftField,        3, 0)

# Row 2
$tlpSources.Controls.Add($chkGhostWhale,       0, 1)
$tlpSources.Controls.Add($chkHopBurnsBlack,    1, 1)
$tlpSources.Controls.Add($chkRadBeer,          2, 1)
$tlpSources.Controls.Add($chkTremblingMadness, 3, 1)

# Row 3
$tlpSources.Controls.Add($chkBeerMerchants,    0, 2)
$tlpSources.Controls.Add($chkBeerRitz,         1, 2)
$tlpSources.Controls.Add($chkShinyBrewery,     2, 2)
$tlpSources.Controls.Add($chkBeerGuerrilla,    3, 2)

# Row 4
$tlpSources.Controls.Add($chkMakemake,         0, 3)
$tlpSources.Controls.Add($chkVaultCity,        1, 3)

$grp.Controls.Add($tlpSources)

$lblKw = New-Object System.Windows.Forms.Label
$lblKw.Text = "Keywords (comma-separated):"
$lblKw.AutoSize = $true
$lblKw.Location = New-Object System.Drawing.Point(12, 165)

$txtKeywords = New-Object System.Windows.Forms.TextBox
$txtKeywords.Location = New-Object System.Drawing.Point(12, 185)
$txtKeywords.Size = New-Object System.Drawing.Size(1018, 20)
$txtKeywords.Text = "triple, quad, tipa, quipa"

# ------------------------
# Options group (TableLayoutPanel) - FIXED clipping
# ------------------------
$grpOptions = New-Object System.Windows.Forms.GroupBox
$grpOptions.Text = "Options"
$grpOptions.Location = New-Object System.Drawing.Point(12, 215)
$grpOptions.Size = New-Object System.Drawing.Size(1018, 70)   # slightly taller for DPI

$tlpOptions = New-Object System.Windows.Forms.TableLayoutPanel
$tlpOptions.Dock = [System.Windows.Forms.DockStyle]::Fill
$tlpOptions.Padding = New-Object System.Windows.Forms.Padding(10,18,10,10)
$tlpOptions.RowCount = 1
$tlpOptions.ColumnCount = 6
$tlpOptions.GrowStyle = [System.Windows.Forms.TableLayoutPanelGrowStyle]::FixedSize
$tlpOptions.AutoSize = $false

# Use Percent column for the first checkbox so text can take space (prevents clipping),
# then autos + fixed widths for numeric fields.
$tlpOptions.ColumnStyles.Clear()
$tlpOptions.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 38))) | Out-Null  # Exclude OOS
$tlpOptions.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 22))) | Out-Null  # Limit per vendor
$tlpOptions.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null      # label
$tlpOptions.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 80))) | Out-Null  # num
$tlpOptions.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize))) | Out-Null      # label
$tlpOptions.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 90))) | Out-Null  # num

function Set-StdMargin($ctl) { $ctl.Margin = New-Object System.Windows.Forms.Padding(6,2,12,2) }

$chkExcludeOOS = New-Object System.Windows.Forms.CheckBox
$chkExcludeOOS.Text = "Exclude out of stock (recommended)"
$chkExcludeOOS.Checked = $true
$chkExcludeOOS.AutoSize = $true
$chkExcludeOOS.Dock = [System.Windows.Forms.DockStyle]::Fill
$chkExcludeOOS.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
Set-StdMargin $chkExcludeOOS

$chkLimitPerVendor = New-Object System.Windows.Forms.CheckBox
$chkLimitPerVendor.Text = "Limit results per vendor"
$chkLimitPerVendor.Checked = $true
$chkLimitPerVendor.AutoSize = $true
$chkLimitPerVendor.Dock = [System.Windows.Forms.DockStyle]::Fill
$chkLimitPerVendor.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
Set-StdMargin $chkLimitPerVendor

$lblLimit = New-Object System.Windows.Forms.Label
$lblLimit.Text = "Max per vendor:"
$lblLimit.AutoSize = $true
$lblLimit.Anchor = [System.Windows.Forms.AnchorStyles]::Left
Set-StdMargin $lblLimit

$numLimit = New-Object System.Windows.Forms.NumericUpDown
$numLimit.Minimum = 1
$numLimit.Maximum = 50
$numLimit.Value = 20
$numLimit.DecimalPlaces = 0
$numLimit.Anchor = [System.Windows.Forms.AnchorStyles]::Left
$numLimit.Width = 70
$numLimit.Enabled = $chkLimitPerVendor.Checked
Set-StdMargin $numLimit

$lblMinAbv = New-Object System.Windows.Forms.Label
$lblMinAbv.Text = "Min ABV:"
$lblMinAbv.AutoSize = $true
$lblMinAbv.Anchor = [System.Windows.Forms.AnchorStyles]::Left
Set-StdMargin $lblMinAbv

$numMinAbv = New-Object System.Windows.Forms.NumericUpDown
$numMinAbv.Minimum = 0
$numMinAbv.Maximum = 20
$numMinAbv.DecimalPlaces = 1
$numMinAbv.Increment = 0.1
$numMinAbv.Value = 9.0
$numMinAbv.Anchor = [System.Windows.Forms.AnchorStyles]::Left
$numMinAbv.Width = 80
Set-StdMargin $numMinAbv

$chkLimitPerVendor.Add_CheckedChanged({
  $numLimit.Enabled = $chkLimitPerVendor.Checked
})

$tlpOptions.Controls.Add($chkExcludeOOS,     0, 0)
$tlpOptions.Controls.Add($chkLimitPerVendor,1, 0)
$tlpOptions.Controls.Add($lblLimit,         2, 0)
$tlpOptions.Controls.Add($numLimit,         3, 0)
$tlpOptions.Controls.Add($lblMinAbv,        4, 0)
$tlpOptions.Controls.Add($numMinAbv,        5, 0)

$grpOptions.Controls.Add($tlpOptions)

# ------------------------
# Buttons + status
# ------------------------
$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Scan Now"
$btnScan.Location = New-Object System.Drawing.Point(12, 295)
$btnScan.Size = New-Object System.Drawing.Size(110, 30)

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = "Stop Scan"
$btnStop.Location = New-Object System.Drawing.Point(130, 295)
$btnStop.Size = New-Object System.Drawing.Size(110, 30)
$btnStop.Enabled = $false

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export CSV"
$btnExport.Location = New-Object System.Drawing.Point(248, 295)
$btnExport.Size = New-Object System.Drawing.Size(110, 30)
$btnExport.Enabled = $false

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text = "Open Link"
$btnOpen.Location = New-Object System.Drawing.Point(366, 295)
$btnOpen.Size = New-Object System.Drawing.Size(110, 30)
$btnOpen.Enabled = $false

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(488, 300)
$progress.Size = New-Object System.Drawing.Size(542, 20)
$progress.Minimum = 0
$progress.Maximum = 100
$progress.Value = 0

$status = New-Object System.Windows.Forms.Label
$status.AutoSize = $true
$status.Location = New-Object System.Drawing.Point(12, 335)
$status.Text = "Ready."

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(12, 350)
$grid.Size = New-Object System.Drawing.Size(1018, 550)
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.SelectionMode = "FullRowSelect"
$grid.MultiSelect = $false
$grid.AutoSizeColumnsMode = "Fill"
$grid.RowTemplate.Height = 28
$grid.AutoGenerateColumns = $false

# Data table (price removed)
$table = New-Object System.Data.DataTable
@("Source","Name","ABV","InStock","Url","SeenAtUtc") | ForEach-Object { [void]$table.Columns.Add($_) }

function New-TextCol([string]$name, [string]$header) {
  $c = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
  $c.Name = $name
  $c.HeaderText = $header
  $c.DataPropertyName = $name
  $c.ReadOnly = $true
  return $c
}

$grid.Columns.Add((New-TextCol "Source" "Source")) | Out-Null
$grid.Columns.Add((New-TextCol "Name" "Name")) | Out-Null
$grid.Columns.Add((New-TextCol "ABV" "ABV")) | Out-Null
$grid.Columns.Add((New-TextCol "InStock" "Stock")) | Out-Null

$grid.Columns["Source"].AutoSizeMode   = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
$grid.Columns["ABV"].AutoSizeMode      = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
$grid.Columns["InStock"].AutoSizeMode  = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells
$grid.Columns["Name"].AutoSizeMode     = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$grid.Columns["Name"].FillWeight       = 80

$urlCol = New-Object System.Windows.Forms.DataGridViewLinkColumn
$urlCol.Name = "Url"
$urlCol.HeaderText = "Url"
$urlCol.DataPropertyName = "Url"
$urlCol.TrackVisitedState = $true
$urlCol.LinkBehavior = [System.Windows.Forms.LinkBehavior]::SystemDefault
$urlCol.UseColumnTextForLinkValue = $false
$urlCol.ReadOnly = $true
$grid.Columns.Add($urlCol) | Out-Null

$grid.Columns.Add((New-TextCol "SeenAtUtc" "Seen (UTC)")) | Out-Null
$grid.Columns["SeenAtUtc"].AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::AllCells

$grid.DataSource = $table.DefaultView

$grid.add_CellContentClick({
  param($sender, $e)
  if ($e.RowIndex -lt 0) { return }
  if ($grid.Columns[$e.ColumnIndex].Name -ne "Url") { return }
  $u = $grid.Rows[$e.RowIndex].Cells["Url"].Value
  if ([string]::IsNullOrWhiteSpace($u)) { return }
  Start-Process $u
})

# Timers
$timer = New-Object System.Windows.Forms.Timer     # polls job output
$timer.Interval = 600

$timerElapsed = New-Object System.Windows.Forms.Timer  # elapsed time tick (NEW)
$timerElapsed.Interval = 500

# State
$global:ScanJob = $null
$global:SeenOutputCount = 0
$global:ResultBuffer = New-Object System.Collections.Generic.List[object]
$global:ScanStartUtc = $null

function Update-ElapsedLabel {
  if ($null -eq $global:ScanStartUtc) { $lblElapsed.Text = "Elapsed: 00:00"; return }
  $elapsed = (Get-Date).ToUniversalTime() - $global:ScanStartUtc
  if ($elapsed.TotalSeconds -lt 0) { $elapsed = [TimeSpan]::Zero }
  $lblElapsed.Text = "Elapsed: " + $elapsed.ToString("mm\:ss")
}

$timerElapsed.Add_Tick({ Update-ElapsedLabel })

function Reset-UiState {
  param([string]$Message = "Ready.")
  try { $timer.Stop() } catch {}
  try { $timerElapsed.Stop() } catch {}

  $global:ScanStartUtc = $null
  Update-ElapsedLabel

  $table.Rows.Clear()
  $global:ResultBuffer = New-Object System.Collections.Generic.List[object]
  $global:SeenOutputCount = 0

  $progress.Value = 0
  $status.Text = $Message

  $btnScan.Enabled = $true
  $btnStop.Enabled = $false
  $btnExport.Enabled = $false
  $btnOpen.Enabled = ($grid.SelectedRows.Count -gt 0)

  if ($global:ScanJob) {
    try { Remove-Job -Job $global:ScanJob -Force -ErrorAction SilentlyContinue } catch {}
    $global:ScanJob = $null
  }
}

$grid.Add_SelectionChanged({
  $btnOpen.Enabled = ($grid.SelectedRows.Count -gt 0)
})

$btnOpen.Add_Click({
  if ($grid.SelectedRows.Count -eq 0) { return }
  $url = $grid.SelectedRows[0].Cells["Url"].Value
  if ([string]::IsNullOrWhiteSpace($url)) { return }
  Start-Process $url
})

$btnExport.Add_Click({
  $dlg = New-Object System.Windows.Forms.SaveFileDialog
  $dlg.Filter = "CSV (*.csv)|*.csv"
  $dlg.FileName = "hazy_hunter_results.csv"
  if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

  $path = $dlg.FileName
  $rows = foreach ($drv in $table.DefaultView) {
    $r = $drv.Row
    [pscustomobject]@{
      Source    = $r["Source"]
      Name      = $r["Name"]
      ABV       = $r["ABV"]
      InStock   = $r["InStock"]
      Url       = $r["Url"]
      SeenAtUtc = $r["SeenAtUtc"]
    }
  }

  $rows | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
  $status.Text = "Exported: $path"
})

$btnStop.Add_Click({
  if (-not $global:ScanJob) { return }

  $status.Text = "Stopping scan..."
  try { $timer.Stop() } catch {}
  try { $timerElapsed.Stop() } catch {}

  try { Stop-Job -Job $global:ScanJob -Force -ErrorAction SilentlyContinue } catch {}
  try { Remove-Job -Job $global:ScanJob -Force -ErrorAction SilentlyContinue } catch {}

  $global:ScanJob = $null
  $global:ResultBuffer = New-Object System.Collections.Generic.List[object]
  $global:SeenOutputCount = 0

  $progress.Value = 0
  $status.Text = "Scan stopped."
  $btnScan.Enabled = $true
  $btnStop.Enabled = $false
  $btnExport.Enabled = $false
  $btnOpen.Enabled = ($grid.SelectedRows.Count -gt 0)

  # Freeze elapsed display at stop time
  Update-ElapsedLabel
})

# ------------------------
# Job script (runs in background process)
# ------------------------
$jobScript = {
  param(
    [hashtable]$SelectedSources,
    [string[]]$Keywords,
    [bool]$ExcludeOutOfStock,
    [bool]$LimitPerVendor,
    [int]$MaxPerVendor,
    [double]$MinAbv
  )

  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch {}
  $ProgressPreference = 'SilentlyContinue'

  $UserAgent = "HazyHunterPS/2.3 (personal use)"
  $TimeoutSec = 20
  $DelayMs = 250
  $StrongHints = @("tipa","qipa","quipa","triple","quad")
  $TipaMinAbv = [double]$MinAbv
  $QipaMinAbv = [double]$MinAbv
  $MaxAbv = 15.0

  $vendorCount = @{
    "Verdant (FAST)"           = 0
    "Cloudwater (FAST)"        = 0
    "Polly's"                  = 0
    "Left Field Beer"          = 0
    "Ghost Whale (FAST)"       = 0
    "Hop Burns & Black (FAST)" = 0
    "RAD Beer (FAST)"          = 0
    "Trembling Madness"        = 0
    "Beer Merchants"           = 0
    "Beer Ritz"                = 0
    "Shiny Brewery (FAST)"     = 0
    "Beer Guerrilla (FAST)"    = 0
    "Makemake (FAST)"         = 0
    "Vault City Brewing (FAST)" = 0
  }

  $resultSeen = New-Object System.Collections.Generic.HashSet[string]

  function Normalize-Url([string]$u) {
    if ([string]::IsNullOrWhiteSpace($u)) { return "" }
    try {
      $uri = [Uri]$u
      $path = $uri.AbsolutePath.TrimEnd("/")
      return ($uri.Scheme + "://" + $uri.Host.ToLowerInvariant() + $path).ToLowerInvariant()
    } catch {
      return ($u.Trim().TrimEnd("/") -replace "\?.*$","").ToLowerInvariant()
    }
  }

  function Vendor-LimitReached([string]$Vendor) {
    if (-not $LimitPerVendor) { return $false }
    if (-not $vendorCount.ContainsKey($Vendor)) { return $false }
    return ($vendorCount[$Vendor] -ge $MaxPerVendor)
  }

  function Inc-Vendor([string]$Vendor) {
    if ($vendorCount.ContainsKey($Vendor)) { $vendorCount[$Vendor]++ }
  }

  function All-VendorsSatisfied([hashtable]$Selected) {
    if (-not $LimitPerVendor) { return $false }
    $need = @()
    if ($Selected.Verdant)           { $need += "Verdant (FAST)" }
    if ($Selected.Cloudwater)        { $need += "Cloudwater (FAST)" }
    if ($Selected.Pollys)            { $need += "Polly's" }
    if ($Selected.LeftField)         { $need += "Left Field Beer" }
    if ($Selected.GhostWhale)        { $need += "Ghost Whale (FAST)" }
    if ($Selected.HopBurnsBlack)     { $need += "Hop Burns & Black (FAST)" }
    if ($Selected.RadBeer)           { $need += "RAD Beer (FAST)" }
    if ($Selected.TremblingMadness)  { $need += "Trembling Madness" }
    if ($Selected.BeerMerchants)     { $need += "Beer Merchants" }
    if ($Selected.BeerRitz)          { $need += "Beer Ritz" }
    if ($Selected.ShinyBrewery)     { $need += "Shiny Brewery (FAST)" }
    if ($Selected.BeerGuerrilla)    { $need += "Beer Guerrilla (FAST)" }
    if ($Selected.Makemake)         { $need += "Makemake (FAST)" }
    if ($Selected.VaultCity)        { $need += "Vault City Brewing (FAST)" }

    if ($need.Count -eq 0) { return $false }
    foreach ($v in $need) {
      if ($vendorCount[$v] -lt $MaxPerVendor) { return $false }
    }
    return $true
  }

  function Emit-Progress([int]$Pct, [string]$Msg) {
    [pscustomobject]@{ Kind="Progress"; Pct=$Pct; Msg=$Msg }
  }
  function Emit-Result($Obj) {
    $Obj | Add-Member -NotePropertyName Kind -NotePropertyValue "Result" -Force
    $Obj
  }

  function Invoke-PoliteWebRequest([string]$Url) {
    if ($DelayMs -gt 0) { Start-Sleep -Milliseconds $DelayMs }
    $headers = @{
      "User-Agent"      = $UserAgent
      "Accept"          = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
      "Accept-Language" = "en-GB,en;q=0.9"
    }
    Invoke-WebRequest -Uri $Url -Headers $headers -TimeoutSec $TimeoutSec -ErrorAction Stop
  }

  function Invoke-PoliteJson([string]$Url) {
    $resp = Invoke-PoliteWebRequest $Url
    ($resp.Content | ConvertFrom-Json)
  }

  function Strip-Html([string]$html) {
    if ([string]::IsNullOrWhiteSpace($html)) { return "" }
    $plain = ($html -replace "<script[\s\S]*?</script>", " " -replace "<style[\s\S]*?</style>", " " -replace "<[^>]+>", " ")
    [regex]::Replace($plain, "\s+", " ").Trim()
  }

  function Parse-ABV([string]$Text) {
    $m = [regex]::Match($Text, "(\d{1,2}(?:\.\d)?)\s*%?\s*ABV", "IgnoreCase")
    if ($m.Success) { return [double]$m.Groups[1].Value }
    $m2 = [regex]::Match($Text, "(\d{1,2}(?:\.\d)?)\s*%", "IgnoreCase")
    if ($m2.Success) { return [double]$m2.Groups[1].Value }
    return $null
  }

  function Abv-Valid([Nullable[double]]$Abv) {
    if ($null -eq $Abv) { return $true }
    return ($Abv -gt 0 -and $Abv -le $MaxAbv)
  }

  function Abv-MatchesTipaQipa([Nullable[double]]$Abv, [string]$Name) {
    $n = $Name.ToLowerInvariant()
    if ($null -eq $Abv) {
      return ($StrongHints | Where-Object { $n.Contains($_) }) -ne $null
    }
    if ($n.Contains("quad") -or $n.Contains("qipa") -or $n.Contains("quipa")) { return ($Abv -ge ($QipaMinAbv - 0.5)) }
    if ($n.Contains("triple") -or $n.Contains("tipa")) { return ($Abv -ge ($TipaMinAbv - 0.5)) }
    return ($Abv -ge $TipaMinAbv)
  }

  function Matches-AnyKeyword([string]$blob, [string[]]$Keywords) {
    if ([string]::IsNullOrWhiteSpace($blob)) { return $false }
    $b = $blob.ToLowerInvariant()
    foreach ($k in $Keywords) {
      $kk = $k.ToLowerInvariant()
      if ([string]::IsNullOrWhiteSpace($kk)) { continue }
      if ($b.Contains($kk)) { return $true }
    }
    return $false
  }

  function Get-ShopifyProductsFromCollection([string]$BaseUrl, [string]$CollectionHandle, [string]$SourceName) {
    $all = New-Object System.Collections.Generic.List[object]
    $page = 1
    while ($true) {
      $u = "$BaseUrl/collections/$CollectionHandle/products.json?limit=250&page=$page"
      try { $json = Invoke-PoliteJson $u } catch { break }
      if ($null -eq $json -or $null -eq $json.products -or $json.products.Count -eq 0) { break }

      foreach ($prod in $json.products) {
        $title = [string]$prod.title
        $bodyPlain = Strip-Html ([string]$prod.body_html)
        $plain = ($title + " " + $bodyPlain).Trim()

        $abv = Parse-ABV $plain

        $inStock = $false
        if ($prod.variants) {
          foreach ($v in $prod.variants) {
            if ($v.available -eq $true) { $inStock = $true; break }
          }
        }

        $handle = [string]$prod.handle
        $url = if ($handle) { "$BaseUrl/products/$handle" } else { "" }

        $all.Add([pscustomobject]@{
          Source    = $SourceName
          Name      = $title
          ABV       = $abv
          InStock   = $inStock
          Url       = $url
          SeenAtUtc = (Get-Date).ToUniversalTime().ToString("s") + "Z"
          RawText   = $plain
        }) | Out-Null
      }

      $page++
      if ($page -gt 10) { break }
    }
    $all
  }

  function Get-ProductDetails([string]$Source, [string]$Url) {
    $resp = Invoke-PoliteWebRequest $Url
    $html = $resp.Content
    $plain = Strip-Html $html

    $title = $Url
    $mTitle = [regex]::Match($html, "<title>(.*?)</title>", "IgnoreCase")
    if ($mTitle.Success) { $title = ($mTitle.Groups[1].Value -replace "\s+", " ").Trim() }

    $abv = Parse-ABV $plain

    $inStock = $null
    if ([regex]::IsMatch($plain, "\bsold out\b", "IgnoreCase") -or [regex]::IsMatch($plain, "out of stock", "IgnoreCase")) {
      $inStock = $false
    } elseif ([regex]::IsMatch($plain, "add to cart", "IgnoreCase") -or [regex]::IsMatch($plain, "\bin stock\b", "IgnoreCase")) {
      $inStock = $true
    }

    [pscustomobject]@{
      Source    = $Source
      Name      = $title
      ABV       = $abv
      InStock   = $inStock
      Url       = $Url
      SeenAtUtc = (Get-Date).ToUniversalTime().ToString("s") + "Z"
      RawText   = $plain
    }
  }

  function Get-PollysCandidates {
    $pages = @(
      "https://pollys.co/shop/",
      "https://pollys.co/shop/page/2/",
      "https://pollys.co/shop/page/3/",
      "https://pollys.co/shop/page/4/",
      "https://pollys.co/shop/page/5/"
    )
    $set = New-Object System.Collections.Generic.HashSet[string]
    foreach ($p in $pages) {
      $resp = Invoke-PoliteWebRequest $p
      $html = $resp.Content
      $matches = [regex]::Matches($html, 'href="(https://pollys\.co/product/[^"]+)"', "IgnoreCase")
      foreach ($m in $matches) { [void]$set.Add($m.Groups[1].Value) }
    }
    $set
  }

  function Get-LeftFieldCandidates {
    $pages = @(
      "https://www.leftfieldbeer.co.uk/buy-beer",
      "https://www.leftfieldbeer.co.uk/buy-beer?page=2",
      "https://www.leftfieldbeer.co.uk/buy-beer?page=3",
      "https://www.leftfieldbeer.co.uk/buy-beer?page=4",
      "https://www.leftfieldbeer.co.uk/buy-beer?page=5"
    )
    $set = New-Object System.Collections.Generic.HashSet[string]
    foreach ($p in $pages) {
      $resp = Invoke-PoliteWebRequest $p
      $html = $resp.Content

      $matches = [regex]::Matches($html, 'href="(https://www\.leftfieldbeer\.co\.uk/product-page/[^"]+)"', "IgnoreCase")
      foreach ($m in $matches) { [void]$set.Add($m.Groups[1].Value) }

      $matches2 = [regex]::Matches($html, 'href="(/product-page/[^"]+)"', "IgnoreCase")
      foreach ($m in $matches2) { [void]$set.Add(("https://www.leftfieldbeer.co.uk" + $m.Groups[1].Value)) }
    }
    $set
  }

  function Get-BeerMerchantsCandidates {
    $pages = @(
      "https://www.beermerchants.com/browse/new-in",
      "https://www.beermerchants.com/browse/new-in?p=2",
      "https://www.beermerchants.com/browse/new-in?p=3",
      "https://www.beermerchants.com/browse/new-in?p=4",
      "https://www.beermerchants.com/browse/new-in?p=5"
    )
    $set = New-Object System.Collections.Generic.HashSet[string]
    foreach ($p in $pages) {
      $resp = Invoke-PoliteWebRequest $p
      $html = $resp.Content
      $matches = [regex]::Matches($html, 'href="(https://www\.beermerchants\.com/[^"#?]+)"', "IgnoreCase")
      foreach ($m in $matches) {
        $u = $m.Groups[1].Value
        if ($u -match "/browse/" -or $u -match "/customer/" -or $u -match "/home" -or $u -match "/cart" -or $u -match "/account") { continue }
        [void]$set.Add($u.TrimEnd("/"))
      }
    }
    $set
  }

  function Get-BeerRitzCandidates {
    $pages = @(
      "https://www.beerritz.co.uk/latest-arrivals",
      "https://www.beerritz.co.uk/latest-arrivals?page=2",
      "https://www.beerritz.co.uk/latest-arrivals?page=3",
      "https://www.beerritz.co.uk/latest-arrivals?page=4",
      "https://www.beerritz.co.uk/latest-arrivals?page=5"
    )
    $set = New-Object System.Collections.Generic.HashSet[string]
    foreach ($p in $pages) {
      $resp = Invoke-PoliteWebRequest $p
      $html = $resp.Content
      $matches = [regex]::Matches($html, 'href="(/[^"]+)"', "IgnoreCase")
      foreach ($m in $matches) {
        $href = $m.Groups[1].Value
        if ($href -match "^/latest-arrivals" -or $href -match "^/shop-by-brewery" -or $href -match "^/uk-beer$" -or $href -match "^/basket" -or $href -match "^/search") { continue }
        if ($href -notmatch "_\d" -and $href -notmatch "-\d+-") { continue }
        [void]$set.Add(("https://www.beerritz.co.uk" + $href).TrimEnd("/"))
      }
    }
    $set
  }

  Emit-Progress 1 "Starting..."

  $kwNorm = @()
  foreach ($k in $Keywords) { if ($k) { $kwNorm += $k.Trim() } }

  if ($SelectedSources.Verdant) {
    Emit-Progress 6 "Fetching Verdant (fast JSON)..."
    $prods = Get-ShopifyProductsFromCollection "https://verdantbrewing.co" "all-beers" "Verdant (FAST)"
    foreach ($p in $prods) {
      if (Vendor-LimitReached $p.Source) { break }
      if (-not (Matches-AnyKeyword $p.RawText $kwNorm)) { continue }
      if ($ExcludeOutOfStock -and $p.InStock -eq $false) { continue }
      if (-not (Abv-Valid $p.ABV)) { continue }
      if (-not (Abv-MatchesTipaQipa $p.ABV $p.Name)) { continue }
      $key = "$($p.Source)|" + (Normalize-Url $p.Url)
      if (-not $resultSeen.Add($key)) { continue }
      Emit-Result $p
      Inc-Vendor $p.Source
    }
  }

  if ($SelectedSources.Cloudwater) {
    Emit-Progress 12 "Fetching Cloudwater (fast JSON)..."
    $prods = Get-ShopifyProductsFromCollection "https://cloudwaterbrew.co" "cloudwater-cans" "Cloudwater (FAST)"
    foreach ($p in $prods) {
      if (Vendor-LimitReached $p.Source) { break }
      if (-not (Matches-AnyKeyword $p.RawText $kwNorm)) { continue }
      if ($ExcludeOutOfStock -and $p.InStock -eq $false) { continue }
      if (-not (Abv-Valid $p.ABV)) { continue }
      if (-not (Abv-MatchesTipaQipa $p.ABV $p.Name)) { continue }
      $key = "$($p.Source)|" + (Normalize-Url $p.Url)
      if (-not $resultSeen.Add($key)) { continue }
      Emit-Result $p
      Inc-Vendor $p.Source
    }
  }

  if ($SelectedSources.GhostWhale) {
    Emit-Progress 18 "Fetching Ghost Whale (fast JSON)..."
    $prods = Get-ShopifyProductsFromCollection "https://shop.ghostwhalelondon.com" "all-beers" "Ghost Whale (FAST)"
    foreach ($p in $prods) {
      if (Vendor-LimitReached $p.Source) { break }
      if (-not (Matches-AnyKeyword $p.RawText $kwNorm)) { continue }
      if ($ExcludeOutOfStock -and $p.InStock -eq $false) { continue }
      if (-not (Abv-Valid $p.ABV)) { continue }
      if (-not (Abv-MatchesTipaQipa $p.ABV $p.Name)) { continue }
      $key = "$($p.Source)|" + (Normalize-Url $p.Url)
      if (-not $resultSeen.Add($key)) { continue }
      Emit-Result $p
      Inc-Vendor $p.Source
    }
  }

  if ($SelectedSources.HopBurnsBlack) {
    Emit-Progress 24 "Fetching Hop Burns & Black (fast JSON)..."
    $prods = Get-ShopifyProductsFromCollection "https://www.hopburnsblack.co.uk" "beers" "Hop Burns & Black (FAST)"
    foreach ($p in $prods) {
      if (Vendor-LimitReached $p.Source) { break }
      if (-not (Matches-AnyKeyword $p.RawText $kwNorm)) { continue }
      if ($ExcludeOutOfStock -and $p.InStock -eq $false) { continue }
      if (-not (Abv-Valid $p.ABV)) { continue }
      if (-not (Abv-MatchesTipaQipa $p.ABV $p.Name)) { continue }
      $key = "$($p.Source)|" + (Normalize-Url $p.Url)
      if (-not $resultSeen.Add($key)) { continue }
      Emit-Result $p
      Inc-Vendor $p.Source
    }
  }

  if ($SelectedSources.RadBeer) {
    Emit-Progress 30 "Fetching RAD Beer (fast JSON)..."
    $prods = Get-ShopifyProductsFromCollection "https://radbeer.com" "beer" "RAD Beer (FAST)"
    foreach ($p in $prods) {
      if (Vendor-LimitReached $p.Source) { break }
      if (-not (Matches-AnyKeyword $p.RawText $kwNorm)) { continue }
      if ($ExcludeOutOfStock -and $p.InStock -eq $false) { continue }
      if (-not (Abv-Valid $p.ABV)) { continue }
      if (-not (Abv-MatchesTipaQipa $p.ABV $p.Name)) { continue }
      $key = "$($p.Source)|" + (Normalize-Url $p.Url)
      if (-not $resultSeen.Add($key)) { continue }
      Emit-Result $p
      Inc-Vendor $p.Source
    }
  }

  if ($SelectedSources.TremblingMadness) {
    Emit-Progress 36 "Fetching Trembling (fast JSON)..."
    $prods = Get-ShopifyProductsFromCollection "https://www.tremblingmadness.co.uk" "new-in" "Trembling Madness"
    foreach ($p in $prods) {
      if (Vendor-LimitReached $p.Source) { break }
      if (-not (Matches-AnyKeyword $p.RawText $kwNorm)) { continue }
      if ($ExcludeOutOfStock -and $p.InStock -eq $false) { continue }
      if (-not (Abv-Valid $p.ABV)) { continue }
      if (-not (Abv-MatchesTipaQipa $p.ABV $p.Name)) { continue }
      $key = "$($p.Source)|" + (Normalize-Url $p.Url)
      if (-not $resultSeen.Add($key)) { continue }
      Emit-Result $p
      Inc-Vendor $p.Source
    }
  }


  if ($SelectedSources.ShinyBrewery) {
    Emit-Progress 38 "Fetching Shiny Brewery (fast JSON)..."
    $prods = Get-ShopifyProductsFromCollection "https://shiny-brewing.myshopify.com" "all" "Shiny Brewery (FAST)"
    foreach ($p in $prods) {
      if (Vendor-LimitReached $p.Source) { break }
      if (-not (Matches-AnyKeyword $p.RawText $kwNorm)) { continue }
      if ($ExcludeOutOfStock -and $p.InStock -eq $false) { continue }
      if (-not (Abv-Valid $p.ABV)) { continue }
      if (-not (Abv-MatchesTipaQipa $p.ABV $p.Name)) { continue }
      $key = "$($p.Source)|" + (Normalize-Url $p.Url)
      if (-not $resultSeen.Add($key)) { continue }
      Emit-Result $p
      Inc-Vendor $p.Source
    }
  }

  if ($SelectedSources.BeerGuerrilla) {
    Emit-Progress 42 "Fetching Beer Guerrilla (fast JSON)..."
    $prods = Get-ShopifyProductsFromCollection "https://beerguerrilla.co.uk" "all" "Beer Guerrilla (FAST)"
    foreach ($p in $prods) {
      if (Vendor-LimitReached $p.Source) { break }
      if (-not (Matches-AnyKeyword $p.RawText $kwNorm)) { continue }
      if ($ExcludeOutOfStock -and $p.InStock -eq $false) { continue }
      if (-not (Abv-Valid $p.ABV)) { continue }
      if (-not (Abv-MatchesTipaQipa $p.ABV $p.Name)) { continue }
      $key = "$($p.Source)|" + (Normalize-Url $p.Url)
      if (-not $resultSeen.Add($key)) { continue }
      Emit-Result $p
      Inc-Vendor $p.Source
    }
  }

  if ($SelectedSources.Makemake) {
    Emit-Progress 44 "Fetching Makemake (fast JSON)..."
    $prods = Get-ShopifyProductsFromCollection "https://makemakebeer.myshopify.com" "available-beers" "Makemake (FAST)"
    foreach ($p in $prods) {
      if (Vendor-LimitReached $p.Source) { break }
      if (-not (Matches-AnyKeyword $p.RawText $kwNorm)) { continue }
      if ($ExcludeOutOfStock -and $p.InStock -eq $false) { continue }
      if (-not (Abv-Valid $p.ABV)) { continue }
      if (-not (Abv-MatchesTipaQipa $p.ABV $p.Name)) { continue }
      $key = "$($p.Source)|" + (Normalize-Url $p.Url)
      if (-not $resultSeen.Add($key)) { continue }
      Emit-Result $p
      Inc-Vendor $p.Source
    }
  }

  if ($SelectedSources.VaultCity) {
    Emit-Progress 44 "Fetching Vault City Brewing (fast JSON)..."
    $prods = Get-ShopifyProductsFromCollection "https://vaultcity.co.uk" "all-beers" "Vault City Brewing (FAST)"
    foreach ($p in $prods) {
      if (Vendor-LimitReached $p.Source) { break }
      if (-not (Matches-AnyKeyword $p.RawText $kwNorm)) { continue }
      if ($ExcludeOutOfStock -and $p.InStock -eq $false) { continue }
      if (-not (Abv-Valid $p.ABV)) { continue }
      if (-not (Abv-MatchesTipaQipa $p.ABV $p.Name)) { continue }
      $key = "$($p.Source)|" + (Normalize-Url $p.Url)
      if (-not $resultSeen.Add($key)) { continue }
      Emit-Result $p
      Inc-Vendor $p.Source
    }
  }


  $slowVendors = @()
  if ($SelectedSources.Pollys)        { $slowVendors += [pscustomobject]@{ Name="Polly's";         Links=(Get-PollysCandidates) } }
  if ($SelectedSources.LeftField)     { $slowVendors += [pscustomobject]@{ Name="Left Field Beer"; Links=(Get-LeftFieldCandidates) } }
  if ($SelectedSources.BeerMerchants) { $slowVendors += [pscustomobject]@{ Name="Beer Merchants";  Links=(Get-BeerMerchantsCandidates) } }
  if ($SelectedSources.BeerRitz)      { $slowVendors += [pscustomobject]@{ Name="Beer Ritz";       Links=(Get-BeerRitzCandidates) } }

  $slowLinks = @()
  foreach ($v in $slowVendors) {
    foreach ($l in ($v.Links | Sort-Object -Unique)) {
      $slowLinks += [pscustomobject]@{ Vendor=$v.Name; Url=$l }
    }
  }

  if ($slowLinks.Count -gt 0) {
    Emit-Progress 45 ("Crawling non-Shopify shops ({0} pages)..." -f $slowLinks.Count)
  }

  $i = 0
  foreach ($item in $slowLinks) {
    $i++
    $pct = 45 + [int](($i / [Math]::Max($slowLinks.Count,1)) * 54)
    $src = $item.Vendor
    $link = $item.Url

    if (All-VendorsSatisfied $SelectedSources) { break }
    if (Vendor-LimitReached $src) { continue }

    Emit-Progress $pct ("Checking: {0} ({1}/{2})" -f $src, $i, $slowLinks.Count)

    try {
      $p = Get-ProductDetails $src $link
      $blob = ($p.Name + " " + $p.RawText)

      if (-not (Matches-AnyKeyword $blob $kwNorm)) { continue }
      if ($ExcludeOutOfStock -and $p.InStock -eq $false) { continue }
      if (-not (Abv-Valid $p.ABV)) { continue }
      if (-not (Abv-MatchesTipaQipa $p.ABV $p.Name)) { continue }

      $key = "$src|" + (Normalize-Url $p.Url)
      if (-not $resultSeen.Add($key)) { continue }

      Emit-Result $p
      Inc-Vendor $src
    } catch {
      continue
    }
  }

  Emit-Progress 100 "Done."
}

# Timer polling: receive output chunks and update UI
$timer.add_Tick({
  if (-not $global:ScanJob) { return }

  $out = @()
  try { $out = @(Receive-Job -Job $global:ScanJob -Keep -ErrorAction SilentlyContinue) } catch { $out = @() }

  if ($out.Count -gt $global:SeenOutputCount) {
    $new = $out[$global:SeenOutputCount..($out.Count-1)]
    $global:SeenOutputCount = $out.Count

    foreach ($o in $new) {
      if ($o.Kind -eq "Progress") {
        $progress.Value = [Math]::Min([Math]::Max([int]$o.Pct,0),100)
        $status.Text = ("{0}% - {1}" -f [int]$o.Pct, [string]$o.Msg)
      }
      elseif ($o.Kind -eq "Result") {
        [void]$global:ResultBuffer.Add($o)
      }
    }
  }

  $state = (Get-Job -Id $global:ScanJob.Id -ErrorAction SilentlyContinue).State
  if ($state -in @("Completed","Failed","Stopped")) {
    $timer.Stop()
    $timerElapsed.Stop()

    if ($state -ne "Completed") {
      $status.Text = "Scan failed/stopped. State: $state"
    }

    # Populate grid (belt-and-braces de-dupe here as well)
    $table.Rows.Clear()
    $gridSeen = New-Object System.Collections.Generic.HashSet[string]

    foreach ($p in $global:ResultBuffer) {
      $k = (($p.Source + "|" + $p.Url).ToLowerInvariant())
      if (-not $gridSeen.Add($k)) { continue }

      $row = $table.NewRow()
      $row["Source"] = $p.Source
      $row["Name"] = $p.Name
      $row["ABV"] = if ($null -ne $p.ABV) { "{0:N1}%" -f $p.ABV } else { "" }
      $row["InStock"] = if ($p.InStock -eq $true) { "In stock" } elseif ($p.InStock -eq $false) { "Sold out" } else { "Unknown" }
      $row["Url"] = $p.Url
      $row["SeenAtUtc"] = $p.SeenAtUtc
      [void]$table.Rows.Add($row)
    }

    # Sort by newest first (ISO string sorts correctly)
    try { $table.DefaultView.Sort = "SeenAtUtc DESC" } catch {}

    $btnScan.Enabled = $true
    $btnStop.Enabled = $false
    $btnExport.Enabled = ($table.Rows.Count -gt 0)
    $btnOpen.Enabled = ($grid.SelectedRows.Count -gt 0)

    $progress.Value = 100
    # Final elapsed update (freeze)
    Update-ElapsedLabel
    $status.Text = "Complete. Matches: $($table.Rows.Count) - " + $lblElapsed.Text

    try { Remove-Job -Job $global:ScanJob -Force -ErrorAction SilentlyContinue } catch {}
    $global:ScanJob = $null
  }
})

$btnScan.Add_Click({
  if ($global:ScanJob) { return }

  $kw = @($txtKeywords.Text.Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ })
  if ($kw.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show(
      "Please provide at least one keyword.",
      "Hazy Hunter",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
    return
  }

  $sel = @{
    Verdant           = [bool]$chkVerdant.Checked
    Cloudwater        = [bool]$chkCloudwater.Checked
    Pollys            = [bool]$chkPollys.Checked
    LeftField         = [bool]$chkLeftField.Checked
    GhostWhale        = [bool]$chkGhostWhale.Checked
    HopBurnsBlack     = [bool]$chkHopBurnsBlack.Checked
    RadBeer           = [bool]$chkRadBeer.Checked
    TremblingMadness  = [bool]$chkTremblingMadness.Checked
    BeerMerchants     = [bool]$chkBeerMerchants.Checked
    BeerRitz          = [bool]$chkBeerRitz.Checked
    ShinyBrewery      = [bool]$chkShinyBrewery.Checked
    BeerGuerrilla     = [bool]$chkBeerGuerrilla.Checked
    Makemake          = [bool]$chkMakemake.Checked
    VaultCity         = [bool]$chkVaultCity.Checked
  }

  if (-not ($sel.Verdant -or $sel.Cloudwater -or $sel.Pollys -or $sel.LeftField -or $sel.GhostWhale -or $sel.HopBurnsBlack -or $sel.RadBeer -or $sel.TremblingMadness -or $sel.BeerMerchants -or $sel.BeerRitz -or $sel.ShinyBrewery -or $sel.BeerGuerrilla -or $sel.Makemake -or $sel.VaultCity)) {
    [System.Windows.Forms.MessageBox]::Show(
      "Select at least one source.",
      "Hazy Hunter",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
    return
  }

  $excludeOos = [bool]$chkExcludeOOS.Checked
  $limitPerVendor = [bool]$chkLimitPerVendor.Checked
  $maxPerVendor = [int]$numLimit.Value
  $minAbv = [double]$numMinAbv.Value

  # reset UI + state
  $table.Rows.Clear()
  $global:ResultBuffer = New-Object System.Collections.Generic.List[object]
  $global:SeenOutputCount = 0
  $progress.Value = 0
  $status.Text = "Starting scan..."
  $btnScan.Enabled = $false
  $btnStop.Enabled = $true
  $btnExport.Enabled = $false
  $btnOpen.Enabled = $false

  # elapsed
  $global:ScanStartUtc = (Get-Date).ToUniversalTime()
  Update-ElapsedLabel
  $timerElapsed.Start()

  $global:ScanJob = Start-Job -ScriptBlock $jobScript -ArgumentList $sel, $kw, $excludeOos, $limitPerVendor, $maxPerVendor, $minAbv
  $timer.Start()
})

$form.Add_FormClosing({
  if ($global:ScanJob) {
    try { Stop-Job -Job $global:ScanJob -Force -ErrorAction SilentlyContinue } catch {}
    try { Remove-Job -Job $global:ScanJob -Force -ErrorAction SilentlyContinue } catch {}
    $global:ScanJob = $null
  }
})

$form.Controls.AddRange(@(
  $lbl, $lblElapsed, $grp, $lblKw, $txtKeywords, $grpOptions,
  $btnScan, $btnStop, $btnExport, $btnOpen, $progress, $status, $grid
))

[void]$form.ShowDialog()
