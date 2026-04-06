# ============================================================
# EXCHANGE 2019 ON-PREM -- SEARCH & DELETE TOOL v6
# ============================================================
# Description : GUI tool to search emails across all Exchange
#               mailboxes by Subject, Sender, Recipient and
#               Date Range -- then Soft Delete (recoverable)
#               or Hard Delete (permanent purge) matching
#               emails from every affected mailbox.
#               Supports Unicode subjects (Sinhala, Arabic etc).
# Author      : Nuwan Gamage
# Version     : 6.0
# Tested On   : Exchange Server 2019 Build 15.2.1748.10
# Requirements: Exchange Management Shell -- run as Administrator
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

$global:MailboxStats   = @()
$global:SearchResults  = @()
$global:DiscoveryMbx   = "Discovery Search Mailbox"

function Write-Log {
    param([string]$msg, [string]$colour = "LightGreen")
    $time = Get-Date -Format "HH:mm:ss"
    $txtLog.SelectionStart  = $txtLog.TextLength
    $txtLog.SelectionLength = 0
    $txtLog.SelectionColor  = [System.Drawing.Color]::$colour
    $txtLog.AppendText("[$time] $msg`r`n")
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-SizeMB {
    param($stats)
    if (-not $stats) { return "N/A" }
    try { return $stats.TotalItemSize.Value.ToMB().ToString("N1") + " MB" } catch {}
    try {
        $str = "$($stats.TotalItemSize)"
        if ($str -match "\(([0-9,]+)\s*bytes\)") {
            $bytes = [long]($Matches[1] -replace ",","")
            return ([math]::Round($bytes/1MB,1)).ToString("N1") + " MB"
        }
        return $str
    } catch { return "N/A" }
}

function Init-Exchange {
    try {
        $reg = Get-PSSnapin -Registered -ErrorAction SilentlyContinue |
               Where-Object { $_.Name -like "*Exchange*" }
        if (-not $reg) {
            Write-Log "ERROR: Exchange snap-in not found. Run from Exchange Management Shell." "Red"
            return
        }
        if (-not (Get-PSSnapin | Where-Object { $_.Name -like "*Exchange*" })) {
            Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn -ErrorAction Stop
        }
        $ver = (Get-ExchangeServer | Select-Object -First 1).AdminDisplayVersion
        Write-Log "Connected -- Exchange $ver" "LightGreen"
        $btnSearch.Enabled   = $true
        $btnExport.Enabled   = $true
        $lblStatus.Text      = "Status: Connected -- Exchange $ver"
        $lblStatus.ForeColor = [System.Drawing.Color]::LightGreen
        Check-Roles
    } catch {
        Write-Log "ERROR loading Exchange: $_" "Red"
    }
}

function Check-Roles {
    $user = $env:USERNAME
    Write-Log "Checking roles for: $user" "White"
    $allOk = $true
    foreach ($role in @("Mailbox Search","Mailbox Import Export")) {
        $a = Get-ManagementRoleAssignment -Role $role -ErrorAction SilentlyContinue |
             Where-Object { $_.RoleAssigneeName -like "*$user*" }
        if ($a) { Write-Log "  [OK] $role" "LightGreen" }
        else    { Write-Log "  [MISSING] $role -- click Fix Roles" "Yellow"; $allOk = $false }
    }
    $cmd = Get-Command Search-Mailbox -ErrorAction SilentlyContinue
    if (-not $cmd) {
        Write-Log "  [MISSING] Search-Mailbox cmdlet" "Red"; $allOk = $false
    } elseif ($cmd.Parameters.ContainsKey("DeleteContent")) {
        Write-Log "  [OK] Search-Mailbox with -DeleteContent ($($cmd.Parameters.Count) params)" "LightGreen"
    } else {
        Write-Log "  [WARNING] Search-Mailbox loaded WITHOUT -DeleteContent" "Yellow"
        Write-Log "  --> LOG OFF from Windows, log back on, reopen EMS, reconnect." "Orange"
        $allOk = $false
    }
    if ($allOk) {
        Write-Log "All checks passed -- ready to delete." "LightGreen"
        $btnEstimate.Enabled = $true
        $btnSoftDel.Enabled  = $true
        $btnHardDel.Enabled  = $true
    } else {
        Write-Log "NOT READY -- fix issues above." "Red"
    }
}

function Fix-Roles {
    $user = $env:USERNAME
    Write-Log "Assigning roles to: $user" "Orange"
    foreach ($role in @("Mailbox Search","Mailbox Import Export")) {
        try {
            $e = Get-ManagementRoleAssignment -Role $role -ErrorAction SilentlyContinue |
                 Where-Object { $_.RoleAssigneeName -like "*$user*" }
            if ($e) { Write-Log "  Already assigned: $role" "Gray" }
            else    { New-ManagementRoleAssignment -Role $role -User $user -ErrorAction Stop | Out-Null; Write-Log "  Assigned: $role" "LightGreen" }
        } catch { Write-Log "  ERROR $role : $_" "Red" }
    }
    try {
        $m = Get-RoleGroupMember "Discovery Management" -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -like "*$user*" }
        if ($m) { Write-Log "  Already in: Discovery Management" "Gray" }
        else    { Add-RoleGroupMember "Discovery Management" -Member $user -ErrorAction Stop; Write-Log "  Added to: Discovery Management" "LightGreen" }
    } catch { Write-Log "  ERROR Discovery Management: $_" "Red" }
    Write-Log "" "White"
    Write-Log "*** LOG OFF from Windows on this server." "Orange"
    Write-Log "*** Log back on, open EMS, run script, click Connect." "Orange"
}

# Build query -- returns a ScriptBlock so Unicode is preserved natively
function Build-Query {
    $parts = @()
    $subj = $txtSubject.Text.Trim()
    $from = $txtSender.Text.Trim()
    $to   = $txtRecipient.Text.Trim()
    if ($subj) {
        $clean = ($subj -replace '"','') -replace "'",""
        $parts += "Subject:`"$clean`""
    }
    if ($from) { $parts += "From:`"$from`"" }
    if ($to)   { $parts += "To:`"$to`"" }
    $d1 = $dtFrom.Value.ToString("MM/dd/yyyy")
    $d2 = $dtTo.Value.ToString("MM/dd/yyyy")
    $parts += "Received:$d1..$d2"
    return ($parts -join " AND ")
}

function Resolve-Mailbox {
    param([string]$addr)
    $mb = Get-Mailbox -Identity $addr -ErrorAction SilentlyContinue
    if ($mb) { return $mb }
    $mb = Get-Mailbox -Filter "EmailAddresses -like '*$addr*'" -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($mb) { return $mb }
    $local = ($addr -split "@")[0]
    $mb = Get-Mailbox -ANR $local -ErrorAction SilentlyContinue | Select-Object -First 1
    return $mb
}

function Start-Search {
    $grid.DataSource = $null
    $global:MailboxStats  = @()
    $global:SearchResults = @()
    $lstAffected.Items.Clear()
    $lblSummary.Text = "Mailboxes: 0  |  Messages: 0"
    Write-Log "=== SEARCH STARTED ===" "Cyan"

    if (-not $txtSubject.Text.Trim() -and -not $txtSender.Text.Trim() -and -not $txtRecipient.Text.Trim()) {
        Write-Log "ERROR: Enter at least Subject, Sender, or Recipient." "Red"; return
    }

    $allResults = [System.Collections.Generic.List[object]]::new()
    try { $servers = Get-TransportService -ErrorAction Stop }
    catch { Write-Log "Cannot get Transport Services: $_" "Red"; return }

    foreach ($srv in $servers) {
        Write-Log "Scanning: $($srv.Name)" "Gray"
        try {
            $logs = Get-MessageTrackingLog -Server $srv.Name -Start $dtFrom.Value -End $dtTo.Value `
                    -ResultSize Unlimited -WarningAction SilentlyContinue -ErrorAction Stop
            foreach ($l in $logs) { $allResults.Add($l) }
            Write-Log "  $($logs.Count) entries from $($srv.Name)" "Gray"
        } catch { Write-Log "  WARNING $($srv.Name): $_" "Yellow" }
    }
    Write-Log "Total tracking entries: $($allResults.Count)" "White"

    $s_subj = $txtSubject.Text.Trim()
    $s_from = $txtSender.Text.Trim()
    $s_to   = $txtRecipient.Text.Trim()

    $filtered = $allResults | Where-Object {
        ($s_subj -eq "" -or $_.MessageSubject -like "*$s_subj*") -and
        ($s_from -eq "" -or $_.Sender         -like "*$s_from*") -and
        ($s_to   -eq "" -or ($_.Recipients -join ",") -like "*$s_to*")
    }

    if (-not $filtered) { Write-Log "No matching messages found." "Yellow"; return }
    Write-Log "Filtered results: $($filtered.Count)" "LightGreen"
    $global:SearchResults = $filtered

    $addrSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($r in $filtered) {
        if ($r.Sender -and $r.Sender -ne "<>") { [void]$addrSet.Add($r.Sender) }
        foreach ($rec in $r.Recipients) { if ($rec -and $rec -ne "<>") { [void]$addrSet.Add($rec) } }
    }
    Write-Log "Unique addresses: $($addrSet.Count) -- resolving..." "White"

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($addr in ($addrSet | Sort-Object)) {
        $mb    = Resolve-Mailbox $addr
        $stats = $null
        if ($mb) { try { $stats = Get-MailboxStatistics -Identity $mb.Alias -ErrorAction SilentlyContinue } catch {} }
        $identity = if ($mb) { $mb.PrimarySmtpAddress.ToString() } else { $addr }
        $rows.Add([PSCustomObject]@{
            SearchIdentity = $identity
            RawAddress     = $addr
            DisplayName    = if ($mb)    { $mb.DisplayName }   else { "(not in Exchange)" }
            Database       = if ($mb)    { "$($mb.Database)" } else { "N/A" }
            Items          = if ($stats) { $stats.ItemCount }  else { "N/A" }
            SizeMB         = Get-SizeMB $stats
            Status         = if ($mb)    { "Verified" }        else { "Not in GAL" }
        })
        if ($mb) { Write-Log "  [FOUND] $addr -> $($mb.DisplayName)" "LightGreen" }
        else     { Write-Log "  [NOT IN GAL] $addr" "Yellow" }
    }

    $global:MailboxStats = $rows
    $bl = New-Object System.ComponentModel.BindingList[PSCustomObject]
    foreach ($r in $rows) { $bl.Add($r) }
    $grid.DataSource = $bl
    $grid.AutoSizeColumnsMode = "AllCells"
    $lstAffected.Items.Clear()
    foreach ($r in $rows) { [void]$lstAffected.Items.Add($r.SearchIdentity) }
    $verified = ($rows | Where-Object { $_.Status -eq "Verified" }).Count
    $lblSummary.Text = "Mailboxes: $($rows.Count)  |  In GAL: $verified  |  Tracking hits: $($filtered.Count)"
    $btnEstimate.Enabled = $true; $btnSoftDel.Enabled = $true
    $btnHardDel.Enabled  = $true; $btnExport.Enabled  = $true
    Write-Log "=== SEARCH COMPLETE -- $($rows.Count) addresses ($verified in GAL) ===" "Cyan"
}

# -------------------------------------------------------
# Core delete -- passes query as a VARIABLE not a string
# This preserves Unicode (Sinhala) characters correctly
# -------------------------------------------------------
function Invoke-Delete {
    param(
        [string]$identity,
        [string]$query,
        [bool]$hardDelete,
        [string]$disc,
        [string]$folder
    )

    if ($hardDelete) {
        # Hard delete -- $query is a real PS variable, Unicode intact
        $result = Search-Mailbox `
            -Identity      $identity `
            -SearchQuery   $query `
            -DeleteContent `
            -Force         `
            -LogLevel      Full `
            -ErrorAction   Stop
    } else {
        # Soft delete -- copy to discovery then delete
        $result = Search-Mailbox `
            -Identity      $identity `
            -SearchQuery   $query `
            -TargetMailbox $disc `
            -TargetFolder  $folder `
            -DeleteContent `
            -Force         `
            -LogLevel      Full `
            -ErrorAction   Stop
    }
    return $result
}

function Run-Delete {
    param($row, [string]$query, [bool]$hardDelete)
    $tag    = if ($hardDelete) { "PURGED" } else { "SOFT-DEL" }
    $disc   = $global:DiscoveryMbx
    $folder = "SoftDel_$(Get-Date -Format 'yyyyMMdd_HHmmss')"

    foreach ($id in (@($row.SearchIdentity, $row.RawAddress) | Select-Object -Unique)) {
        try {
            $r = Invoke-Delete -identity $id -query $query -hardDelete $hardDelete -disc $disc -folder $folder
            $cnt = if ($r.ResultItemsCount) { $r.ResultItemsCount } else { "0" }
            Write-Log "  [$tag] $id : $cnt item(s)" "LightGreen"
            return $true
        } catch {
            Write-Log "  WARN $id : $_" "Yellow"
        }
    }
    Write-Log "  FAILED: $($row.RawAddress)" "Red"
    return $false
}

function Start-Estimate {
    if (-not $global:MailboxStats -or $global:MailboxStats.Count -eq 0) { Write-Log "Run Search first." "Yellow"; return }
    $query = Build-Query
    Write-Log "=== ESTIMATE (nothing deleted) ===" "Cyan"
    Write-Log "Query: $query" "White"
    $total = 0
    foreach ($row in $global:MailboxStats) {
        foreach ($id in (@($row.SearchIdentity, $row.RawAddress) | Select-Object -Unique)) {
            try {
                $r = Search-Mailbox -Identity $id -SearchQuery $query -EstimateResultOnly -ErrorAction Stop
                $cnt = if ($r.ResultItemsCount) { $r.ResultItemsCount } else { 0 }
                Write-Log "  $id : $cnt item(s)  $($r.ResultItemsSize)" "LightGreen"
                $total += [int]$cnt
                break
            } catch { Write-Log "  SKIP $id : $_" "Yellow" }
        }
    }
    Write-Log "=== ESTIMATE COMPLETE -- Total: $total item(s) ===" "Cyan"
}

function Start-SoftDelete {
    if (-not $global:MailboxStats -or $global:MailboxStats.Count -eq 0) { Write-Log "Run Search first." "Yellow"; return }
    $global:DiscoveryMbx = $txtDiscovery.Text.Trim()
    $query = Build-Query
    Write-Log "Query will be: $query" "White"
    $msg = "SOFT DELETE`n`nEmails move to Recoverable Items.`nUsers CAN restore via Outlook.`n`nQuery     : $query`nMailboxes : $($global:MailboxStats.Count)`nDiscovery : $($global:DiscoveryMbx)`n`nProceed?"
    $ans = [System.Windows.Forms.MessageBox]::Show($msg,"Confirm Soft Delete",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Warning)
    if ($ans -ne [System.Windows.Forms.DialogResult]::Yes) { Write-Log "Cancelled." "Yellow"; return }
    Write-Log "=== SOFT DELETE STARTED ===" "Orange"
    Write-Log "Query: $query" "White"
    $progressBar.Maximum = $global:MailboxStats.Count; $progressBar.Value = 0
    $ok = 0; $fail = 0
    foreach ($row in $global:MailboxStats) {
        Write-Log "  Processing: $($row.SearchIdentity)" "Gray"
        if (Run-Delete -row $row -query $query -hardDelete $false) { $ok++ } else { $fail++ }
        $progressBar.Value++
        [System.Windows.Forms.Application]::DoEvents()
    }
    $progressBar.Value = 0
    Write-Log "=== SOFT DELETE DONE -- Success: $ok  Errors: $fail ===" "LightGreen"
    [System.Windows.Forms.MessageBox]::Show("Soft delete complete.`nSuccess: $ok`nErrors: $fail","Done",
        [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

function Start-HardDelete {
    if (-not $global:MailboxStats -or $global:MailboxStats.Count -eq 0) { Write-Log "Run Search first." "Yellow"; return }
    $query = Build-Query
    $ans1 = [System.Windows.Forms.MessageBox]::Show(
        "*** HARD DELETE -- PERMANENT PURGE ***`n`nEmails will be PERMANENTLY deleted.`nCANNOT be undone.`n`nQuery     : $query`nMailboxes : $($global:MailboxStats.Count)`n`nAre you sure?",
        "WARNING: Hard Delete",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Stop)
    if ($ans1 -ne [System.Windows.Forms.DialogResult]::Yes) { Write-Log "Cancelled." "Yellow"; return }
    $ans2 = [System.Windows.Forms.MessageBox]::Show(
        "SECOND CONFIRMATION`n`nPermanently purge from $($global:MailboxStats.Count) mailboxes?",
        "Second Confirmation",[System.Windows.Forms.MessageBoxButtons]::OKCancel,[System.Windows.Forms.MessageBoxIcon]::Stop)
    if ($ans2 -ne [System.Windows.Forms.DialogResult]::OK) { Write-Log "Cancelled." "Yellow"; return }
    $typed = [Microsoft.VisualBasic.Interaction]::InputBox("Type  HARDDELETE  (all caps) to confirm:","Final Safety Check","")
    if ($typed -cne "HARDDELETE") { Write-Log "Cancelled -- word did not match." "Yellow"; return }
    Write-Log "=== HARD DELETE STARTED ===" "Red"
    Write-Log "Query: $query" "White"
    $progressBar.Maximum = $global:MailboxStats.Count; $progressBar.Value = 0
    $ok = 0; $fail = 0
    foreach ($row in $global:MailboxStats) {
        Write-Log "  Purging: $($row.SearchIdentity)" "Orange"
        if (Run-Delete -row $row -query $query -hardDelete $true) { $ok++ } else { $fail++ }
        $progressBar.Value++
        [System.Windows.Forms.Application]::DoEvents()
    }
    $progressBar.Value = 0
    Write-Log "=== HARD DELETE DONE -- Success: $ok  Errors: $fail ===" "Red"
    [System.Windows.Forms.MessageBox]::Show("Hard delete complete.`nSuccess: $ok`nErrors: $fail","Done",
        [System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
}

function Export-Report {
    if (-not $global:MailboxStats -or $global:MailboxStats.Count -eq 0) { Write-Log "Nothing to export." "Yellow"; return }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "Text File (*.txt)|*.txt|CSV File (*.csv)|*.csv|All Files (*.*)|*.*"
    $dlg.FileName = "Exchange_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $path = $dlg.FileName
    if ($path -like "*.csv") {
        $global:MailboxStats | Export-Csv -Path $path -NoTypeInformation -Encoding UTF8
    } else {
        $sep = "=" * 70
        $lines = @("Exchange Search and Delete Report","Generated : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",$sep,
            "Subject   : $($txtSubject.Text)","Sender    : $($txtSender.Text)",
            "Recipient : $($txtRecipient.Text)","From Date : $($dtFrom.Value.ToString('yyyy-MM-dd'))",
            "To Date   : $($dtTo.Value.ToString('yyyy-MM-dd'))",$sep,
            "AFFECTED MAILBOXES ($($global:MailboxStats.Count))","")
        foreach ($r in $global:MailboxStats) {
            $lines += "$($r.SearchIdentity.PadRight(52)) | $($r.DisplayName.PadRight(30)) | $($r.Status)"
        }
        $lines | Out-File -FilePath $path -Encoding UTF8
    }
    Write-Log "Report saved: $path" "LightGreen"
}

function Clear-All {
    $txtSubject.Text = ""; $txtSender.Text = ""; $txtRecipient.Text = ""
    $dtFrom.Value = (Get-Date).AddDays(-7); $dtTo.Value = Get-Date
    $grid.DataSource = $null; $lstAffected.Items.Clear()
    $lblSummary.Text = "Mailboxes: 0  |  Messages: 0"
    $txtLog.Clear(); $global:MailboxStats = @(); $global:SearchResults = @()
    Write-Log "Cleared." "Gray"
}

function Copy-Selected {
    if ($lstAffected.SelectedItem) {
        [System.Windows.Forms.Clipboard]::SetText($lstAffected.SelectedItem.ToString())
        Write-Log "Copied: $($lstAffected.SelectedItem)" "Gray"
    }
}

# ============================================================
# GUI
# ============================================================

$form               = New-Object System.Windows.Forms.Form
$form.Text          = "Exchange 2019 -- Search and Delete Tool v6"
$form.Size          = New-Object System.Drawing.Size(1400,860)
$form.MinimumSize   = New-Object System.Drawing.Size(1200,720)
$form.StartPosition = "CenterScreen"
$form.BackColor     = [System.Drawing.Color]::FromArgb(30,30,35)
$form.ForeColor     = [System.Drawing.Color]::WhiteSmoke
$form.Font          = New-Object System.Drawing.Font("Segoe UI",9)

$statusStrip           = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = [System.Drawing.Color]::FromArgb(20,20,25)
$lblStatus             = New-Object System.Windows.Forms.ToolStripStatusLabel
$lblStatus.Text        = "Status: Not connected  --  Click Connect"
$lblStatus.ForeColor   = [System.Drawing.Color]::Tomato
$statusStrip.Items.Add($lblStatus) | Out-Null
$form.Controls.Add($statusStrip)

$pnlTop           = New-Object System.Windows.Forms.GroupBox
$pnlTop.Text      = "Search Filters"
$pnlTop.Location  = New-Object System.Drawing.Point(8,8)
$pnlTop.Size      = New-Object System.Drawing.Size(1360,140)
$pnlTop.BackColor = [System.Drawing.Color]::FromArgb(40,40,50)
$pnlTop.ForeColor = [System.Drawing.Color]::WhiteSmoke
$form.Controls.Add($pnlTop)

function MkLabel($text,$x,$y) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text=$text; $l.Location=New-Object System.Drawing.Point($x,$y)
    $l.AutoSize=$true; $l.ForeColor=[System.Drawing.Color]::WhiteSmoke
    $pnlTop.Controls.Add($l)
}
function MkTextBox($x,$y,$w) {
    $t = New-Object System.Windows.Forms.TextBox
    $t.Location=New-Object System.Drawing.Point($x,$y); $t.Size=New-Object System.Drawing.Size($w,22)
    $pnlTop.Controls.Add($t); return $t
}
function MkDatePicker($x,$y,$w,$val) {
    $d = New-Object System.Windows.Forms.DateTimePicker
    $d.Location=New-Object System.Drawing.Point($x,$y); $d.Width=$w; $d.Value=$val
    $pnlTop.Controls.Add($d); return $d
}

MkLabel "Subject:"    10 24; $txtSubject   = MkTextBox  72 20 230
MkLabel "Sender:"    317 24; $txtSender    = MkTextBox 372 20 230
MkLabel "Recipient:" 617 24; $txtRecipient = MkTextBox 680 20 250
MkLabel "From Date:"  10 60; $dtFrom = MkDatePicker  82 57 165 ((Get-Date).AddDays(-7))
MkLabel "To Date:"   262 60; $dtTo   = MkDatePicker 320 57 165 (Get-Date)
MkLabel "Discovery Mailbox:" 502 60; $txtDiscovery = MkTextBox 640 57 280
$txtDiscovery.Text = $global:DiscoveryMbx

function MkBtn($text,$x,$w,$bg,$fg) {
    $b = New-Object System.Windows.Forms.Button
    $b.Text=$text; $b.Location=New-Object System.Drawing.Point($x,100); $b.Size=New-Object System.Drawing.Size($w,28)
    $b.BackColor=[System.Drawing.Color]::$bg; $b.ForeColor=[System.Drawing.Color]::$fg
    $b.FlatStyle="Flat"; $b.FlatAppearance.BorderSize=0
    $b.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
    $pnlTop.Controls.Add($b); return $b
}

$btnConnect  = MkBtn "1. Connect"       10 105 "SteelBlue"     "White"
$btnFixRoles = MkBtn "Fix Roles"       123 100 "DarkGoldenrod"  "White"
$btnChkRoles = MkBtn "Check Roles"     231 100 "DimGray"       "White"
$btnSearch   = MkBtn "2. Search"       339  95 "DarkSlateGray"  "White"
$btnEstimate = MkBtn "Estimate"        442  90 "DimGray"       "White"
$btnSoftDel  = MkBtn "3a. Soft Delete" 540 125 "DarkOrange"    "White"
$btnHardDel  = MkBtn "3b. Hard Delete" 673 125 "Firebrick"     "White"
$btnExport   = MkBtn "Export"          806  80 "DarkGreen"     "White"
$btnClear    = MkBtn "Clear"           894  75 "SlateGray"     "White"

foreach ($b in @($btnSearch,$btnEstimate,$btnSoftDel,$btnHardDel,$btnExport)) { $b.Enabled = $false }

$btnConnect.Add_Click({  $global:DiscoveryMbx=$txtDiscovery.Text.Trim(); Init-Exchange })
$btnFixRoles.Add_Click({ Fix-Roles })
$btnChkRoles.Add_Click({ Check-Roles })
$btnSearch.Add_Click({   Start-Search })
$btnEstimate.Add_Click({ Start-Estimate })
$btnSoftDel.Add_Click({  Start-SoftDelete })
$btnHardDel.Add_Click({  Start-HardDelete })
$btnExport.Add_Click({   Export-Report })
$btnClear.Add_Click({    Clear-All })

$tip = New-Object System.Windows.Forms.ToolTip
$tip.SetToolTip($btnSoftDel,  "Moves emails to Recoverable Items. Users can restore via Outlook.")
$tip.SetToolTip($btnHardDel,  "PERMANENTLY purges emails. NOT recoverable. Triple confirmation required.")
$tip.SetToolTip($btnEstimate, "Safe dry run -- counts items only, nothing deleted.")
$tip.SetToolTip($txtDiscovery,"Target mailbox for backup copies during Soft Delete.")

$split = New-Object System.Windows.Forms.SplitContainer
$split.Location=New-Object System.Drawing.Point(8,156); $split.Size=New-Object System.Drawing.Size(1360,615)
$split.SplitterDistance=940; $split.BackColor=[System.Drawing.Color]::FromArgb(30,30,35)
$form.Controls.Add($split)

$pnlGrid = New-Object System.Windows.Forms.GroupBox
$pnlGrid.Text="Affected Mailboxes"; $pnlGrid.Dock="Fill"
$pnlGrid.BackColor=[System.Drawing.Color]::FromArgb(40,40,50); $pnlGrid.ForeColor=[System.Drawing.Color]::WhiteSmoke
$split.Panel1.Controls.Add($pnlGrid)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock="Fill"; $grid.ReadOnly=$true; $grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false
$grid.AutoSizeColumnsMode="AllCells"; $grid.SelectionMode="FullRowSelect"
$grid.BackgroundColor=[System.Drawing.Color]::FromArgb(30,30,40)
$grid.DefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(30,30,40)
$grid.DefaultCellStyle.ForeColor=[System.Drawing.Color]::WhiteSmoke
$grid.AlternatingRowsDefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(42,42,55)
$grid.ColumnHeadersDefaultCellStyle.BackColor=[System.Drawing.Color]::FromArgb(20,90,160)
$grid.ColumnHeadersDefaultCellStyle.ForeColor=[System.Drawing.Color]::White
$grid.ColumnHeadersDefaultCellStyle.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
$grid.EnableHeadersVisualStyles=$false; $grid.GridColor=[System.Drawing.Color]::FromArgb(55,55,75)
$pnlGrid.Controls.Add($grid)

$lblSummary=New-Object System.Windows.Forms.Label; $lblSummary.Text="Mailboxes: 0  |  Messages: 0"
$lblSummary.Dock="Bottom"; $lblSummary.Height=22
$lblSummary.Font=New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
$lblSummary.ForeColor=[System.Drawing.Color]::LightSkyBlue; $lblSummary.BackColor=[System.Drawing.Color]::FromArgb(20,20,30)
$pnlGrid.Controls.Add($lblSummary)

$progressBar=New-Object System.Windows.Forms.ProgressBar; $progressBar.Dock="Bottom"
$progressBar.Height=12; $progressBar.Style="Continuous"; $pnlGrid.Controls.Add($progressBar)

$rightSplit=New-Object System.Windows.Forms.SplitContainer; $rightSplit.Dock="Fill"
$rightSplit.Orientation="Horizontal"; $rightSplit.SplitterDistance=290
$rightSplit.BackColor=[System.Drawing.Color]::FromArgb(30,30,35); $split.Panel2.Controls.Add($rightSplit)

$pnlList=New-Object System.Windows.Forms.GroupBox; $pnlList.Text="Affected Addresses (double-click to copy)"
$pnlList.Dock="Fill"; $pnlList.BackColor=[System.Drawing.Color]::FromArgb(40,40,50)
$pnlList.ForeColor=[System.Drawing.Color]::WhiteSmoke; $rightSplit.Panel1.Controls.Add($pnlList)

$lstAffected=New-Object System.Windows.Forms.ListBox; $lstAffected.Dock="Fill"
$lstAffected.Font=New-Object System.Drawing.Font("Consolas",8.5)
$lstAffected.BackColor=[System.Drawing.Color]::FromArgb(22,22,32)
$lstAffected.ForeColor=[System.Drawing.Color]::LightGreen
$lstAffected.Add_DoubleClick({ Copy-Selected }); $pnlList.Controls.Add($lstAffected)

$btnCopyAddr=New-Object System.Windows.Forms.Button; $btnCopyAddr.Text="Copy Selected"
$btnCopyAddr.Dock="Bottom"; $btnCopyAddr.Height=24
$btnCopyAddr.BackColor=[System.Drawing.Color]::FromArgb(50,50,70); $btnCopyAddr.ForeColor=[System.Drawing.Color]::White
$btnCopyAddr.FlatStyle="Flat"; $btnCopyAddr.Add_Click({ Copy-Selected }); $pnlList.Controls.Add($btnCopyAddr)

$pnlLog=New-Object System.Windows.Forms.GroupBox; $pnlLog.Text="Activity Log"
$pnlLog.Dock="Fill"; $pnlLog.BackColor=[System.Drawing.Color]::FromArgb(18,18,22)
$pnlLog.ForeColor=[System.Drawing.Color]::WhiteSmoke; $rightSplit.Panel2.Controls.Add($pnlLog)

$txtLog=New-Object System.Windows.Forms.RichTextBox; $txtLog.Dock="Fill"; $txtLog.ReadOnly=$true
$txtLog.Font=New-Object System.Drawing.Font("Consolas",8.5)
$txtLog.BackColor=[System.Drawing.Color]::FromArgb(10,10,16); $txtLog.ForeColor=[System.Drawing.Color]::LightGreen
$pnlLog.Controls.Add($txtLog)

$btnClearLog=New-Object System.Windows.Forms.Button; $btnClearLog.Text="Clear Log"
$btnClearLog.Dock="Bottom"; $btnClearLog.Height=22
$btnClearLog.BackColor=[System.Drawing.Color]::FromArgb(50,50,70); $btnClearLog.ForeColor=[System.Drawing.Color]::White
$btnClearLog.FlatStyle="Flat"; $btnClearLog.Add_Click({ $txtLog.Clear() }); $pnlLog.Controls.Add($btnClearLog)

Write-Log "Exchange 2019 Search and Delete Tool v6 ready." "Cyan"
Write-Log "v6 fix: query passed as variable -- Unicode subjects (Sinhala etc) work correctly." "LightGreen"
Write-Log "Step 1: Connect  |  Step 2: Search  |  Step 3: Estimate  |  Step 4: Soft/Hard Delete" "White"

$form.ShowDialog() | Out-Null
