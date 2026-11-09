<#
.SYNOPSIS
    Posts time entries from the "Daily Work Activity Summary" email into ConnectWise Manage.

.DESCRIPTION
    1. Reads Reports\CWTime-yyyy-MM-dd.csv (saved to OneDrive by the Cowork daily report) for the
       report date (default: yesterday). If that file is missing, falls back to the email in classic Outlook.
    2. Parses the "CONNECTWISE TIME ENTRIES" table.
    3. Maps each row:  "#1234567" in the Ticket/Meeting column  -> that service (or project) ticket
                       anything else                             -> a ticket (TicketId) or charge code (ChargeCodeId),
                                                                    by Entry Type / Customer rules
                       Lunch (SkipEntryTypes)                    -> skipped
    4. Skips rows that already exist in ConnectWise (same member + same start time) so re-runs are safe.
    5. Preview mode: shows the plan and asks you to type YES.  Auto mode: posts, logs, and emails you a summary.

.EXAMPLE
    .\Sync-CWTimeEntries.ps1 -SaveKey                        # store the private key (encrypted for your Windows account)
    .\Sync-CWTimeEntries.ps1 -ListChargeCodes                # show charge code IDs for the config file
    .\Sync-CWTimeEntries.ps1 -ParseOnly -ReportDate 2026-11-04   # parse the email only, no ConnectWise calls
    .\Sync-CWTimeEntries.ps1 -ReportDate 2026-11-04          # preview + confirm + post
    .\Sync-CWTimeEntries.ps1 -RegisterTask                   # create the daily Windows scheduled task (Auto mode)
#>
[CmdletBinding()]
param(
    [datetime] $ReportDate,                # default: previous workday (Monday -> Friday)
    [ValidateSet('Preview','Auto')] [string] $Mode,
    [string]   $ConfigPath = (Join-Path $PSScriptRoot 'CWTimeConfig.json'),
    [string]   $FromFile,                # test: parse a saved .htm/.txt copy of the email instead of Outlook
    [switch]   $ParseOnly,
    [switch]   $ListChargeCodes,
    [switch]   $SaveKey,
    [switch]   $RegisterTask
)

$ErrorActionPreference = 'Stop'
if (-not $PSBoundParameters.ContainsKey('ReportDate')) {
    $ReportDate = (Get-Date).Date.AddDays(-1)
    while ($ReportDate.DayOfWeek -eq 'Saturday' -or $ReportDate.DayOfWeek -eq 'Sunday') { $ReportDate = $ReportDate.AddDays(-1) }
}
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
$EnUS = [Globalization.CultureInfo]'en-US'

# ------------------------------------------------------------------ config / logging
$Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json
if (-not $Mode) { $Mode = $Config.Mode }
$BaseUri = "$($Config.Server)/v4_6_release/apis/3.0"
$KeyFile = Join-Path $PSScriptRoot 'cw-privatekey.xml'
$LogDir  = Join-Path $PSScriptRoot 'logs'
if (-not (Test-Path $LogDir)) { New-Item -ItemType Directory -Path $LogDir | Out-Null }
$LogFile = Join-Path $LogDir ("CWTime-{0:yyyy-MM-dd}.log" -f $ReportDate)

function Log([string]$Msg, [string]$Color = 'Gray') {
    $line = "{0:yyyy-MM-dd HH:mm:ss}  {1}" -f (Get-Date), $Msg
    Add-Content -Path $LogFile -Value $line
    Write-Host $Msg -ForegroundColor $Color
}

# ------------------------------------------------------------------ one-off actions
if ($SaveKey) {
    Read-Host "Enter TimeEntryAPI Private Key" -AsSecureString | Export-Clixml -Path $KeyFile
    Write-Host "Private key saved to $KeyFile (readable only by $env:USERNAME on $env:COMPUTERNAME)." -ForegroundColor Green
    return
}

if ($RegisterTask) {
    $Exe = (Get-Command pwsh -ErrorAction SilentlyContinue).Source
    if (-not $Exe) { $Exe = 'powershell.exe' }
    $Action  = New-ScheduledTaskAction -Execute $Exe `
               -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Mode Auto" `
               -WorkingDirectory $PSScriptRoot
    $Trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At $Config.ScheduleTime
    $Set     = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 15)
    Register-ScheduledTask -TaskName 'ConnectWise Daily Time Entries' -Action $Action -Trigger $Trigger `
        -Settings $Set -Description 'Posts the previous workday''s Daily Work Activity Summary time entries to ConnectWise' -Force | Out-Null
    Write-Host "Scheduled task 'ConnectWise Daily Time Entries' registered (Mon-Fri at $($Config.ScheduleTime), posts the previous workday)." -ForegroundColor Green
    return
}

# ------------------------------------------------------------------ ConnectWise helpers
function Get-CWHeaders {
    if (Test-Path $KeyFile) { $Secure = Import-Clixml $KeyFile }
    elseif ($Mode -eq 'Auto') { throw "No saved private key. Run: .\Sync-CWTimeEntries.ps1 -SaveKey" }
    else { $Secure = Read-Host "Enter TimeEntryAPI Private Key" -AsSecureString }
    $BSTR = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
    try {
        $Plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR)
        $Enc   = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("$($Config.Company)+$($Config.PublicKey):$Plain"))
    } finally {
        [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR); $Plain = $null
    }
    @{ Authorization = "Basic $Enc"; clientId = $Config.ClientId; Accept = 'application/vnd.connectwise.com+json' }
}

function Invoke-CW([string]$Method, [string]$Path, $Body) {
    $p = @{ Uri = "$BaseUri/$Path"; Method = $Method; Headers = $script:Headers }
    if ($Body) { $p.Body = ($Body | ConvertTo-Json -Depth 5); $p.ContentType = 'application/json' }
    $r = Invoke-RestMethod @p
    foreach ($item in $r) { $item }      # unroll JSON arrays (PowerShell 7 returns them as one object)
}

function Get-CWErrorText($Err) {
    $t = $Err.Exception.Message
    if ($Err.ErrorDetails.Message) { $t += " | " + $Err.ErrorDetails.Message }
    $t
}

function ConvertTo-UtcKey($Value) {
    if ($Value -is [datetime]) { $d = $Value } else { $d = [datetime]::Parse([string]$Value, $EnUS) }
    $d.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm')
}

$TicketTypeCache = @{}
function Get-TicketChargeType([int]$Id) {
    if ($TicketTypeCache.ContainsKey($Id)) { return $TicketTypeCache[$Id] }
    $type = $null
    try   { Invoke-CW GET "service/tickets/$Id`?fields=id" | Out-Null; $type = 'ServiceTicket' }
    catch {
        try { Invoke-CW GET "project/tickets/$Id`?fields=id" | Out-Null; $type = 'ProjectTicket' } catch { }
    }
    $TicketTypeCache[$Id] = $type
    $type
}

if ($ListChargeCodes) {
    $script:Headers = Get-CWHeaders
    Invoke-CW GET 'time/chargeCodes?pageSize=200&orderBy=name' |
        Select-Object id, name, @{n='Billable';e={$_.billTime}}, @{n='Inactive';e={$_.inactiveFlag}} |
        Format-Table -AutoSize
    return
}

# ------------------------------------------------------------------ 1. get the report
function Get-ReportFromOutlook([datetime]$Date) {
    $want = 'Daily Work Activity Summary'
    $dateText = $Date.ToString('dddd, MMMM d, yyyy', $EnUS)
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace('MAPI')
    foreach ($folderId in 6, 5) {               # 6 = Inbox, 5 = Sent Items
        $items = $ns.GetDefaultFolder($folderId).Items
        $items.Sort('[ReceivedTime]', $true)
        $hits = $items.Restrict("@SQL=""urn:schemas:httpmail:subject"" LIKE '%$want%'")
        foreach ($m in $hits) {
            if ($m.Subject -like "*$dateText*") {
                return [pscustomobject]@{ Subject = $m.Subject; Html = $m.HTMLBody; Text = $m.Body }
            }
        }
    }
    $null
}

$CsvFolder = $null
if ($Config.CsvFolder) {
    $CsvFolder = $Config.CsvFolder
    if (-not [IO.Path]::IsPathRooted($CsvFolder)) { $CsvFolder = Join-Path $PSScriptRoot $CsvFolder }
}
$CsvFile = $null
if ($CsvFolder) { $CsvFile = Join-Path $CsvFolder ("CWTime-{0:yyyy-MM-dd}.csv" -f $ReportDate) }

$Rows = @()
if ($CsvFile -and -not $FromFile -and (Test-Path $CsvFile)) {
    # Preferred source: CSV saved to OneDrive by the Cowork daily report
    $Report = [pscustomobject]@{ Subject = "CSV $CsvFile"; Html = $null; Text = $null }
    foreach ($c in (Import-Csv -Path $CsvFile -Encoding UTF8)) {
        $Rows += ,@($c.Start, $c.End, $c.Hours, $c.EntryType, $c.TicketOrMeeting, $c.Customer, $c.Notes)
    }
}
elseif ($FromFile) {
    $raw = Get-Content $FromFile -Raw -Encoding UTF8
    $isHtml = $raw -match '(?i)<table'
    $Report = [pscustomobject]@{ Subject = "Daily Work Activity Summary – " + $ReportDate.ToString('dddd, MMMM d, yyyy', $EnUS)
                                 Html = $(if ($isHtml) { $raw }); Text = $(if (-not $isHtml) { $raw }) }
} else {
    if ($CsvFile) { Log "No CSV at $CsvFile - falling back to Outlook." Yellow }
    try { $Report = Get-ReportFromOutlook $ReportDate }
    catch { Log "Could not read Outlook ($($_.Exception.Message)). Nothing posted." Red; return }
}

if (-not $Report) {
    Log ("No 'Daily Work Activity Summary' found for {0:dddd, MMM d, yyyy}. Nothing posted." -f $ReportDate) Yellow
    return
}
Log "Report: $($Report.Subject)" Cyan

# ------------------------------------------------------------------ 2. parse the time table
function ConvertTo-PlainText([string]$s) {
    $s = $s -replace '(?i)<br\s*/?>', ' ' -replace '<[^>]+>', ''
    $s = [System.Net.WebUtility]::HtmlDecode($s) -replace [char]0xA0, ' '
    ($s -replace '\s+', ' ').Trim()
}

$TimeRx = '^\d{1,2}:\d{2}\s?[AP]M$'

if ($Rows.Count -gt 0) { }   # already loaded from CSV
elseif ($Report.Html -and $Report.Html.IndexOf('CONNECTWISE TIME ENTRIES') -ge 0) {
    $start = $Report.Html.IndexOf('CONNECTWISE TIME ENTRIES')
    $tStart = $Report.Html.IndexOf('<table', $start, [StringComparison]::OrdinalIgnoreCase)
    $tEnd   = $Report.Html.IndexOf('</table>', $tStart, [StringComparison]::OrdinalIgnoreCase)
    $table  = $Report.Html.Substring($tStart, $tEnd - $tStart)
    foreach ($tr in [regex]::Matches($table, '(?is)<tr[^>]*>(.*?)</tr>')) {
        $cells = @([regex]::Matches($tr.Groups[1].Value, '(?is)<t[dh][^>]*>(.*?)</t[dh]>') | ForEach-Object { ConvertTo-PlainText $_.Groups[1].Value })
        if ($cells.Count -ge 7 -and $cells[0] -match $TimeRx) { $Rows += ,$cells }
    }
} else {
    # Plain-text fallback (tab-separated cells, as Outlook renders tables in .Body)
    $text = [string]$Report.Text
    $i = $text.IndexOf('CONNECTWISE TIME ENTRIES')
    if ($i -ge 0) {
        foreach ($line in ($text.Substring($i) -split "`r?`n")) {
            $cells = @($line -split "`t" | ForEach-Object { $_.Trim() })
            if ($cells.Count -ge 7 -and $cells[0] -match $TimeRx) {
                $Rows += ,(@($cells[0..5]) + (($cells[6..($cells.Count-1)]) -join ' '))
            }
        }
    }
}

if ($Rows.Count -eq 0) { Log "Could not find any rows in the CONNECTWISE TIME ENTRIES table. Nothing posted." Red; return }

# ------------------------------------------------------------------ 3. build the plan
$Day = $ReportDate.ToString('yyyy-MM-dd')
$Plan = foreach ($r in $Rows) {
    $s = [datetime]::ParseExact("$Day $($r[0] -replace '\s?([AP]M)',' $1')", 'yyyy-MM-dd h:mm tt', $EnUS)
    $e = [datetime]::ParseExact("$Day $($r[1] -replace '\s?([AP]M)',' $1')", 'yyyy-MM-dd h:mm tt', $EnUS)
    if ($e -le $s) { $e = $e.AddDays(1) }
    $type = $r[3]; $name = $r[4]; $cust = $r[5]; $notes = $r[6]

    $item = [ordered]@{
        Start = $s; End = $e; Hours = [math]::Round(($e - $s).TotalHours, 2)
        Type = $type; Item = $name; Customer = $cust; Notes = $notes
        ChargeToType = $null; ChargeToId = $null; Billable = ''; Action = 'Post'; Reason = ''
    }

    # Billing: any real customer that is not GDS -> billable; GDS / blank / unidentified -> internal default
    $internal = [string]::IsNullOrWhiteSpace($cust) -or $cust -in @('—','-')
    foreach ($pat in $Config.Billing.InternalCustomers) { if ($cust -like $pat) { $internal = $true } }
    if ($internal) { $item.Billable = [string]$Config.Billing.InternalOption }
    else           { $item.Billable = [string]$Config.Billing.CustomerOption }

    if ($Config.SkipEntryTypes -contains $type) {
        $item.Action = 'Skip'; $item.Reason = 'Skipped type'
    }
    elseif ($name -match '#\s*(\d{4,})') {
        $item.ChargeToType = 'Ticket'; $item.ChargeToId = [int]$Matches[1]
    }
    else {
        $rule = $Config.ChargeCodeRules | Where-Object { $type -like $_.EntryType -and $cust -like $_.Customer } | Select-Object -First 1
        if ($rule -and [int]$rule.TicketId -gt 0) {
            $item.ChargeToType = 'Ticket'; $item.ChargeToId = [int]$rule.TicketId
            $prefix = @($name, $cust) | Where-Object { $_ -and $_ -ne '—' -and $_ -ne '-' }
            if ($prefix) { $item.Notes = "[" + ($prefix -join ' | ') + "] " + $notes }
        }
        elseif ($rule -and [int]$rule.ChargeCodeId -gt 0) {
            $item.ChargeToType = 'ChargeCode'; $item.ChargeToId = [int]$rule.ChargeCodeId
            $prefix = @($name, $cust) | Where-Object { $_ -and $_ -ne '—' -and $_ -ne '-' }
            if ($prefix) { $item.Notes = "[" + ($prefix -join ' | ') + "] " + $notes }
        } else {
            $item.Action = 'Skip'; $item.Reason = "No charge code rule for '$type' / '$cust'"
        }
    }
    [pscustomobject]$item
}

function Show-Plan($P) {
    $P | Select-Object @{n='Start';e={$_.Start.ToString('h:mm tt')}}, @{n='End';e={$_.End.ToString('h:mm tt')}}, Hours, Type,
        @{n='ChargeTo';e={ if ($_.ChargeToId) { "$($_.ChargeToType) $($_.ChargeToId)" } }},
        @{n='Customer';e={ if ($_.Customer.Length -gt 28) { $_.Customer.Substring(0,28) } else { $_.Customer } }},
        @{n='Billing';e={ if ($_.Billable) { $_.Billable } else { '(default)' } }},
        Action, Reason | Format-Table -AutoSize | Out-String -Width 220 | Write-Host
}

if ($ParseOnly) {
    Log "Parsed $($Plan.Count) rows (parse only - ConnectWise not contacted)." Cyan
    Show-Plan $Plan
    return
}

# ------------------------------------------------------------------ 4. connect, resolve tickets, de-duplicate
$script:Headers = Get-CWHeaders
try {
    foreach ($p in $Plan | Where-Object { $_.Action -eq 'Post' -and $_.ChargeToType -eq 'Ticket' }) {
        $t = Get-TicketChargeType $p.ChargeToId
        if ($t) { $p.ChargeToType = $t } else { $p.Action = 'Skip'; $p.Reason = "Ticket #$($p.ChargeToId) not found" }
    }

    $fromUtc = $ReportDate.Date.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $toUtc   = $ReportDate.Date.AddDays(1).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    $cond = "member/identifier=`"$($Config.MemberIdentifier)`" and timeStart>=[$fromUtc] and timeStart<[$toUtc]"
    $existing = @(Invoke-CW GET ("time/entries?pageSize=200&fields=id,timeStart&conditions=" + [uri]::EscapeDataString($cond)))
    $existingKeys = @{}
    foreach ($x in $existing) { if ($x.timeStart) { $existingKeys[(ConvertTo-UtcKey $x.timeStart)] = $x.id } }

    foreach ($p in $Plan | Where-Object { $_.Action -eq 'Post' }) {
        $k = ConvertTo-UtcKey $p.Start
        if ($existingKeys.ContainsKey($k)) { $p.Action = 'Skip'; $p.Reason = "Already in ConnectWise (entry $($existingKeys[$k]))" }
    }

    $ToPost = @($Plan | Where-Object { $_.Action -eq 'Post' })
    Log ("Plan: {0} to post ({1:N2} h), {2} skipped." -f $ToPost.Count, ($ToPost | Measure-Object Hours -Sum).Sum, ($Plan.Count - $ToPost.Count)) Cyan
    Show-Plan $Plan

    if ($ToPost.Count -eq 0) { Log "Nothing new to post." Yellow; return }

    if ($Mode -eq 'Preview') {
        if ((Read-Host "Type YES to post these $($ToPost.Count) entries") -ne 'YES') { Log "Cancelled - nothing posted." Yellow; return }
    }

    # -------------------------------------------------------------- 5. post
    $ok = 0; $fail = 0
    foreach ($p in $ToPost) {
        $body = @{
            chargeToType = $p.ChargeToType
            chargeToId   = $p.ChargeToId
            member       = @{ identifier = $Config.MemberIdentifier }
            timeStart    = $p.Start.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
            timeEnd      = $p.End.ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
            notes        = $p.Notes
        }
        if ($p.Billable) { $body.billableOption = $p.Billable }
        if ($p.ChargeToType -like '*Ticket') {
            $body.addToDetailDescriptionFlag = [bool]$Config.TicketNotes.AddToDiscussion
            $body.addToInternalAnalysisFlag  = [bool]$Config.TicketNotes.AddToInternal
            $body.addToResolutionFlag        = $false
        }
        try {
            $r = Invoke-CW POST 'time/entries' $body
            $p.Reason = "Created entry $($r.id)"; $ok++
            Log ("  OK    {0:h:mm tt}-{1:h:mm tt}  {2} {3}  -> entry {4}" -f $p.Start, $p.End, $p.ChargeToType, $p.ChargeToId, $r.id) Green
        } catch {
            $p.Action = 'Failed'; $p.Reason = Get-CWErrorText $_; $fail++
            Log ("  FAIL  {0:h:mm tt}-{1:h:mm tt}  {2} {3}  {4}" -f $p.Start, $p.End, $p.ChargeToType, $p.ChargeToId, $p.Reason) Red
        }
    }
    Log "Done: $ok posted, $fail failed." $(if ($fail) { 'Red' } else { 'Green' })

    # -------------------------------------------------------------- 6. email summary (Auto mode)
    if ($Mode -eq 'Auto' -and $Config.EmailSummary) {
        $rowsHtml = ($Plan | ForEach-Object {
            "<tr><td>{0:h:mm tt}</td><td>{1:h:mm tt}</td><td>{2}</td><td>{3}</td><td>{4} {5}</td><td>{6}</td><td>{7}</td></tr>" -f `
                $_.Start, $_.End, $_.Hours, $_.Type, $_.ChargeToType, $_.ChargeToId, $_.Action,
                [System.Net.WebUtility]::HtmlEncode($_.Reason)
        }) -join ''
        try {
        $ol = New-Object -ComObject Outlook.Application
        $mail = $ol.CreateItem(0)
        $mail.To = $Config.SummaryEmailTo
        $mail.Subject = "ConnectWise time posted for {0:dddd, MMM d}: {1} posted, {2} failed" -f $ReportDate, $ok, $fail
        $mail.HTMLBody = "<p>Report: $([System.Net.WebUtility]::HtmlEncode($Report.Subject))</p><table border=1 cellpadding=4 style='border-collapse:collapse'><tr><th>Start</th><th>End</th><th>Hrs</th><th>Type</th><th>Charged to</th><th>Result</th><th>Detail</th></tr>$rowsHtml</table>"
        $mail.Send()
        } catch {
            Log "Summary email not sent (Outlook unavailable: $($_.Exception.Message)). See this log instead." Yellow
        }
    }
}
catch {
    Log ("ERROR: " + (Get-CWErrorText $_)) Red
    throw
}
finally {
    $script:Headers = $null
}
