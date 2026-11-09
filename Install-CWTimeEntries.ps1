<#
.SYNOPSIS
    One-time setup for the ConnectWise Daily Time Entry automation.

.DESCRIPTION
    - Copies Sync-CWTimeEntries.ps1 to your install folder (inside OneDrive) and creates Reports\ and logs\
    - Builds CWTimeConfig.json from your answers
    - Stores your ConnectWise API private key encrypted for your Windows account only
    - Tests the ConnectWise connection and lets you pick a charge code for Project Work
    - Registers the Windows scheduled task (Mon-Fri, posts the previous workday)
    - Writes Daily-Report-Prompt.txt (personalised) and copies it to your clipboard
      so you can paste it into a Copilot Cowork scheduled task.

.EXAMPLE
    .\Install-CWTimeEntries.ps1
#>
[CmdletBinding()]
param([string] $InstallPath)

$ErrorActionPreference = 'Stop'
$Src = $PSScriptRoot

function Ask([string]$Prompt, [string]$Default) {
    $suffix = ''
    if ($Default) { $suffix = " [$Default]" }
    $a = Read-Host "$Prompt$suffix"
    if ([string]::IsNullOrWhiteSpace($a)) { return $Default }
    $a.Trim()
}
function AskInt([string]$Prompt, [int]$Default = -1) {
    while ($true) {
        $d = ''; if ($Default -ge 0) { $d = "$Default" }
        $a = Ask $Prompt $d
        $n = 0
        if ([int]::TryParse(($a -replace '^#',''), [ref]$n)) { return $n }
        Write-Host "  Please enter a number." -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "=== ConnectWise Daily Time Entry automation - setup ===" -ForegroundColor Cyan
Write-Host "Press Enter to accept the value in [brackets]."
Write-Host ""

# ------------------------------------------------------------------ install folder (must be inside OneDrive)
$OneDriveRoot = $env:OneDriveCommercial
if (-not $OneDriveRoot) { $OneDriveRoot = $env:OneDrive }
if (-not $InstallPath) {
    if ($OneDriveRoot) { $def = Join-Path $OneDriveRoot 'Documents\CW Time Entry Automation' }
    else               { $def = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'CW Time Entry Automation' }
    $InstallPath = Ask 'Install folder (must be inside your work OneDrive)' $def
}
foreach ($d in $InstallPath, (Join-Path $InstallPath 'Reports'), (Join-Path $InstallPath 'logs')) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d | Out-Null }
}

# ------------------------------------------------------------------ about you
Write-Host "`n-- About you --" -ForegroundColor Cyan
$upn = $null
try { $upn = (whoami /upn 2>$null) } catch { }
$Email   = Ask 'Your work email' $upn
$Name    = Ask 'Your full name (as it should appear in the report)' ''
$Title   = Ask 'Your job title' ''
$OrgName = Ask 'Company name' 'Global Data Systems'
$Member  = Ask 'Your ConnectWise member identifier (your CW login name)' (($Email -split '@')[0])

# ------------------------------------------------------------------ ConnectWise API
Write-Host "`n-- ConnectWise API (get these from your ConnectWise admin) --" -ForegroundColor Cyan
$Server    = Ask 'ConnectWise URL' 'https://connectwise.getgds.com'
$Company   = Ask 'ConnectWise company ID' 'GDS'
$ClientId  = Ask 'ConnectWise clientId' 'f41f250c-b38c-4806-abf9-28a6c339eb89'
$PublicKey = ''
while (-not $PublicKey) { $PublicKey = Ask 'API public key' '' }
$KeyFile = Join-Path $InstallPath 'cw-privatekey.xml'
Read-Host 'API private key (hidden; stored encrypted for your Windows account only)' -AsSecureString | Export-Clixml -Path $KeyFile

# ------------------------------------------------------------------ where time goes
Write-Host "`n-- Where time is posted --" -ForegroundColor Cyan
Write-Host "Rows that already contain a ticket number (#1234567) always go to that ticket."
$AdminTicket   = AskInt 'Ticket number for ADMIN time (email, Teams, queue review)'
$MeetingTicket = AskInt 'Ticket number for MEETING time'
$Internal      = Ask 'Customer names treated as internal / not billable (comma-separated, * = wildcard)' 'GDS*,Global Data Systems*,Internal,Unidentified*'
$RunTime       = Ask 'Time to post each weekday (after your morning report arrives)' '9:00AM'
$EmailSummary  = (Ask 'Email yourself a summary after each run? Needs CLASSIC Outlook (y/n)' 'n') -match '^[Yy]'

# ------------------------------------------------------------------ install files + config
Copy-Item (Join-Path $Src 'Sync-CWTimeEntries.ps1') $InstallPath -Force
$Sync      = Join-Path $InstallPath 'Sync-CWTimeEntries.ps1'
$ConfigOut = Join-Path $InstallPath 'CWTimeConfig.json'

function Save-Config([int]$ProjectCode) {
    $cfg = [ordered]@{
        Server           = $Server
        Company          = $Company
        PublicKey        = $PublicKey
        ClientId         = $ClientId
        MemberIdentifier = $Member
        CsvFolder        = 'Reports'
        Mode             = 'Preview'
        ScheduleTime     = $RunTime
        EmailSummary     = $EmailSummary
        SummaryEmailTo   = $Email
        SkipEntryTypes   = @('Lunch')
        Billing          = [ordered]@{
            CustomerOption    = 'Billable'
            InternalOption    = ''
            InternalCustomers = @($Internal -split '\s*,\s*' | Where-Object { $_ })
        }
        TicketNotes      = [ordered]@{ AddToDiscussion = $false; AddToInternal = $false }
        ChargeCodeRules  = @(
            [ordered]@{ EntryType = 'Admin';        Customer = '*'; TicketId     = $AdminTicket },
            [ordered]@{ EntryType = 'Meeting';      Customer = '*'; TicketId     = $MeetingTicket },
            [ordered]@{ EntryType = 'Project Work'; Customer = '*'; ChargeCodeId = $ProjectCode },
            [ordered]@{ EntryType = '*';            Customer = '*'; ChargeCodeId = 0 }
        )
    }
    $cfg | ConvertTo-Json -Depth 6 | Set-Content -Path $ConfigOut -Encoding UTF8
}
Save-Config 0
Get-ChildItem $InstallPath -File | Unblock-File

# ------------------------------------------------------------------ test connection + choose charge code
Write-Host "`n-- Testing ConnectWise connection --" -ForegroundColor Cyan
$ProjectCode = 0
try {
    & $Sync -ListChargeCodes -ConfigPath $ConfigOut
    Write-Host "Connected." -ForegroundColor Green
    $ProjectCode = AskInt 'Charge code ID for PROJECT WORK rows with no ticket (0 = skip those rows)' 0
}
catch {
    Write-Host "Could not reach ConnectWise: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Check the keys / VPN, then fix CWTimeConfig.json or re-run setup. Continuing." -ForegroundColor Yellow
}
Save-Config $ProjectCode

# ------------------------------------------------------------------ scheduled task
Write-Host "`n-- Creating the scheduled task --" -ForegroundColor Cyan
try { & $Sync -RegisterTask -ConfigPath $ConfigOut }
catch { Write-Host "Could not create the scheduled task: $($_.Exception.Message)" -ForegroundColor Red }

# ------------------------------------------------------------------ personalised Cowork prompt
$ReportsRel = $null
if ($OneDriveRoot -and $InstallPath.StartsWith($OneDriveRoot, [StringComparison]::OrdinalIgnoreCase)) {
    $ReportsRel = ($InstallPath.Substring($OneDriveRoot.Length).TrimStart('\') -replace '\\','/') + '/Reports'
} else {
    $ReportsRel = '<<YOUR ONEDRIVE PATH>>/Reports'
    Write-Host "WARNING: the install folder is not inside OneDrive - Cowork cannot deliver the CSV there." -ForegroundColor Yellow
}
$TzWin = (Get-TimeZone).Id
$Tz = $TzWin
$iana = $null
try { if ([TimeZoneInfo]::TryConvertWindowsIdToIanaId($TzWin, [ref]$iana)) { $Tz = $iana } } catch { }
$First = ($Name -split '\s+')[0]
if (-not $First) { $First = 'me' }

$tpl = Get-Content (Join-Path $Src 'Daily-Report-Prompt-TEMPLATE.txt') -Raw -Encoding UTF8
$prompt = $tpl
$map = [ordered]@{ '{{NAME}}' = $Name; '{{FIRST_NAME}}' = $First; '{{EMAIL}}' = $Email; '{{TITLE}}' = $Title
                   '{{COMPANY}}' = $OrgName; '{{TIMEZONE}}' = $Tz; '{{REPORTS_PATH}}' = $ReportsRel }
foreach ($k in $map.Keys) { $prompt = $prompt.Replace($k, [string]$map[$k]) }
$PromptFile = Join-Path $InstallPath 'Daily-Report-Prompt.txt'
Set-Content -Path $PromptFile -Value $prompt -Encoding UTF8
try { Set-Clipboard -Value $prompt } catch { }

# ------------------------------------------------------------------ done
Write-Host ""
Write-Host "=== Setup complete ===" -ForegroundColor Green
Write-Host "Installed to : $InstallPath"
Write-Host "CSV folder   : $ReportsRel  (in OneDrive)"
Write-Host "Task         : 'ConnectWise Daily Time Entries' - Mon-Fri at $RunTime, posts the previous workday"
Write-Host ""
Write-Host "NEXT STEPS" -ForegroundColor Cyan
Write-Host " 1. In Copilot Cowork, start a new task and paste the prompt (already on your clipboard;"
Write-Host "    also saved as Daily-Report-Prompt.txt). Ask Cowork to run it every weekday at 8:00 AM."
Write-Host " 2. After the first report arrives, test without posting:"
Write-Host "      cd `"$InstallPath`""
Write-Host "      .\Sync-CWTimeEntries.ps1 -ParseOnly"
Write-Host " 3. Then post once by hand with a preview (type YES to confirm):"
Write-Host "      .\Sync-CWTimeEntries.ps1"
Write-Host " From then on the scheduled task posts automatically. Check the logs folder for results."
