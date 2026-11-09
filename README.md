# ConnectWise Daily Time Entry Automation

Posts your time to ConnectWise every weekday morning, using the
"Daily Work Activity Summary" report that Copilot Cowork builds from your
email, Teams, calendar and files.

How it works:

1. **8:00 AM** (Mon–Fri): a Copilot Cowork scheduled task emails you the
   Daily Work Activity Summary for the previous workday. It also saves the
   time-entry table as `Reports\CWTime-yyyy-MM-dd.csv` in your OneDrive.
2. **9:00 AM** (Mon–Fri): a Windows scheduled task on your PC reads that CSV
   and posts each row to ConnectWise. On Monday it posts Friday.

| Row in the report | Goes to |
|---|---|
| Has a ticket number (`#1234567`) | That ticket (service or project) |
| Admin | Your Admin ticket |
| Meeting | Your Meeting ticket |
| Project Work with no ticket | The charge code you pick (or skipped) |
| Lunch | Skipped |

Billing: rows for a real customer are posted **Billable**. Rows for internal
customers (GDS…, Internal, Unidentified…) use ConnectWise's default billing.

Re-running a day is safe: entries that start at the same time as an entry
already in ConnectWise are skipped.

---

## What you need first

- **Windows PC** on the GDS network or VPN, with **OneDrive for Business** syncing.
- **PowerShell 7** (recommended): `winget install Microsoft.PowerShell`
- **Copilot Cowork** access.
- **ConnectWise API keys** (public + private) from your ConnectWise admin.
  The API member needs rights to read tickets and charge codes and to
  add time entries. Get the private key from the admin directly. Don't
  receive it by email or chat.
- **Two ConnectWise tickets**: one for your Admin time and one for your
  Meeting time. Note both ticket numbers.
- Your **ConnectWise member identifier** (your CW login name, e.g. `chadl`).

## Setup (about 10 minutes)

1. Extract this zip to any temporary folder.
2. Open **PowerShell 7** (not as Administrator) in that folder and run:

   ```powershell
   Get-ChildItem | Unblock-File
   .\Install-CWTimeEntries.ps1
   ```

   Answer the questions. The installer:
   - installs into `OneDrive\Documents\CW Time Entry Automation` (you can change it)
   - saves your private key encrypted for your Windows account only
   - tests the ConnectWise connection and lists charge codes so you can pick
     one for Project Work
   - creates the scheduled task **ConnectWise Daily Time Entries**
     (Mon–Fri, 9:00 AM)
   - writes your personal **Daily-Report-Prompt.txt** and copies it to your
     clipboard

3. **Create the Cowork report.** Open Copilot Cowork, start a new task, paste
   the prompt (it's on your clipboard, or open `Daily-Report-Prompt.txt`) and
   send it. Make sure Cowork sets it up as a scheduled task that runs every
   weekday at 8:00 AM.

4. **Test it** the morning after the first report arrives. In the install
   folder:

   ```powershell
   .\Sync-CWTimeEntries.ps1 -ParseOnly      # read the CSV and show the plan; posts nothing
   .\Sync-CWTimeEntries.ps1                 # preview, type YES to post
   ```

   After that, the scheduled task posts automatically each weekday.

## Everyday use

| Task | Command (run in the install folder) |
|---|---|
| See what would be posted for a day | `.\Sync-CWTimeEntries.ps1 -ParseOnly -ReportDate 2026-11-06` |
| Post a specific day with a preview | `.\Sync-CWTimeEntries.ps1 -ReportDate 2026-11-06` |
| List charge codes | `.\Sync-CWTimeEntries.ps1 -ListChargeCodes` |
| Replace the saved private key | `.\Sync-CWTimeEntries.ps1 -SaveKey` |
| Run the scheduled task now | `Start-ScheduledTask -TaskName 'ConnectWise Daily Time Entries'` |
| Check results | open `logs\CWTime-yyyy-MM-dd.log` |

**Manual runs** always show a preview and wait for `YES`. The **scheduled task**
posts without asking.

To post a day by hand that has no CSV, open the report email, copy the
CONNECTWISE TIME ENTRIES table into a CSV in the format below, save it in
`Reports\`, and run that day.

CSV format: `Reports\CWTime-yyyy-MM-dd.csv`

```
"Date","Start","End","Hours","EntryType","TicketOrMeeting","Customer","Notes"
"2026-11-06","8:00 AM","9:00 AM","1.00","Admin","—","GDS Internal","Email, Teams, ticket queue review"
```

## Changing settings (`CWTimeConfig.json`)

| Setting | What it does |
|---|---|
| `ChargeCodeRules` | Where rows without a ticket go. Rules are checked from top to bottom; the first match wins. Use `TicketId` or `ChargeCodeId`; `0` means skip. `Customer` accepts wildcards, e.g. `"Royal*"`. |
| `Billing.InternalCustomers` | Customer names that are **not** billable. |
| `SkipEntryTypes` | Row types never posted (default: Lunch). |
| `TicketNotes` | Set to `true` to also add the note to the ticket's discussion or internal notes. |
| `EmailSummary` | Emails a results summary after each run. Works only with **classic** Outlook; leave `false` if you use new Outlook. |
| `ScheduleTime` | Used when the task is (re)created with `.\Sync-CWTimeEntries.ps1 -RegisterTask`. |

## Troubleshooting

| Symptom | Fix |
|---|---|
| "…is not digitally signed" | `Unblock-File .\Sync-CWTimeEntries.ps1` (if your PC requires signed scripts by Group Policy, ask IT to sign it) |
| 401 Unauthorized | Wrong company ID / public key / private key. Re-run `-SaveKey` or fix the config. |
| 403 Forbidden | The API member's security role is missing a permission (tickets, charge codes or time entry add). |
| "No CSV at …Reports\CWTime-….csv" | The Cowork report didn't save the CSV, or OneDrive hasn't synced yet. Check the Reports folder in OneDrive online. |
| Task didn't run | The PC was off or you weren't signed in. It runs at your next sign-in. Check Task Scheduler → *ConnectWise Daily Time Entries* → History. |
| "Summary email not sent (Outlook unavailable)" | Only the summary email failed; your time was posted. You're on new Outlook or PowerShell is running as Administrator. Set `EmailSummary` to `false`. |
| A row was skipped: "No charge code rule…" | Add a rule for that EntryType in `ChargeCodeRules`. |

## Files

| File | Purpose |
|---|---|
| `Install-CWTimeEntries.ps1` | One-time setup (run from the extracted zip) |
| `Sync-CWTimeEntries.ps1` | The posting script (copied to your install folder) |
| `Daily-Report-Prompt-TEMPLATE.txt` | Cowork report instructions with placeholders (the installer fills them in) |
| `CWTimeConfig.json` | Your settings (created by the installer) |
| `cw-privatekey.xml` | Your encrypted private key (created by the installer; works only for your Windows account on that PC; never share it) |
