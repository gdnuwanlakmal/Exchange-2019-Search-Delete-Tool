# Exchange 2019 Search & Delete Tool
A PowerShell GUI tool for Microsoft Exchange Server 2019 (on-premises only) that allows administrators to search for emails across all mailboxes using message tracking logs, preview affected users, and perform soft or hard delete operations. 
Built and battle-tested on Exchange Server 2019 (Build 15.2.1748.10).

<img width="1381" height="825" alt="image" src="https://github.com/user-attachments/assets/2d7b8001-91a5-4329-acb2-1bcf1c232910" />

### The tool features a dark-themed GUI with:

* Filter panel (Subject, Sender, Recipient, Date Range, Discovery Mailbox)
* Affected mailboxes DataGrid with display name, database, item count, size
* Affected addresses side list (click to copy)
* Colour-coded activity log (green = success, orange = warning, red = error)
* Progress bar during delete operations

### Features

* Search emails across all Exchange transport servers simultaneously
* Filter by Subject (supports Unicode / Sinhala / any language), Sender, Recipient, Date Range
* Resolve affected mailboxes against the Global Address List with 3-stage fallback
* Estimate (dry run) -- shows item counts per mailbox without deleting anything
* Soft Delete -- moves email to user Recoverable Items folder (user can restore via Outlook)
* Hard Delete -- permanently purges email from database (triple confirmation required)
* Export affected mailbox list to TXT or CSV
* Built-in role check and one-click role assignment (Fix Roles button)
* Coloured activity log with timestamps saved in the UI

### Requirements

* Windows Server with Exchange Server 2019 (on-premises)
* Must be run from Exchange Management Shell (EMS) as Administrator
* PowerShell 5.1 (included with Windows Server 2016/2019)
* .NET Framework 4.7.2 or later
* A configured Discovery Search Mailbox (created by default in Exchange 2019)

### Required Roles
These roles must be assigned to the account running the script.
The tool includes a Fix Roles button that assigns them automatically,
but a full Windows logoff/logon is required afterwards.

Run once in Exchange Management Shell before first use:
```shell
New-ManagementRoleAssignment -Role "Mailbox Search"        -User "Administrator"
New-ManagementRoleAssignment -Role "Mailbox Import Export" -User "Administrator"
Add-RoleGroupMember "Discovery Management" -Member "Administrator"
```
### IMPORTANT: After assigning roles, you must LOG OFF from Windows on the
### Exchange Server and log back on. Simply closing and reopening EMS is not
### enough because the Kerberos security token is not refreshed until logoff. 

### Installation

1. Download ExchangeTool_v6.ps1
2. Copy to your Exchange Server desktop
3. Open Exchange Management Shell as Administrator
4. Run:

```shell
.\ExchangeTool_v6.ps1
```
### If you see a security warning, press R to run once, or unblock permanently:

```shell
Unblock-File .\ExchangeTool_v6.ps1
.\ExchangeTool_v6.ps1
```
### Usage
### Step 1 -- Connect
Click the Connect button. The tool loads the Exchange snap-in and checks
whether all required roles and cmdlets are available in the current session.

### Step 2 -- Fix Roles (first time only)
If the role check shows [MISSING], click Fix Roles. The tool assigns the
required roles automatically. Then:

* Close the script window
* Log off from Windows on the Exchange Server
* Log back on
* Reopen EMS and run the script again

### Step 3 -- Search
Enter at least one filter:

* Subject -- partial match, supports any language including Unicode
* Sender -- partial email address match
* Recipient -- partial email address match
* From Date / To Date -- message tracking date range

Click Search. The tool scans transport logs on all Exchange servers
and resolves each affected address against the GAL.

### Step 4 -- Estimate (recommended before deleting)
Click Estimate to see exactly how many items would be deleted per mailbox.
Nothing is deleted. Verify the counts match your expectation before proceeding.

### Step 5a -- Soft Delete
Click 3a. Soft Delete. Matching emails are:

1. Copied to the Discovery Search Mailbox as a backup
2. Moved to each user's Recoverable Items folder

Users can still recover the email via Outlook > Recover Deleted Items
for up to 14 days (or 30 days if your organisation has extended this).

### Step 5b -- Hard Delete
Click 3b. Hard Delete. Matching emails are permanently purged from the
Exchange database. This requires three confirmations:

1. First warning dialog -- click Yes
2. Second confirmation dialog -- click OK
3. Type HARDDELETE (all caps) in the text box and press OK


There is no recovery from Hard Delete. Use Estimate and Soft Delete first
to verify you are targeting the correct emails.

### Export
Click Export to save the list of affected mailboxes as a TXT or CSV file.
The export includes email address, display name, database, and GAL status.

### How It Works
### Search
The tool queries Get-MessageTrackingLog on every transport server in the
organisation across the specified date range. Results are filtered in memory
against the Subject, Sender, and Recipient criteria. Unique email addresses
from matching log entries are then resolved against Exchange using
Get-Mailbox with a three-stage fallback:

1. Direct identity lookup
2. EmailAddresses filter
3. Ambiguous Name Resolution (ANR)

### Delete
Both Soft Delete and Hard Delete use the Search-Mailbox cmdlet with
the -SearchQuery parameter (KQL syntax) and -DeleteContent -Force.

The query is passed as a native PowerShell variable (not a constructed
string) to preserve Unicode characters correctly. This is important for
non-Latin subjects such as Sinhala, Arabic, Chinese, etc.

Soft Delete additionally passes -TargetMailbox and -TargetFolder
to save a copy to the Discovery Search Mailbox before deletion.

Hard Delete omits the target mailbox parameters, causing Exchange to
purge the items permanently from the database.

### Why Invoke-Expression was removed
Earlier versions used Invoke-Expression to build the Search-Mailbox command
as a text string. This caused Unicode subjects to be mangled during string
construction, resulting in the cmdlet returning Success: True with 0 items
deleted. The fix was to call Search-Mailbox directly as a cmdlet using the
$query variable, which preserves all Unicode characters natively.

### Troubleshooting
### -DeleteContent parameter not found
This means the Mailbox Import Export role is assigned but the Windows
Kerberos token has not been refreshed.
Solution:

1. Click Fix Roles (roles may already be assigned, that is fine)
2. Close the script and EMS
3. Log off from Windows on the Exchange Server
4. Log back on as Administrator
5. Open EMS and run the script again

### Search finds emails but delete shows 0 items
The KQL query is not matching inside the mailbox index. This can happen if:

* The email subject contains special characters being escaped in the query
* The date range in the query does not align with the mailbox index timezone
* The email has not yet been indexed by Exchange Search

Run Estimate first to confirm items are found. If Estimate shows 0 but
transport logs show matches, the email may be in Sent Items rather than
Inbox. Try removing the date range filter to broaden the search.

### ToMB method not found
This occurs when Get-MailboxStatistics returns a Deserialized object from
a remote session. The tool handles this automatically by parsing the size
string as a fallback.

### Script runs but GUI does not appear
Ensure you are running from Exchange Management Shell, not standard
PowerShell. The Exchange snap-in must be registered on the machine.

### Security Notes

* The script does not connect to the internet or send data externally
* No credentials are stored or logged
* All operations run under the currently logged-on Windows session
* Soft Delete is reversible by the user via Outlook
* Hard Delete is permanent and requires triple confirmation including
* typing HARDDELETE in a confirmation box
* The script only calls Exchange cmdlets: Search-Mailbox,
* Get-Mailbox, Get-MessageTrackingLog, Get-MailboxStatistics,
* New-ManagementRoleAssignment, Add-RoleGroupMember
* It does not modify user accounts, mailbox settings, or Active Directory


### Delete Type Comparison
<img width="668" height="281" alt="image" src="https://github.com/user-attachments/assets/1c5421d3-ef35-4dbf-b9f2-07b5c839f7b2" />


### Exchange Version Compatibility
<img width="522" height="192" alt="image" src="https://github.com/user-attachments/assets/0dcbf10a-8520-4a4f-943b-fe62c79861f8" />


### Known Limitations

* Search-Mailbox returns a maximum of 10,000 items per mailbox per run.
For mailboxes with more than 10,000 matching items, run the delete
multiple times until Estimate returns 0.

* Transport log scanning across large organisations with high mail volume
can take several minutes (tested: ~167,000 entries across 3 servers
in approximately 6 minutes).

* The tool targets emails in all folders including Inbox, Sent Items,
Deleted Items, and subfolders. Use specific Subject/Sender/Recipient
filters to avoid unintended matches.


### License
MIT License. Free to use, modify, and distribute.
Not affiliated with or endorsed by Microsoft.

### Author Notes

This tool was developed and tested iteratively on a live Exchange 2019
on-premises deployment. The key technical challenges solved during
development were:

1. Role assignment requires a full Windows logoff to refresh the Kerberos
token before -DeleteContent becomes available in Search-Mailbox.

2. Unicode subjects (Sinhala script in this case) must be passed as a
native PowerShell variable to Search-Mailbox. Building the command
via Invoke-Expression mangles non-ASCII characters and causes silent
zero-item matches.

3. Get-MailboxStatistics returns a Deserialized object in some session
configurations, requiring a string-parsing fallback for TotalItemSize.
