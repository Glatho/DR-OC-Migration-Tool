# Teams DID Migration Toolkit

`Teams-DID-Migration-Toolkit.ps1` — an interactive PowerShell tool for migrating a Microsoft Teams tenant's phone numbers (DIDs) from **Direct Routing** to **Operator Connect**.

It has three menu-driven actions: back up every phone number in the tenant, unassign (remove) DIDs from users, and reassign DIDs from a backup once your Operator Connect provider has ported them in.

## Contents

- [What it does](#what-it-does)
- [Prerequisites](#prerequisites)
- [Running the script](#running-the-script)
- [The migration workflow](#the-migration-workflow)
- [Menu options in detail](#menu-options-in-detail)
  - [1) Backup](#1-backup)
  - [2) Remove DIDs](#2-remove-dids)
  - [3) Assign DIDs](#3-assign-dids)
- [Output files](#output-files)
- [Safety features](#safety-features)
- [Troubleshooting](#troubleshooting)
- [Known limitations](#known-limitations)

## What it does

| Option | Action |
|---|---|
| **1) Backup** | Exports every phone number in the tenant — regardless of whether it's assigned to a user, a resource account (auto attendant/call queue), a policy, or sitting unassigned — to a timestamped CSV in this folder. |
| **2) Remove DIDs** | Unassigns DIDs from users/resource accounts. This is also the step that **frees a number for your Operator Connect provider to add/port in** — Microsoft's documented migration process requires a number to be unassigned from Direct Routing before an operator can add it as an OC number. |
| **3) Assign DIDs** | Reads the latest backup for the currently-connected tenant, finds which of those numbers are now sitting unassigned in the tenant (i.e. have landed as Operator Connect numbers after porting), and reassigns them to the same users. |

## Prerequisites

- **MicrosoftTeams PowerShell module.** The script checks for it and offers to install it from PSGallery if missing. `Set/Get/Remove-CsPhoneNumberAssignment` need module version 3.0.0+.
- **Teams Administrator or Global Administrator** role on the account you sign in with.
- **Windows PowerShell 5.1** is recommended for the interactive grid picker (`Out-GridView`). If it's not available (e.g. plain PowerShell 7 without the GraphicalTools module), the script automatically falls back to a numbered list you can type selections into (e.g. `1,3,5-8`).

## Running the script

```powershell
.\Teams-DID-Migration-Toolkit.ps1
```

You'll be prompted to sign in — a browser window opens for you to authenticate as a Teams/Global Administrator. Once connected, the script reads the tenant's name and uses it to tag every file this run produces, so backups and reports from different customer tenants never get mixed up if they land in the same folder.

## The migration workflow

The three menu options are designed to be run in this order over the course of a migration:

1. **Run option 1** to snapshot the current (Direct Routing) assignments. Keep this file — option 3 depends on it later.
2. **Run option 2** to unassign the old Direct Routing DIDs from their users. This is what frees each number so your operator's side can pick it up.
3. **Submit/confirm the port order with your Operator Connect provider** and wait for them to complete it. This is a manual, external step the script cannot perform — it can take anywhere from minutes to days depending on the provider. Once done, the numbers will appear in the tenant as unassigned Operator Connect numbers (check **Teams admin center > Voice > Phone numbers**).
4. **Run option 3** to reassign the now-Operator-Connect numbers back to the same users, using the option 1 backup as the source of truth.

Steps 1–2 can be run in small test batches first (a couple of numbers) before committing to a full cutover.

## Menu options in detail

### 1) Backup

Pulls **every** number in the tenant (no status filter), so the export includes:

- Numbers assigned to users
- Numbers assigned to resource accounts (auto attendants, call queues)
- Numbers assigned to policies (e.g. shared calling routing policy emergency numbers)
- Numbers sitting completely unassigned

Each row includes: `UserPrincipalName`, `DisplayName`, `AccountType` (User / Auto Attendant / Call Queue / Resource Account), `TelephoneNumber`, `NumberType` (DirectRouting / CallingPlan / OperatorConnect), `PstnAssignmentStatus`, `Provider`, `LocationId`, voice policy details (`OnlineVoiceRoutingPolicy`, `TeamsCallingPolicy`, `TenantDialPlan`), and more.

The log line after running shows a breakdown by status, e.g. `Breakdown by status: UserAssigned=142, Unassigned=8, VoiceApplicationAssigned=3` — a quick sanity check that the count matches what you'd expect from Teams admin center.

Saved as `<Tenant>_TeamsPhoneNumberBackup_<timestamp>.csv`.

### 2) Remove DIDs

Three ways to pick which DIDs to remove:

1. **Type/paste numbers** — best for 1–10 DIDs.
2. **Import a CSV** with a `TelephoneNumber` column — best for large bulk removals (300–400+).
3. **Interactive picker** — browse every currently-assigned DID and multi-select (grid view, or numbered fallback).

After selecting, you'll be shown a summary table and asked to type:

- `REMOVE` — apply the changes for real, or
- `PREVIEW` — dry-run everything (logs what *would* happen, makes no changes), or
- anything else / Enter — cancel.

If you type `REMOVE`, you're then asked two more things:

- Whether to send Teams' built-in email notification to each affected user about the removal.
- Whether to also clear each user's **Online Voice Routing Policy** for removed Direct Routing numbers — this is Microsoft's documented Step 2 for this migration (that policy points at the old DR trunk and has no purpose once the number's gone). Only applied to `DirectRouting` numbers assigned to users.

After a real (non-preview) run with successes, the script prints a "next steps" reminder about coordinating the port with your operator before running option 3.

### 3) Assign DIDs

- Finds the most recent backup CSV **for the tenant you're currently connected to**.
- Asks which number type to look for in the tenant's unassigned inventory (defaults to `OperatorConnect`).
- Matches backup rows against numbers that are now sitting unassigned of that type, and shows you the matches to multi-select.
- Same `ASSIGN` / `PREVIEW` / cancel confirmation as option 2, and the same optional notification prompt.

## Output files

Every file this script produces is prefixed with the connected tenant's name, e.g. `ACCESS-4-PTY-LTD_...`.

| File pattern | What it is |
|---|---|
| `<Tenant>_TeamsPhoneNumberBackup_<timestamp>.csv` | Full inventory snapshot from option 1. |
| `<Tenant>_RemoveDidsPlan_<runid>.csv` / `<Tenant>_AssignDidsPlan_<runid>.csv` | The full intended target list for a bulk run, written before any changes are made. |
| `<Tenant>_RemoveDidsResults_<runid>.csv` / `<Tenant>_AssignDidsResults_<runid>.csv` | What actually happened, one row appended per number as it's processed — includes `Status` (Success/Failed/DryRun), `Verified` (Yes/No/Skipped), and for removals, `VoiceRoutingPolicyCleared`. |
| `TeamsDIDToolkit_<timestamp>.log` | Full timestamped log of everything the script did during that session. |

**Phone numbers in these CSVs are written in an Excel-safe format** (`="+61399672273"`). This stops Excel from silently stripping the leading `+` if someone opens and re-saves the file. The script automatically un-wraps this when it reads its own CSVs back in — you don't need to do anything, but don't manually strip the `="..."` wrapper yourself if you're editing a CSV by hand before feeding it back into option 2.

## Safety features

- **Preview/dry-run mode** on every bulk action — see exactly what would happen before committing.
- **Explicit typed confirmation** (`REMOVE` / `ASSIGN`) required before any change is made.
- **Post-change verification** — after every removal/assignment, the script re-checks live tenant data to confirm the change actually took effect (Teams number changes can take a few seconds to propagate), and records this in the results CSV.
- **Resumable runs** — if a bulk run is interrupted (closed window, crash, network drop), the next time you pick the same menu option the script detects the incomplete run and offers to resume just the numbers that didn't already succeed, rather than starting over. This also means a `PREVIEW` run can later be "resumed" as a real `REMOVE`/`ASSIGN` run.
- **Tenant-tagged filenames** — prevents a backup or report from one customer tenant ever being matched against another if both land in this folder.
- **Full audit trail** — a Plan file (intent) and Results file (outcome) for every bulk action, plus a session log.

## Troubleshooting

**"Cannot process argument transformation on parameter 'ResultSize'"** or **"A parameter cannot be found that matches parameter name 'NumberType'/'TelephoneNumber'/'PhoneNumber'"**
Different builds of the MicrosoftTeams PowerShell module have used different parameter names for the same thing over time (e.g. `TelephoneNumber` vs `PhoneNumber`, `NumberType` vs `PhoneNumberType`). The script resolves the actual parameter names your installed module uses at connect time via `Get-Command`, and logs what it found right after signing in — look for a log line like:

```
Remove-CsPhoneNumberAssignment parameter names on this module: PhoneNumber=[PhoneNumber] NumberType=[PhoneNumberType]
```

If you still hit a "parameter cannot be found" error on a *different* parameter name (e.g. `LocationId`, `AssignmentCategory`), it's the same underlying issue on a parameter the script doesn't yet resolve dynamically — note the exact parameter name from the error for a fix.

**Blank `UserPrincipalName` / `DisplayName` / `AccountType` in the backup**
Usually means the identity lookup couldn't match any users. Check the log for a line like `Retrieved N user(s), resolved an identity key for 0 of them` — if it says 0, run `Get-CsOnlineUser -ResultSize 1 | Get-Member` and check what identity-like property names (`Identity`, `ObjectId`, etc.) are actually present on your module version.

**Fewer numbers than expected in the backup**
Check the "Breakdown by status" log line after a backup run. If a status you expected (e.g. `VoiceApplicationAssigned` for resource accounts) shows 0 when you know numbers exist, verify with `Get-CsOnlineApplicationInstance` directly.

**Module install fails / PSGallery untrusted prompt**
The script calls `Install-Module -Repository PSGallery`, which may prompt to trust the repository the first time. Approve it, or run `Set-PSRepository -Name PSGallery -InstallationPolicy Trusted` first if you manage this centrally.

**`Out-GridView` not available**
The script automatically falls back to a numbered picker (type `1,3,5-8` or `all`) — no action needed, but note it lacks sorting/filtering, so for very large lists a CSV import (option 2, method 2) may be easier to work with.

## Known limitations

- The actual number **porting into Operator Connect is done by your provider**, outside this script, on their own timeline — the script can only tell you when a number has landed (it'll show up as unassigned `OperatorConnect` in the tenant).
- Numbers assigned to **Teams shared calling routing policy instances** (rather than a user or resource account) won't resolve to a friendly name in reports — they'll show the raw policy instance ID instead.
- This script has not been tested against GCC High / DoD / 21Vianet clouds — Direct Routing/Operator Connect availability and cmdlet behaviour can differ there.
