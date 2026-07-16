<#
================================================================================
MICROSOFT TEAMS PHONE NUMBER (DID) MIGRATION TOOLKIT
================================================================================
PURPOSE:
  Interactive toolkit for migrating a Teams tenant's PSTN phone numbers (DIDs)
  from Direct Routing to Operator Connect. Provides three menu-driven actions:

    1) BACKUP  - Exports EVERY phone number in the tenant regardless of
                 assignment status (assigned to a user, assigned to a
                 resource account, assigned to a policy, or sitting
                 unassigned), its type (Direct Routing / Calling Plan /
                 Operator Connect), provider, and the owning user's or
                 resource account's details, to a timestamped CSV saved in
                 this script's folder.

    2) REMOVE  - Unassigns DIDs from users/resource accounts. This is also the
                 step that FREES the number for your Operator Connect provider
                 to add/port into the tenant - Microsoft's documented DR ->
                 OC migration process requires the number to be unassigned
                 from Direct Routing before the operator can add it as an OC
                 number. Works for small batches (1-2 DIDs, typed by hand) up
                 to large bulk runs (300-400+, via CSV import or an
                 interactive multi-select picker). Numbers are unassigned,
                 NOT deleted from the tenant - Remove-CsPhoneNumberAssignment
                 keeps them available in the tenant's number inventory for
                 later reassignment. Optionally also clears the affected
                 users' Online Voice Routing Policy (Microsoft's documented
                 Step 2 for this migration, since that policy is Direct-
                 Routing-specific). Supports a PREVIEW (dry-run) mode and can
                 resume an interrupted run.

    3) ASSIGN  - Reads the most recent backup CSV for the CURRENTLY CONNECTED
                 tenant, finds which of those numbers are now sitting
                 unassigned in the tenant (i.e. have landed as Operator
                 Connect numbers after porting), matches them back to the
                 users/resource accounts who previously held them, and
                 reassigns them after an explicit confirmation. Also supports
                 PREVIEW mode and resuming an interrupted run.

  Typical Direct Routing -> Operator Connect migration flow using this script:
    a) Run option 1 to snapshot the current (Direct Routing) assignments.
    b) Run option 2 to unassign the old Direct Routing DIDs from users. This
       is what frees each number so the operator's side can pick it up.
    c) Submit/confirm the port order with your Operator Connect provider and
       wait for them to complete it - this is a manual, external step (can
       take days) that this script cannot perform for you. Once done, the
       numbers appear in the tenant as unassigned Operator Connect numbers.
    d) Run option 3 to reassign the (now Operator Connect) numbers to the
       same users, using the backup as the source of truth.

PREREQUISITES:
  - MicrosoftTeams PowerShell module (Set/Get/Remove-CsPhoneNumberAssignment
    require module version 3.0.0+; this script will offer to install the
    current version from PSGallery if it's missing).
  - The signed-in account must hold the Teams Administrator or Global
    Administrator role.
  - Windows PowerShell 5.1 is recommended for the interactive grid picker
    (Out-GridView). If Out-GridView isn't available, the script automatically
    falls back to a numbered list you can select from by typing e.g. "1,3,5-8".

REFERENCES (Microsoft Learn):
  - Get-CsPhoneNumberAssignment:
    https://learn.microsoft.com/powershell/module/microsoftteams/get-csphonenumberassignment
  - Set-CsPhoneNumberAssignment:
    https://learn.microsoft.com/powershell/module/microsoftteams/set-csphonenumberassignment
  - Remove-CsPhoneNumberAssignment:
    https://learn.microsoft.com/powershell/module/microsoftteams/remove-csphonenumberassignment
  - Get-CsOnlineApplicationInstance (resource accounts):
    https://learn.microsoft.com/powershell/module/microsoftteams/get-csonlineapplicationinstance
  - Assign, change, or remove a phone number for a user:
    https://learn.microsoft.com/microsoftteams/assign-change-or-remove-a-phone-number-for-a-user

SAFEGUARDS:
  - Every bulk action offers a PREVIEW (dry-run) mode and always requires an
    explicit typed confirmation ("REMOVE" / "ASSIGN") before making changes.
  - Every change is verified immediately afterwards against live tenant data
    (Teams number changes can be briefly async), and every run writes a
    timestamped .log, a "Plan" CSV (what was intended), and a "Results" CSV
    (what actually happened) to this folder, so every change is auditable.
  - If a bulk run is interrupted (crash, closed window, network drop), the
    next time you pick the same menu option you'll be offered the chance to
    resume exactly where it left off, retrying only the items that didn't
    already succeed.
  - Output filenames are tagged with the connected tenant's name, so backups
    and reports from different customer tenants stored in this same folder
    never get cross-matched by mistake.
  - Backup/Plan/Results CSVs write phone numbers in an Excel-safe format
    (="+614...") so opening and re-saving them in Excel can't silently strip
    the leading "+". The script un-wraps this automatically when it reads
    its own CSVs back in - you don't need to do anything.
================================================================================
#>

[CmdletBinding()]
param()

#region ---- Setup ----
$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
$LogFile    = Join-Path $ScriptRoot ("TeamsDIDToolkit_{0}.log" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$script:TenantTag = 'Tenant'   # replaced with the real tenant name once connected

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info', 'Warn', 'Error')][string]$Level = 'Info'
    )
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -Path $LogFile -Value $line
    switch ($Level) {
        'Warn'  { Write-Host $line -ForegroundColor Yellow }
        'Error' { Write-Host $line -ForegroundColor Red }
        default { Write-Host $line }
    }
}

function Normalize-PhoneNumber {
    param([string]$Number)
    if ([string]::IsNullOrWhiteSpace($Number)) { return $Number }
    $n = ($Number.Trim() -replace '[\s\-\(\)]', '')
    if ($n -notmatch '^\+') { $n = '+' + $n.TrimStart('+') }
    return $n
}

# Normalizes a GUID-like value (trim + lowercase) so lookups aren't broken by
# casing differences between cmdlets/module versions.
function Get-NormalizedGuid {
    param($Value)
    if ($null -eq $Value) { return $null }
    $s = $Value.ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    return $s.ToLowerInvariant()
}

# Teams PowerShell Module 3.0.0+ renamed Get-CsOnlineUser's "ObjectId" output
# property to "Identity" (still the raw object GUID, not the cmdlet's -Identity
# input parameter). Property names have shifted between module versions before
# and may again, so check both rather than hard-coding one.
function Get-EntityObjectId {
    param($Obj)
    $raw = $null
    if ($Obj.PSObject.Properties.Name -contains 'Identity' -and $Obj.Identity) { $raw = $Obj.Identity }
    elseif ($Obj.PSObject.Properties.Name -contains 'ObjectId' -and $Obj.ObjectId) { $raw = $Obj.ObjectId }
    return Get-NormalizedGuid $raw
}

# Excel silently strips the leading "+" from a phone number when a CSV is
# opened and re-saved. Wrapping the value as ="+614..." forces Excel to treat
# it as literal text. Unwrap-PhoneFromCsv reverses this when we read our own
# CSVs back in, so the wrapping is completely transparent to the rest of the script.
function Format-PhoneForCsv {
    param([string]$Number)
    if ([string]::IsNullOrWhiteSpace($Number)) { return $Number }
    return ('="{0}"' -f $Number)
}

function Unwrap-PhoneFromCsv {
    param([string]$Value)
    if ($null -eq $Value) { return $Value }
    if ($Value -match '^="(.*)"$') { return $Matches[1] }
    return $Value
}

function Export-CsvExcelSafe {
    param(
        [Parameter(Mandatory)][array]$Rows,
        [Parameter(Mandatory)][string]$Path,
        [string[]]$PhoneNumberProperties = @('TelephoneNumber'),
        [switch]$Append
    )
    $converted = foreach ($row in $Rows) {
        $clone = $row.PSObject.Copy()
        foreach ($p in $PhoneNumberProperties) {
            if (($clone.PSObject.Properties.Name -contains $p) -and $clone.$p) {
                $clone.$p = Format-PhoneForCsv $clone.$p
            }
        }
        $clone
    }
    if ($Append) {
        $converted | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8 -Append
    } else {
        $converted | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
    }
}

function New-OutputPath {
    param([Parameter(Mandatory)][string]$Suffix)
    return Join-Path $ScriptRoot ("{0}_{1}" -f $script:TenantTag, $Suffix)
}

# Out-GridView isn't guaranteed on every install (e.g. PowerShell 7 without the
# GraphicalTools module, or Server Core). Falls back to a numbered picker that
# accepts ranges like "1,3,5-8" so bulk selection still works without a GUI.
function Select-Rows {
    param(
        [Parameter(Mandatory)][array]$Rows,
        [Parameter(Mandatory)][string]$Title
    )
    if (Get-Command Out-GridView -ErrorAction SilentlyContinue) {
        return @($Rows | Out-GridView -Title $Title -PassThru)
    }

    Write-Host ""
    Write-Host "Out-GridView isn't available on this system - using numbered selection instead." -ForegroundColor Yellow
    Write-Host $Title -ForegroundColor Cyan
    for ($idx = 0; $idx -lt $Rows.Count; $idx++) {
        $r = $Rows[$idx]
        Write-Host ("[{0,4}] {1,-30} {2,-16} {3,-16} {4}" -f ($idx + 1), $r.UserPrincipalName, $r.TelephoneNumber, $r.NumberType, $r.Provider)
    }
    $sel = Read-Host "Enter row numbers to select (e.g. 1,3,5-8), or 'all'"
    if ($sel.Trim().ToLower() -eq 'all') { return $Rows }

    $indices = New-Object System.Collections.Generic.List[int]
    foreach ($part in ($sel -split ',')) {
        $part = $part.Trim()
        if ($part -match '^(\d+)-(\d+)$') {
            $indices.AddRange([int[]]($matches[1]..$matches[2]))
        } elseif ($part -match '^\d+$') {
            $indices.Add([int]$part)
        }
    }
    $indices = $indices | Sort-Object -Unique | Where-Object { $_ -ge 1 -and $_ -le $Rows.Count }
    return $indices | ForEach-Object { $Rows[$_ - 1] }
}
#endregion

#region ---- Module + connection ----
function Ensure-TeamsModule {
    $mod = Get-Module -ListAvailable -Name MicrosoftTeams | Sort-Object Version -Descending | Select-Object -First 1
    if (-not $mod) {
        Write-Log "MicrosoftTeams PowerShell module was not found." -Level Warn
        $answer = Read-Host "Install it now from PSGallery for the current user? (Y/N)"
        if ($answer -match '^(y|yes)$') {
            Install-Module -Name MicrosoftTeams -Scope CurrentUser -Force -AllowClobber -Repository PSGallery
        } else {
            throw "The MicrosoftTeams module is required to continue."
        }
    }
    Import-Module MicrosoftTeams -ErrorAction Stop
}

# Different MicrosoftTeams module versions/builds have used different parameter
# names for the same thing on these three cmdlets (TelephoneNumber vs PhoneNumber,
# NumberType vs PhoneNumberType) - documentation for the "current" names doesn't
# reliably match what's actually installed. Rather than guess, ask the loaded
# module directly via Get-Command and use whichever name it actually has.
function Resolve-PhoneCmdletParams {
    $map = @{}
    foreach ($cmdName in 'Get-CsPhoneNumberAssignment', 'Set-CsPhoneNumberAssignment', 'Remove-CsPhoneNumberAssignment') {
        $cmd = Get-Command $cmdName -ErrorAction Stop
        $pnParam = @('TelephoneNumber', 'PhoneNumber') | Where-Object { $cmd.Parameters.ContainsKey($_) } | Select-Object -First 1
        $ntParam = @('NumberType', 'PhoneNumberType') | Where-Object { $cmd.Parameters.ContainsKey($_) } | Select-Object -First 1
        if (-not $pnParam) {
            throw "Could not find a phone-number parameter (tried TelephoneNumber/PhoneNumber) on $cmdName. Available parameters: $($cmd.Parameters.Keys -join ', ')"
        }
        if ($cmdName -ne 'Get-CsPhoneNumberAssignment' -and -not $ntParam) {
            throw "Could not find a number-type parameter (tried NumberType/PhoneNumberType) on $cmdName. Available parameters: $($cmd.Parameters.Keys -join ', ')"
        }
        $map[$cmdName] = [PSCustomObject]@{ PhoneNumberParam = $pnParam; NumberTypeParam = $ntParam }
    }
    return $map
}

function Get-TenantTag {
    # Used to prefix every output filename so backups/reports from different
    # customer tenants never get cross-matched if they land in the same folder.
    $name = $null
    try {
        $tenant = Get-CsTenant -ErrorAction Stop
        $name = $tenant.DisplayName
        if ([string]::IsNullOrWhiteSpace($name) -and $tenant.Domains) { $name = @($tenant.Domains)[0] }
    } catch {
        $name = $null
    }
    if ([string]::IsNullOrWhiteSpace($name)) { $name = 'Tenant' }
    $safe = ($name -replace '[^a-zA-Z0-9]+', '-').Trim('-')
    if ([string]::IsNullOrWhiteSpace($safe)) { $safe = 'Tenant' }
    return $safe
}

function Connect-TeamsAdmin {
    Write-Host ""
    Write-Host "=== Microsoft Teams Admin Sign-In ===" -ForegroundColor Cyan
    Write-Host "A browser window will open for you to sign in with a Teams Administrator or Global Administrator account."
    Read-Host "Press Enter to continue"
    try {
        Connect-MicrosoftTeams -ErrorAction Stop | Out-Null
    } catch {
        Write-Log "Failed to connect to Microsoft Teams: $($_.Exception.Message)" -Level Error
        throw
    }
    try {
        $null = Get-CsOnlineUser -ResultSize 1 -ErrorAction Stop
        Write-Log "Connected to the Teams admin center."
    } catch {
        Write-Log "Connected, but a quick permissions check failed - the account may lack Teams Administrator rights: $($_.Exception.Message)" -Level Warn
    }
    $script:TenantTag = Get-TenantTag
    Write-Log "Connected to tenant '$($script:TenantTag)' - output files will be prefixed with this tag."
}
#endregion

#region ---- Shared data helpers ----
function Get-AllPhoneAssignments {
    # Get-CsPhoneNumberAssignment caps at 500 results by default (1000 max per
    # call) - page with -Skip/-Top so this works for tenants with hundreds of DIDs.
    param(
        [string]$NumberType,
        [string]$PstnAssignmentStatus
    )
    $results = [System.Collections.Generic.List[object]]::new()
    $top = 1000
    $skip = 0
    while ($true) {
        $params = @{ Top = $top; Skip = $skip; ErrorAction = 'Stop' }
        $ntParam = $script:PhoneCmdletParams['Get-CsPhoneNumberAssignment'].NumberTypeParam
        if ($NumberType -and $ntParam) { $params[$ntParam] = $NumberType }
        if ($PstnAssignmentStatus) { $params.PstnAssignmentStatus = $PstnAssignmentStatus }
        $batch = @(Get-CsPhoneNumberAssignment @params)
        if ($batch.Count -eq 0) { break }
        $results.AddRange($batch)
        if ($batch.Count -lt $top) { break }
        $skip += $top
    }
    return $results
}

function Get-ResourceAccountTypeName {
    param([string]$ApplicationId)
    switch ($ApplicationId) {
        'ce933385-9390-45d1-9512-c8d228074e07' { return 'Auto Attendant' }
        '11cd3e2e-fccb-42ad-ad00-878b93575e07' { return 'Call Queue' }
        default { return 'Resource Account' }
    }
}

function Get-IdentityLookup {
    # Get-CsPhoneNumberAssignment returns AssignedPstnTargetId as an object GUID,
    # not a UPN - build a lookup covering BOTH regular users and resource
    # accounts (auto attendants / call queues, which also hold DIDs and are
    # commonly part of a DR -> OC migration) so reports/matching show real names.
    Write-Log "Retrieving tenant users for phone-number-to-identity resolution..."
    $lookup = @{}
    $usersSeen = 0
    $usersKeyed = 0
    # No -ResultSize: omitting it returns every user by default, and its type
    # varies between module versions (Unlimited enum vs plain uint32), so a
    # literal value here is more likely to break than help.
    Get-CsOnlineUser -ErrorAction Stop | ForEach-Object {
        $usersSeen++
        $key = Get-EntityObjectId $_
        if ($key) {
            $usersKeyed++
            $lookup[$key] = [PSCustomObject]@{
                UserPrincipalName        = $_.UserPrincipalName
                DisplayName              = $_.DisplayName
                AccountType              = 'User'
                UsageLocation            = $_.UsageLocation
                EnterpriseVoiceEnabled   = $_.EnterpriseVoiceEnabled
                OnlineVoiceRoutingPolicy = $_.OnlineVoiceRoutingPolicy
                TeamsCallingPolicy       = $_.TeamsCallingPolicy
                TenantDialPlan           = $_.TenantDialPlan
                TeamsIPPhonePolicy       = $_.TeamsIPPhonePolicy
                AccountEnabled           = $_.AccountEnabled
            }
        }
    }
    Write-Log "Retrieved $usersSeen user(s), resolved an identity key for $usersKeyed of them."
    if ($usersSeen -gt 0 -and $usersKeyed -eq 0) {
        Write-Log "None of the retrieved users had a usable Identity/ObjectId property - phone-number-to-user matching will fail. This usually means the MicrosoftTeams module version on this machine exposes that property under a different name than expected; run 'Get-CsOnlineUser -ResultSize 1 | Get-Member' and let the script maintainer know what identity-like property names are present." -Level Error
    }

    Write-Log "Retrieving resource accounts (auto attendants / call queues)..."
    try {
        Get-CsOnlineApplicationInstance -ErrorAction Stop | ForEach-Object {
            $key = Get-EntityObjectId $_
            if ($key -and -not $lookup.ContainsKey($key)) {
                $lookup[$key] = [PSCustomObject]@{
                    UserPrincipalName        = $_.UserPrincipalName
                    DisplayName              = $_.DisplayName
                    AccountType              = Get-ResourceAccountTypeName $_.ApplicationId
                    UsageLocation            = ''
                    EnterpriseVoiceEnabled   = ''
                    OnlineVoiceRoutingPolicy = ''
                    TeamsCallingPolicy       = ''
                    TenantDialPlan           = ''
                    TeamsIPPhonePolicy       = ''
                    AccountEnabled           = ''
                }
            }
        }
    } catch {
        Write-Log "Could not retrieve resource accounts: $($_.Exception.Message)" -Level Warn
    }
    return $lookup
}

# Polls live tenant state briefly after a change, since Teams number
# operations can take a few seconds to fully propagate even after the cmdlet
# returns success. Keeps bulk runs fast: most numbers verify on the first check.
function Wait-ForNumberState {
    param(
        [Parameter(Mandatory)][string]$TelephoneNumber,
        [Parameter(Mandatory)][scriptblock]$SuccessTest,
        [int]$MaxAttempts = 3,
        [int]$DelaySeconds = 2
    )
    $getParams = @{ ErrorAction = 'SilentlyContinue' }
    $getParams[$script:PhoneCmdletParams['Get-CsPhoneNumberAssignment'].PhoneNumberParam] = $TelephoneNumber
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        Start-Sleep -Seconds $DelaySeconds
        $live = @(Get-CsPhoneNumberAssignment @getParams)
        if ($live.Count -gt 0 -and (& $SuccessTest $live[0])) { return $true }
    }
    return $false
}

# Looks for a Plan file (the full intended target list) from a prior run of
# the same option that doesn't have a matching completed Results row for
# every planned number, so an interrupted bulk run can be resumed instead of
# restarted from scratch. Rows logged as 'DryRun' or 'Failed' are treated as
# still pending, so a resume also retries failures and turns a prior preview into a real run.
function Get-ResumableRun {
    param([Parameter(Mandatory)][string]$Prefix)
    $planFile = Get-ChildItem -Path $ScriptRoot -Filter "$($script:TenantTag)_${Prefix}Plan_*.csv" -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $planFile) { return $null }
    if ($planFile.Name -notmatch '(?<runid>\d{8}_\d{6})\.csv$') { return $null }
    $runId = $Matches.runid
    $resultsPath = New-OutputPath "${Prefix}Results_$runId.csv"
    $planRows = @(Import-Csv $planFile.FullName)
    if ($planRows.Count -eq 0) { return $null }

    $doneNumbers = @()
    if (Test-Path $resultsPath) {
        $doneNumbers = @(Import-Csv $resultsPath | Where-Object { $_.Status -eq 'Success' } |
            ForEach-Object { Normalize-PhoneNumber (Unwrap-PhoneFromCsv $_.TelephoneNumber) })
    }

    $pending = @($planRows | Where-Object {
        (Normalize-PhoneNumber (Unwrap-PhoneFromCsv $_.TelephoneNumber)) -notin $doneNumbers
    })
    if ($pending.Count -eq 0) { return $null }

    [PSCustomObject]@{
        RunId        = $runId
        ResultsPath  = $resultsPath
        Pending      = $pending
        TotalPlanned = $planRows.Count
    }
}
#endregion

#region ---- Option 1: Backup ----
function Invoke-BackupNumbers {
    Write-Host ""
    Write-Host "=== Backing up phone number assignments ===" -ForegroundColor Cyan

    $identityLookup = Get-IdentityLookup

    # No -PstnAssignmentStatus filter here on purpose: this is meant to be a
    # full inventory snapshot, so it must include every status - Unassigned,
    # UserAssigned, ConferenceAssigned, VoiceApplicationAssigned (resource
    # accounts), ThirdPartyAppAssigned and PolicyAssigned - not just numbers
    # assigned directly to a user.
    Write-Log "Retrieving every phone number in the tenant, any assignment status (this can take a while for large tenants)..."
    $allNumbers = Get-AllPhoneAssignments
    Write-Log "$($allNumbers.Count) number(s) found across all assignment statuses."
    $statusCounts = $allNumbers | Group-Object PstnAssignmentStatus | ForEach-Object { "$($_.Name)=$($_.Count)" }
    Write-Log ("Breakdown by status: " + ($statusCounts -join ', '))

    $exportTime = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $unresolvedAssigned = 0
    $rows = foreach ($a in $allNumbers) {
        $u = $null
        if ($a.AssignedPstnTargetId) {
            $key = Get-NormalizedGuid $a.AssignedPstnTargetId
            if ($key) { $u = $identityLookup[$key] }
            if (-not $u -and $a.PstnAssignmentStatus -ne 'Unassigned') { $unresolvedAssigned++ }
        }
        [PSCustomObject]@{
            ExportDateTime           = $exportTime
            UserPrincipalName        = $u.UserPrincipalName
            DisplayName              = $u.DisplayName
            AccountType              = $u.AccountType
            TelephoneNumber          = $a.TelephoneNumber
            NumberType               = $a.NumberType
            PstnAssignmentStatus     = $a.PstnAssignmentStatus
            Provider                 = $a.PstnPartnerName
            OperatorId               = $a.OperatorId
            ActivationState          = $a.ActivationState
            AssignmentCategory       = $a.AssignmentCategory
            City                     = $a.City
            IsoCountryCode           = $a.IsoCountryCode
            LocationId               = $a.LocationId
            NumberSource             = $a.NumberSource
            AssignedObjectId         = $a.AssignedPstnTargetId
            UsageLocation            = $u.UsageLocation
            EnterpriseVoiceEnabled   = $u.EnterpriseVoiceEnabled
            OnlineVoiceRoutingPolicy = $u.OnlineVoiceRoutingPolicy
            TeamsCallingPolicy       = $u.TeamsCallingPolicy
            TenantDialPlan           = $u.TenantDialPlan
            TeamsIPPhonePolicy       = $u.TeamsIPPhonePolicy
            AccountEnabled           = $u.AccountEnabled
        }
    }

    if ($unresolvedAssigned -gt 0) {
        Write-Log "$unresolvedAssigned assigned number(s) could not be matched to a user or resource account (AssignedPstnTargetId had no match in the lookup - it may be a Teams shared calling routing policy instance rather than a user/resource account)." -Level Warn
    }

    if (-not $rows) {
        Write-Log "No phone numbers were found in the tenant - nothing to back up." -Level Warn
        return
    }

    $outPath = New-OutputPath ("TeamsPhoneNumberBackup_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    Export-CsvExcelSafe -Rows $rows -Path $outPath -PhoneNumberProperties @('TelephoneNumber')

    Write-Log "Backup complete: $($rows.Count) numbers exported to $outPath"
    Write-Host "Saved: $outPath" -ForegroundColor Green
}
#endregion

#region ---- Option 2: Remove DIDs ----
function Resolve-TargetForNumber {
    param([string]$Number, [hashtable]$IdentityLookup)
    $getParams = @{ ErrorAction = 'SilentlyContinue' }
    $getParams[$script:PhoneCmdletParams['Get-CsPhoneNumberAssignment'].PhoneNumberParam] = $Number
    $live = @(Get-CsPhoneNumberAssignment @getParams)
    if ($live.Count -eq 0) { return $null }
    $live = $live[0]
    if ($live.PstnAssignmentStatus -eq 'Unassigned') { return $null }

    $identity = $Number
    $key = Get-NormalizedGuid $live.AssignedPstnTargetId
    if ($key -and $IdentityLookup.ContainsKey($key) -and $IdentityLookup[$key].UserPrincipalName) {
        $identity = $IdentityLookup[$key].UserPrincipalName
    } elseif ($live.AssignedPstnTargetId) {
        $identity = $live.AssignedPstnTargetId
    }
    [PSCustomObject]@{
        Identity        = $identity
        TelephoneNumber = $live.TelephoneNumber
        NumberType      = $live.NumberType
    }
}

function Invoke-RemoveDids {
    Write-Host ""
    Write-Host "=== Remove DIDs ===" -ForegroundColor Cyan

    $targets = [System.Collections.Generic.List[object]]::new()
    $runId = $null
    $reportPath = $null
    $resumed = $false

    $resume = Get-ResumableRun -Prefix 'RemoveDids'
    if ($resume) {
        Write-Host "Found an incomplete removal run (ID $($resume.RunId)) with $($resume.Pending.Count) of $($resume.TotalPlanned) DID(s) not yet completed." -ForegroundColor Yellow
        $rAns = Read-Host "Resume that run instead of starting a new selection? (Y/N)"
        if ($rAns -match '^(y|yes)$') {
            foreach ($p in $resume.Pending) {
                $targets.Add([PSCustomObject]@{
                    Identity        = $p.Identity
                    TelephoneNumber = Unwrap-PhoneFromCsv $p.TelephoneNumber
                    NumberType      = $p.NumberType
                })
            }
            $runId      = $resume.RunId
            $reportPath = $resume.ResultsPath
            $resumed    = $true
            Write-Log "Resuming removal run $runId with $($targets.Count) pending DID(s)."
        }
    }

    if (-not $resumed) {
        Write-Host "Choose how to select the DIDs to remove:"
        Write-Host "  1) Type/paste phone numbers (comma or newline separated) - best for 1-10 DIDs"
        Write-Host "  2) Import a CSV file with a 'TelephoneNumber' column - best for large bulk removals"
        Write-Host "  3) Interactive picker - browse all currently assigned DIDs and multi-select"
        $choice = Read-Host "Selection (1/2/3)"

        $identityLookup = Get-IdentityLookup

        switch ($choice) {
            '1' {
                $raw = Read-Host "Enter phone number(s), separated by commas"
                $numbers = $raw -split '[,;\r\n]' | Where-Object { $_.Trim() -ne '' } | ForEach-Object { Normalize-PhoneNumber $_ }
                foreach ($n in $numbers) {
                    $t = Resolve-TargetForNumber -Number $n -IdentityLookup $identityLookup
                    if ($t) { $targets.Add($t) } else { Write-Log "Skipped $n - not found or not currently assigned." -Level Warn }
                }
            }
            '2' {
                $csvPath = Read-Host "Full path to the CSV file (must contain a 'TelephoneNumber' column)"
                if (-not (Test-Path $csvPath)) { Write-Log "File not found: $csvPath" -Level Error; return }
                foreach ($row in (Import-Csv -Path $csvPath)) {
                    $n = Normalize-PhoneNumber (Unwrap-PhoneFromCsv $row.TelephoneNumber)
                    $t = Resolve-TargetForNumber -Number $n -IdentityLookup $identityLookup
                    if ($t) { $targets.Add($t) } else { Write-Log "Skipped $n - not found or not currently assigned." -Level Warn }
                }
            }
            '3' {
                # No -PstnAssignmentStatus filter: query everything, then exclude only
                # Unassigned, so resource-account (VoiceApplicationAssigned), conference,
                # and policy-assigned numbers show up here too, not just plain user assignments.
                Write-Log "Retrieving all currently assigned DIDs for the picker (this can take a while for large tenants)..."
                $assigned = @(Get-AllPhoneAssignments) | Where-Object { $_.PstnAssignmentStatus -ne 'Unassigned' }
                $picker = foreach ($a in $assigned) {
                    $u = $null
                    $key = Get-NormalizedGuid $a.AssignedPstnTargetId
                    if ($key) { $u = $identityLookup[$key] }
                    $displayIdentity = $u.UserPrincipalName
                    if (-not $displayIdentity) { $displayIdentity = $a.AssignedPstnTargetId }
                    [PSCustomObject]@{
                        UserPrincipalName = $displayIdentity
                        DisplayName       = $u.DisplayName
                        TelephoneNumber   = $a.TelephoneNumber
                        NumberType        = $a.NumberType
                        Provider          = $a.PstnPartnerName
                    }
                }
                $selected = Select-Rows -Rows (@($picker) | Sort-Object UserPrincipalName) -Title "Select DID(s) to REMOVE - Ctrl/Shift+Click for multiple, then OK"
                if (-not $selected) { Write-Log "No rows selected - nothing to remove." -Level Warn; return }
                foreach ($row in $selected) {
                    $targets.Add([PSCustomObject]@{ Identity = $row.UserPrincipalName; TelephoneNumber = $row.TelephoneNumber; NumberType = $row.NumberType })
                }
            }
            default { Write-Log "Invalid selection." -Level Warn; return }
        }

        if ($targets.Count -eq 0) { Write-Log "No valid DIDs resolved - nothing to remove." -Level Warn; return }

        $runId      = Get-Date -Format 'yyyyMMdd_HHmmss'
        $planPath   = New-OutputPath "RemoveDidsPlan_$runId.csv"
        $reportPath = New-OutputPath "RemoveDidsResults_$runId.csv"
        Export-CsvExcelSafe -Rows ($targets | Select-Object Identity, TelephoneNumber, NumberType) -Path $planPath -PhoneNumberProperties @('TelephoneNumber')
    }

    Write-Host ""
    Write-Host "The following $($targets.Count) DID(s) will be UNASSIGNED:" -ForegroundColor Yellow
    $targets | Format-Table Identity, TelephoneNumber, NumberType -AutoSize | Out-String | Write-Host

    $modeAnswer = Read-Host "Type REMOVE to apply these changes, PREVIEW to dry-run without changing anything, or press Enter to cancel"
    $mode = $modeAnswer.Trim().ToUpper()
    if ($mode -ne 'REMOVE' -and $mode -ne 'PREVIEW') { Write-Log "Removal cancelled by operator." -Level Warn; return }
    $dryRun = ($mode -eq 'PREVIEW')

    $notify = $false
    $clearPolicy = $false
    if (-not $dryRun) {
        $notifyAnswer = Read-Host "Send Teams' built-in email notification to each affected user about the removal? (Y/N)"
        $notify = $notifyAnswer -match '^(y|yes)$'

        # Microsoft's documented "Move numbers from Direct Routing to Operator Connect"
        # process (learn.microsoft.com/microsoftteams/operator-connect-configure) has a
        # Step 2 after unassigning the number: clear the user's Online Voice Routing
        # Policy, since it points at the old Direct Routing SBC/trunk and has no purpose
        # once that number is gone. Only applies to DirectRouting numbers assigned to users.
        $clearPolicyAnswer = Read-Host "Also clear each user's Online Voice Routing Policy for removed Direct Routing numbers, as Microsoft recommends for DR -> Operator Connect migrations? (Y/N)"
        $clearPolicy = $clearPolicyAnswer -match '^(y|yes)$'
    }

    $i = 0
    foreach ($t in $targets) {
        $i++
        Write-Progress -Activity "Removing DIDs" -Status "$i of $($targets.Count): $($t.TelephoneNumber)" -PercentComplete (($i / $targets.Count) * 100)

        if ($dryRun) {
            Write-Log "[PREVIEW] Would remove $($t.TelephoneNumber) from $($t.Identity)"
            $row = [PSCustomObject]@{ Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); Identity = $t.Identity; TelephoneNumber = $t.TelephoneNumber; NumberType = $t.NumberType; Status = 'DryRun'; Verified = 'Skipped'; VoiceRoutingPolicyCleared = 'Skipped'; Error = '' }
        } else {
            try {
                $rmMap = $script:PhoneCmdletParams['Remove-CsPhoneNumberAssignment']
                $rmParams = @{ Identity = $t.Identity; ErrorAction = 'Stop' }
                $rmParams[$rmMap.PhoneNumberParam] = $t.TelephoneNumber
                $rmParams[$rmMap.NumberTypeParam] = $t.NumberType
                if ($notify) { $rmParams.Notify = $true }
                Remove-CsPhoneNumberAssignment @rmParams

                $verified = Wait-ForNumberState -TelephoneNumber $t.TelephoneNumber -SuccessTest { param($x) $x.PstnAssignmentStatus -eq 'Unassigned' }
                $verifiedText = 'No'
                if ($verified) { $verifiedText = 'Yes' }
                if ($verified) {
                    Write-Log "Removed $($t.TelephoneNumber) from $($t.Identity)"
                } else {
                    Write-Log "Removed $($t.TelephoneNumber) from $($t.Identity), but could not verify the change afterwards - check manually." -Level Warn
                }

                $policyCleared = 'N/A'
                if ($clearPolicy -and $t.NumberType -eq 'DirectRouting') {
                    try {
                        Grant-CsOnlineVoiceRoutingPolicy -Identity $t.Identity -PolicyName $null -ErrorAction Stop
                        $policyCleared = 'Yes'
                        Write-Log "Cleared OnlineVoiceRoutingPolicy for $($t.Identity)"
                    } catch {
                        $policyCleared = 'Failed'
                        Write-Log "Could not clear OnlineVoiceRoutingPolicy for $($t.Identity) (expected for resource accounts): $($_.Exception.Message)" -Level Warn
                    }
                }

                $row = [PSCustomObject]@{ Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); Identity = $t.Identity; TelephoneNumber = $t.TelephoneNumber; NumberType = $t.NumberType; Status = 'Success'; Verified = $verifiedText; VoiceRoutingPolicyCleared = $policyCleared; Error = '' }
            } catch {
                Write-Log "FAILED to remove $($t.TelephoneNumber) from $($t.Identity): $($_.Exception.Message)" -Level Error
                $row = [PSCustomObject]@{ Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); Identity = $t.Identity; TelephoneNumber = $t.TelephoneNumber; NumberType = $t.NumberType; Status = 'Failed'; Verified = 'No'; VoiceRoutingPolicyCleared = 'N/A'; Error = $_.Exception.Message }
            }
        }

        Export-CsvExcelSafe -Rows @($row) -Path $reportPath -PhoneNumberProperties @('TelephoneNumber') -Append
        if (-not $dryRun) { Start-Sleep -Milliseconds 250 }   # throttling headroom for large bulk runs
    }
    Write-Progress -Activity "Removing DIDs" -Completed

    $allResults = @(Import-Csv $reportPath)
    $ok      = @($allResults | Where-Object Status -eq 'Success').Count
    $fail    = @($allResults | Where-Object Status -eq 'Failed').Count
    $preview = @($allResults | Where-Object Status -eq 'DryRun').Count
    Write-Host "Done: $ok succeeded, $fail failed, $preview previewed. Full report: $reportPath" -ForegroundColor Green

    if (-not $dryRun -and $ok -gt 0) {
        Write-Host ""
        Write-Host "NEXT STEPS - these $ok number(s) are now unassigned, which is what your Operator Connect" -ForegroundColor Cyan
        Write-Host "provider needs before they can add/port them into this tenant:" -ForegroundColor Cyan
        Write-Host "  1. Confirm/submit the port order for these DIDs with your Operator Connect provider." -ForegroundColor Cyan
        Write-Host "  2. Wait for the provider to complete the port. Removal itself can take up to 10 minutes" -ForegroundColor Cyan
        Write-Host "     to propagate (rarely, up to 24 hours) - the porting time on top of that is entirely" -ForegroundColor Cyan
        Write-Host "     up to the provider, and can be days." -ForegroundColor Cyan
        Write-Host "  3. Confirm in Teams admin center under Voice > Phone numbers that each number now shows" -ForegroundColor Cyan
        Write-Host "     Number type = Operator Connect before running option 3 - running option 3 too early" -ForegroundColor Cyan
        Write-Host "     will simply find no match yet for that number." -ForegroundColor Cyan
    }
}
#endregion

#region ---- Option 3: Assign DIDs ----
function Invoke-AssignDids {
    Write-Host ""
    Write-Host "=== Assign DIDs (from latest backup) ===" -ForegroundColor Cyan

    $targets = [System.Collections.Generic.List[object]]::new()
    $runId = $null
    $reportPath = $null
    $resumed = $false

    $resume = Get-ResumableRun -Prefix 'AssignDids'
    if ($resume) {
        Write-Host "Found an incomplete assignment run (ID $($resume.RunId)) with $($resume.Pending.Count) of $($resume.TotalPlanned) DID(s) not yet completed." -ForegroundColor Yellow
        $rAns = Read-Host "Resume that run instead of starting a new selection? (Y/N)"
        if ($rAns -match '^(y|yes)$') {
            foreach ($p in $resume.Pending) {
                $targets.Add([PSCustomObject]@{
                    UserPrincipalName = $p.UserPrincipalName
                    TelephoneNumber   = Unwrap-PhoneFromCsv $p.TelephoneNumber
                    NumberType        = $p.NumberType
                    LocationId        = $p.LocationId
                })
            }
            $runId      = $resume.RunId
            $reportPath = $resume.ResultsPath
            $resumed    = $true
            Write-Log "Resuming assignment run $runId with $($targets.Count) pending DID(s)."
        }
    }

    if (-not $resumed) {
        $latest = Get-ChildItem -Path $ScriptRoot -Filter "$($script:TenantTag)_TeamsPhoneNumberBackup_*.csv" -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $latest) {
            Write-Log "No backup file found for tenant '$($script:TenantTag)' in $ScriptRoot. Run option 1 first while connected to this tenant." -Level Error
            return
        }
        Write-Log "Using latest backup: $($latest.Name) (last written $($latest.LastWriteTime))"
        $backupRows = Import-Csv -Path $latest.FullName

        $targetTypeInput = Read-Host "Number type to look for in the tenant's unassigned inventory [OperatorConnect]"
        $targetType = if ([string]::IsNullOrWhiteSpace($targetTypeInput)) { 'OperatorConnect' } else { $targetTypeInput }

        Write-Log "Retrieving unassigned $targetType numbers currently in the tenant (this can take a while)..."
        $unassigned = Get-AllPhoneAssignments -PstnAssignmentStatus 'Unassigned' -NumberType $targetType
        if ($unassigned.Count -eq 0) {
            Write-Log "No unassigned $targetType numbers found in the tenant. Port/activate the numbers with the carrier first." -Level Warn
            return
        }
        $unassignedByNumber = @{}
        foreach ($u in $unassigned) { $unassignedByNumber[(Normalize-PhoneNumber $u.TelephoneNumber)] = $u }

        $matches = foreach ($row in $backupRows) {
            $num = Normalize-PhoneNumber (Unwrap-PhoneFromCsv $row.TelephoneNumber)
            if ($unassignedByNumber.ContainsKey($num) -and -not [string]::IsNullOrWhiteSpace($row.UserPrincipalName)) {
                $live = $unassignedByNumber[$num]
                [PSCustomObject]@{
                    UserPrincipalName  = $row.UserPrincipalName
                    DisplayName        = $row.DisplayName
                    AccountType        = $row.AccountType
                    TelephoneNumber    = $live.TelephoneNumber
                    NumberType         = $live.NumberType
                    Provider           = $live.PstnPartnerName
                    PreviousNumberType = $row.NumberType
                    LocationId         = $row.LocationId
                }
            }
        }

        if (-not $matches) {
            Write-Log "No backed-up numbers matched an unassigned $targetType number in the tenant." -Level Warn
            return
        }

        Write-Host ""
        Write-Host "Found $(@($matches).Count) DID(s) from the backup that are now unassigned $targetType numbers:" -ForegroundColor Yellow
        $selected = Select-Rows -Rows (@($matches) | Sort-Object UserPrincipalName) -Title "Select DID(s) to ASSIGN back to their user - Ctrl/Shift+Click for multiple, then OK"
        if (-not $selected) { Write-Log "No rows selected - nothing to assign." -Level Warn; return }

        foreach ($row in @($selected)) {
            $targets.Add([PSCustomObject]@{
                UserPrincipalName = $row.UserPrincipalName
                TelephoneNumber   = $row.TelephoneNumber
                NumberType        = $row.NumberType
                LocationId        = $row.LocationId
            })
        }

        $runId      = Get-Date -Format 'yyyyMMdd_HHmmss'
        $planPath   = New-OutputPath "AssignDidsPlan_$runId.csv"
        $reportPath = New-OutputPath "AssignDidsResults_$runId.csv"
        Export-CsvExcelSafe -Rows ($targets | Select-Object UserPrincipalName, TelephoneNumber, NumberType, LocationId) -Path $planPath -PhoneNumberProperties @('TelephoneNumber')
    }

    Write-Host ""
    Write-Host "The following $($targets.Count) DID(s) will be ASSIGNED:" -ForegroundColor Yellow
    $targets | Format-Table UserPrincipalName, TelephoneNumber, NumberType -AutoSize | Out-String | Write-Host

    $modeAnswer = Read-Host "Type ASSIGN to apply these changes, PREVIEW to dry-run without changing anything, or press Enter to cancel"
    $mode = $modeAnswer.Trim().ToUpper()
    if ($mode -ne 'ASSIGN' -and $mode -ne 'PREVIEW') { Write-Log "Assignment cancelled by operator." -Level Warn; return }
    $dryRun = ($mode -eq 'PREVIEW')

    $notify = $false
    if (-not $dryRun) {
        $notifyAnswer = Read-Host "Send Teams' built-in email notification to each affected user about the new number? (Y/N)"
        $notify = $notifyAnswer -match '^(y|yes)$'
    }

    $i = 0
    foreach ($m in $targets) {
        $i++
        Write-Progress -Activity "Assigning DIDs" -Status "$i of $($targets.Count): $($m.TelephoneNumber) -> $($m.UserPrincipalName)" -PercentComplete (($i / $targets.Count) * 100)

        if ($dryRun) {
            Write-Log "[PREVIEW] Would assign $($m.TelephoneNumber) to $($m.UserPrincipalName)"
            $row = [PSCustomObject]@{ Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); UserPrincipalName = $m.UserPrincipalName; TelephoneNumber = $m.TelephoneNumber; NumberType = $m.NumberType; Status = 'DryRun'; Verified = 'Skipped'; Error = '' }
        } else {
            try {
                $setMap = $script:PhoneCmdletParams['Set-CsPhoneNumberAssignment']
                $params = @{ Identity = $m.UserPrincipalName; ErrorAction = 'Stop' }
                $params[$setMap.PhoneNumberParam] = $m.TelephoneNumber
                $params[$setMap.NumberTypeParam] = $m.NumberType
                if (-not [string]::IsNullOrWhiteSpace($m.LocationId) -and $m.LocationId -ne '00000000-0000-0000-0000-000000000000') {
                    $params.LocationId = $m.LocationId
                }
                if ($notify) { $params.Notify = $true }
                Set-CsPhoneNumberAssignment @params

                $verified = Wait-ForNumberState -TelephoneNumber $m.TelephoneNumber -SuccessTest { param($x) $x.PstnAssignmentStatus -ne 'Unassigned' }
                $verifiedText = 'No'
                if ($verified) { $verifiedText = 'Yes' }
                if ($verified) {
                    Write-Log "Assigned $($m.TelephoneNumber) to $($m.UserPrincipalName)"
                } else {
                    Write-Log "Assigned $($m.TelephoneNumber) to $($m.UserPrincipalName), but could not verify the change afterwards - check manually." -Level Warn
                }
                $row = [PSCustomObject]@{ Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); UserPrincipalName = $m.UserPrincipalName; TelephoneNumber = $m.TelephoneNumber; NumberType = $m.NumberType; Status = 'Success'; Verified = $verifiedText; Error = '' }
            } catch {
                Write-Log "FAILED to assign $($m.TelephoneNumber) to $($m.UserPrincipalName): $($_.Exception.Message)" -Level Error
                $row = [PSCustomObject]@{ Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'); UserPrincipalName = $m.UserPrincipalName; TelephoneNumber = $m.TelephoneNumber; NumberType = $m.NumberType; Status = 'Failed'; Verified = 'No'; Error = $_.Exception.Message }
            }
        }

        Export-CsvExcelSafe -Rows @($row) -Path $reportPath -PhoneNumberProperties @('TelephoneNumber') -Append
        if (-not $dryRun) { Start-Sleep -Milliseconds 250 }
    }
    Write-Progress -Activity "Assigning DIDs" -Completed

    $allResults = @(Import-Csv $reportPath)
    $ok      = @($allResults | Where-Object Status -eq 'Success').Count
    $fail    = @($allResults | Where-Object Status -eq 'Failed').Count
    $preview = @($allResults | Where-Object Status -eq 'DryRun').Count
    Write-Host "Done: $ok succeeded, $fail failed, $preview previewed. Full report: $reportPath" -ForegroundColor Green
}
#endregion

#region ---- Main menu ----
function Show-MainMenu {
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host " Microsoft Teams Phone Number (DID) Migration Toolkit" -ForegroundColor Cyan
    Write-Host " Tenant: $($script:TenantTag)  |  Direct Routing -> Operator Connect" -ForegroundColor Cyan
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host " 1) Backup all users' phone number assignments to CSV"
    Write-Host " 2) Remove DIDs (unassign from users)"
    Write-Host " 3) Assign DIDs (reassign from latest backup)"
    Write-Host " 0) Disconnect and exit"
    Write-Host "================================================================"
    Read-Host "Select an option"
}

try {
    Ensure-TeamsModule
    $script:PhoneCmdletParams = Resolve-PhoneCmdletParams
    foreach ($cmdName in $script:PhoneCmdletParams.Keys) {
        $p = $script:PhoneCmdletParams[$cmdName]
        Write-Log "$cmdName parameter names on this module: PhoneNumber=[$($p.PhoneNumberParam)] NumberType=[$($p.NumberTypeParam)]"
    }
    Connect-TeamsAdmin

    $exit = $false
    do {
        $choice = Show-MainMenu
        switch ($choice) {
            '1'     { Invoke-BackupNumbers }
            '2'     { Invoke-RemoveDids }
            '3'     { Invoke-AssignDids }
            '0'     { $exit = $true }
            default { Write-Host "Invalid selection." -ForegroundColor Yellow }
        }
    } while (-not $exit)
}
finally {
    try { Disconnect-MicrosoftTeams -ErrorAction SilentlyContinue | Out-Null } catch {}
    Write-Log "Session ended."
}
#endregion
