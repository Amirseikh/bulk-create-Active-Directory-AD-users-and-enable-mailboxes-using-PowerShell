<#
.SYNOPSIS
    AD User Provisioning Tool - GUI application for bulk-creating Active Directory
    users from an Excel sheet, with optional Exchange mailbox enablement.

.DESCRIPTION
    Configuration (saved to JSON next to the script):
      - Forest:       Forest Name, Global Catalog Server, Forest Domain DN
      - Provisioning: Username + password of the account used to create users
                      (password stored DPAPI-encrypted, current user+machine only)
      - Domain:       Domain DNs (multiple), Domain Controllers/Servers (multiple)
      - Exchange:     Exchange Forest, Exchange Hosts (multiple), Exchange User,
                      Mailbox Databases (multiple, each with max mailbox count)
                      + distribution mode (fill sequentially / round-robin)
      - Search:       Search DN (container), Iterated Search Filters (multiple)

    Excel column mapping (per approved mapping sheet):
      AD Attribute      <- Excel Column(s)
      employeeID        <- employeeID
      employeeNumber    <- employeeNumber
      givenName         <- First name
      sn                <- Second name + Third name + Last name
      cn                <- First name + Second name + Third name + Last name
      displayName       <- First name + Second name + Third name + Last name
      name              <- First name + Second name + Third name + Last name
      mobile            <- mobile
      sAMAccountName    <- sAMAccountName
      mail              <- mail
      title             <- title
      occupation        <- occupation        (schema-extended attribute)
      cluster           <- cluster           (schema-extended attribute)
      clusterCode       <- clusterCode       (schema-extended attribute)
      facilty           <- facilty           (schema-extended attribute)
      faciltyCode       <- faciltyCode       (schema-extended attribute)
      employeeType      <- employeeType

    Rules:
      - Random 12-character passwords (upper + lower + digit + special)
      - "User must change password at next logon" is set
      - Username & email are taken exactly as provided in the sheet
      - Default target OU: OU=Users,OU=YourOrg,DC=contoso,DC=com

.NOTES
    Requires: RSAT ActiveDirectory module. Run as a user with rights to
    create objects in the target OU. For Exchange enablement the account
    supplied under "Exchange User" needs Recipient Management rights.
#>

#Requires -Version 5.1

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

$script:ScriptDir  = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:ConfigPath = Join-Path $script:ScriptDir 'ADProvisioning.config.json'
$script:Results    = New-Object System.Collections.ObjectModel.ObservableCollection[object]
$script:Preview    = @()
$script:ADPassEnc  = $null   # DPAPI-encrypted provisioning password (from config or PasswordBox)

# ------------------------------------------------------------------ #
#  Helper functions                                                  #
# ------------------------------------------------------------------ #

function New-RandomPassword {
    param([int]$Length = 12)
    $upper   = 'ABCDEFGHJKLMNPQRSTUVWXYZ'.ToCharArray()
    $lower   = 'abcdefghijkmnpqrstuvwxyz'.ToCharArray()
    $digits  = '23456789'.ToCharArray()
    $special = '!@#$%^&*-_=+?'.ToCharArray()
    $all     = $upper + $lower + $digits + $special

    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    function Get-RandomChar([char[]]$set) {
        $b = New-Object byte[] 4
        $rng.GetBytes($b)
        $set[[math]::Abs([BitConverter]::ToInt32($b,0)) % $set.Count]
    }

    # Guarantee at least one of each required class
    $chars = @(
        (Get-RandomChar $upper), (Get-RandomChar $lower),
        (Get-RandomChar $digits), (Get-RandomChar $special)
    )
    for ($i = $chars.Count; $i -lt $Length; $i++) { $chars += Get-RandomChar $all }

    # Fisher-Yates shuffle so mandatory chars aren't always in front
    for ($i = $chars.Count - 1; $i -gt 0; $i--) {
        $b = New-Object byte[] 4; $rng.GetBytes($b)
        $j = [math]::Abs([BitConverter]::ToInt32($b,0)) % ($i + 1)
        $tmp = $chars[$i]; $chars[$i] = $chars[$j]; $chars[$j] = $tmp
    }
    $rng.Dispose()
    -join $chars
}

function Join-NameParts {
    param([string[]]$Parts)
    ($Parts | Where-Object { $_ -and $_.Trim() }) -join ' '
}

function Import-UserSheet {
    param([string]$Path)
    $ext = [IO.Path]::GetExtension($Path).ToLower()
    if ($ext -eq '.csv') {
        return @(Import-Csv -Path $Path)
    }
    # Prefer ImportExcel module, fall back to Excel COM
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        try {
            Install-Module ImportExcel -Scope CurrentUser -Force -ErrorAction Stop
        } catch { }
    }
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        return @(Import-Excel -Path $Path)
    }
    # COM fallback (requires Excel installed)
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    try {
        $wb  = $excel.Workbooks.Open($Path, $null, $true)
        $ws  = $wb.Worksheets.Item(1)
        $rng = $ws.UsedRange
        $rows = $rng.Rows.Count; $cols = $rng.Columns.Count
        $headers = for ($c = 1; $c -le $cols; $c++) { "$($rng.Cells.Item(1,$c).Text)".Trim() }
        $out = for ($r = 2; $r -le $rows; $r++) {
            $o = [ordered]@{}
            for ($c = 1; $c -le $cols; $c++) {
                $o[$headers[$c-1]] = "$($rng.Cells.Item($r,$c).Text)".Trim()
            }
            [pscustomobject]$o
        }
        $wb.Close($false)
        return @($out)
    } finally {
        $excel.Quit()
        [Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    }
}

function Get-Cell {
    param($Row, [string]$Name)
    $prop = $Row.PSObject.Properties |
        Where-Object { $_.Name.Trim() -ieq $Name } |
        Select-Object -First 1
    if ($prop -and $null -ne $prop.Value) { return "$($prop.Value)".Trim() }
    return ''
}

function ConvertTo-UserRecord {
    param($Row)
    $first  = Get-Cell $Row 'First name'
    $second = Get-Cell $Row 'Second name'
    $third  = Get-Cell $Row 'Third name'
    $last   = Get-Cell $Row 'Last name'
    $full   = Join-NameParts @($first, $second, $third, $last)

    [pscustomobject]@{
        employeeID     = Get-Cell $Row 'employeeID'
        employeeNumber = Get-Cell $Row 'employeeNumber'
        givenName      = $first
        sn             = Join-NameParts @($second, $third, $last)
        cn             = $full
        displayName    = $full
        name           = $full
        mobile         = Get-Cell $Row 'mobile'
        sAMAccountName = Get-Cell $Row 'sAMAccountName'
        mail           = Get-Cell $Row 'mail'
        title          = Get-Cell $Row 'title'
        occupation     = Get-Cell $Row 'occupation'
        cluster        = Get-Cell $Row 'cluster'
        clusterCode    = Get-Cell $Row 'clusterCode'
        facilty        = Get-Cell $Row 'facilty'
        faciltyCode    = Get-Cell $Row 'faciltyCode'
        employeeType   = Get-Cell $Row 'employeeType'
    }
}

function Get-ProvisioningCredential {
    <#
        Resolves the AD credential used for searches and user creation:
        1. Password typed in the UI PasswordBox (this session)
        2. DPAPI-encrypted password stored in the config file
        3. Interactive prompt as a last resort
        Returns $null when no username is configured (run as current user).
    #>
    param($Config)
    if (-not $Config.ADUser) { return $null }
    $secure = $null
    if ($pwdADPass.SecurePassword.Length -gt 0) {
        $secure = $pwdADPass.SecurePassword
    } elseif ($Config.ADPasswordEnc) {
        try { $secure = ConvertTo-SecureString "$($Config.ADPasswordEnc)" }
        catch { Write-Warning 'Stored password could not be decrypted on this machine/user - prompting.' }
    }
    if (-not $secure) {
        return (Get-Credential -UserName $Config.ADUser -Message 'Provisioning account password')
    }
    New-Object System.Management.Automation.PSCredential($Config.ADUser, $secure)
}

function Test-UserExists {
    <#
        Iterates the configured search filters (templates) against each
        Search DN using the Global Catalog when available.
        Filter templates support {sAMAccountName}, {mail}, {employeeID},
        {employeeNumber} placeholders, e.g. "(sAMAccountName={sAMAccountName})".

        For EACH query, tries a list of candidate servers in order (GC first,
        then a plain configured DC, then let ADWS auto-locate one) and uses
        the first one that executes without throwing. This tolerates a
        GC-specific query failing while a plain LDAP query on the same DC
        would have worked.

        Returns .Verified = $true if AT LEAST ONE search actually executed
        (hit or no-hit). If every attempt threw, .Verified = $false - the
        caller must treat this as "could not check", NOT as "doesn't exist".
    #>
    param($User, $Config, $Credential)

    $filters = @($Config.SearchFilters)
    if (-not $filters -or $filters.Count -eq 0) {
        $filters = @('(sAMAccountName={sAMAccountName})', '(mail={mail})')
    }
    $searchBases = @($Config.SearchDNs)
    if (-not $searchBases -or $searchBases.Count -eq 0) { $searchBases = @($Config.DomainDNs) }

    $gcServer = $null
    if ($Config.GlobalCatalog) {
        $gc = ($Config.GlobalCatalog -replace ':\d+$', '').Trim()   # tolerate ":3268" typed by mistake
        $gcServer = "${gc}:3268"
    }
    $candidateServers = @()
    if ($gcServer) { $candidateServers += $gcServer }
    if ($Config.Servers -and $Config.Servers.Count -gt 0) { $candidateServers += $Config.Servers[0] }
    $candidateServers += $null   # last resort: no explicit -Server, let ADWS pick

    $anySucceeded = $false
    foreach ($base in $searchBases) {
        if (-not $base) { continue }
        foreach ($tpl in $filters) {
            $ldap = $tpl -replace '\{sAMAccountName\}', $User.sAMAccountName `
                         -replace '\{mail\}',           $User.mail `
                         -replace '\{employeeID\}',     $User.employeeID `
                         -replace '\{employeeNumber\}', $User.employeeNumber
            # Skip filters whose placeholder resolved to empty (e.g. "(mail=)")
            if ($ldap -match '=\)') { continue }

            foreach ($srv in $candidateServers) {
                try {
                    $p = @{ LDAPFilter = $ldap; SearchBase = $base; ErrorAction = 'Stop' }
                    if ($srv)        { $p.Server = $srv }
                    if ($Credential) { $p.Credential = $Credential }
                    $hit = Get-ADObject @p | Select-Object -First 1
                    $anySucceeded = $true
                    if ($hit) {
                        return [pscustomobject]@{ Exists = $true; Filter = $ldap; DN = $hit.DistinguishedName; Verified = $true }
                    }
                    break   # this base/filter answered (no hit) - stop trying other servers for it
                } catch {
                    Write-Warning "Search failed [$base] [$ldap] via [$(if ($srv) {$srv} else {'auto'})]: $($_.Exception.Message)"
                    # fall through to the next candidate server for this same base/filter
                }
            }
        }
    }
    [pscustomobject]@{ Exists = $false; Filter = $null; DN = $null; Verified = $anySucceeded }
}

function Connect-ExchangeSession {
    # Opens ONE remote PowerShell session, trying each configured host in order.
    param($Config, $Credential)
    foreach ($exHost in @($Config.ExchangeHosts)) {
        if (-not $exHost) { continue }
        try {
            $p = @{
                ConfigurationName = 'Microsoft.Exchange'
                ConnectionUri     = "http://$exHost/PowerShell/"
                Authentication    = 'Kerberos'
                ErrorAction       = 'Stop'
            }
            if ($Credential) { $p.Credential = $Credential }
            return New-PSSession @p
        } catch {
            Write-Warning "Exchange [$exHost] failed: $($_.Exception.Message)"
        }
    }
    return $null
}

function Import-ExchangeCommands {
    <#
        Exchange remote sessions run a RESTRICTED "no-language mode" runspace:
        any scriptblock sent via Invoke-Command - even one using $using: - gets
        rewritten into a param() block behind the scenes, which the restricted
        runspace also rejects ("ScriptsNotAllowed"). Invoke-Command is simply
        not usable against this endpoint, with or without workarounds.

        The Microsoft-documented pattern for this exact situation is IMPLICIT
        REMOTING: Import-PSSession creates local proxy functions for the named
        cmdlets. Calling Enable-Mailbox / Get-Mailbox afterwards looks like a
        normal local call - no scriptblock is ever sent, so the restriction
        never applies. This returns the temporary module so it can be removed
        again in a finally block once the run is done.

        IMPORTANT: Import-PSSession only creates a proxy for a command if the
        connecting account's Exchange RBAC role actually exposes it in that
        session. If the account lacks rights (e.g. no Recipient Management
        role), the cmdlet is silently NOT imported - no error here, just a
        later "term not recognized". This function checks explicitly and
        throws immediately with an actionable message instead.
    #>
    param($Session)
    $requested = 'Get-Mailbox','Enable-Mailbox'
    $imported = Import-PSSession -Session $Session -CommandName $requested `
        -AllowClobber -DisableNameChecking -ErrorAction Stop

    $missing = $requested | Where-Object { -not (Get-Command $_ -Module $imported.Name -ErrorAction SilentlyContinue) }
    if ($missing) {
        throw "Connected to Exchange, but the account is not permitted to run: $($missing -join ', '). " +
              "This is an RBAC role issue, not a network/syntax problem - the Exchange User account " +
              "likely needs the 'Recipient Management' (or 'Mail Recipients') management role assigned. " +
              "An Exchange admin can check with: Get-ManagementRoleAssignment -RoleAssignee '<ExchangeUser>'"
    }
    return $imported
}

function Get-MailboxDbTracker {
    <#
        Builds an ordered tracker of the configured mailbox databases with
        the configured maximum, and a Count that is incremented locally as
        mailboxes are created THIS RUN - so capacity is respected without
        re-querying Exchange per user.

        By default (CountExisting = $false) Count starts at 0 for every
        database, so MaxMailboxes is "how many NEW mailboxes this run may
        place in this database" - it is never reduced by mailboxes that
        already exist there from earlier runs or other sources. This matches
        "if I say create 100, create 100" rather than treating MaxMailboxes
        as a hard total-capacity ceiling for the database.

        Pass -CountExisting to go back to the old behaviour and seed Count
        from a live Get-Mailbox query, i.e. enforce MaxMailboxes as a true
        ceiling on (pre-existing + new) mailboxes in the database.
        Requires Import-ExchangeCommands to have been called first in that
        case, so Get-Mailbox here is the imported proxy function, not a
        local cmdlet.
    #>
    param($Config, [switch]$CountExisting)
    $tracker = [ordered]@{}
    foreach ($db in @($Config.MailboxDatabases)) {
        if (-not $db.Name) { continue }
        $count = 0
        if ($CountExisting) {
            try {
                $count = @(Get-Mailbox -Database $db.Name -ResultSize Unlimited -ErrorAction Stop).Count
            } catch {
                Write-Warning "Could not read mailbox count for [$($db.Name)]: $($_.Exception.Message)"
            }
        }
        $tracker[$db.Name] = [pscustomobject]@{
            Max   = [int]$db.MaxMailboxes
            Count = [int]$count
        }
    }
    return $tracker
}

function Select-MailboxDatabase {
    # Fill       : first database (in configured order) that still has capacity
    # RoundRobin : database with the fewest mailboxes that still has capacity
    param($Tracker, [string]$Mode = 'Fill')
    $available = @($Tracker.Keys | Where-Object { $Tracker[$_].Count -lt $Tracker[$_].Max })
    if ($available.Count -eq 0) { return $null }
    if ($Mode -eq 'RoundRobin') {
        return ($available | Sort-Object { $Tracker[$_].Count } | Select-Object -First 1)
    }
    return $available[0]
}

function New-ProvisionedUser {
    param($User, $Config, [string]$TargetOU, [switch]$EnableMailbox,
          $ExchangeSession, $DbTracker, [string]$DbMode = 'Fill', $ADCredential, [switch]$DryRun)

    $result = [pscustomobject]@{
        sAMAccountName = $User.sAMAccountName
        DisplayName    = $User.displayName
        Mail           = $User.mail
        Password       = ''
        Database       = ''
        Status         = ''
        Detail         = ''
    }

    try {
        if (-not $User.sAMAccountName) { throw 'sAMAccountName is empty' }
        if (-not $User.givenName)      { throw 'First name is empty' }

        # Duplicate check across configured scope
        $dup = Test-UserExists -User $User -Config $Config -Credential $ADCredential
        if ($dup.Exists) {
            $result.Status = 'Skipped'
            $result.Detail = "Already exists: $($dup.DN) (matched $($dup.Filter))"
            return $result
        }
        if (-not $dup.Verified -and -not $DryRun) {
            # Every duplicate-check search failed - refuse to create rather than
            # risk a real duplicate. Fix AD connectivity/config and retry.
            $result.Status = 'Failed'
            $result.Detail = 'Duplicate check could not be verified (all AD searches failed) - refusing to create for safety. Check Global Catalog / Search DN / Servers / credentials.'
            return $result
        }

        $password = New-RandomPassword -Length 12
        $secure   = ConvertTo-SecureString $password -AsPlainText -Force

        # UPN = email as provided; fall back to sam@domain derived from target DN
        $upn = $User.mail
        if (-not $upn) {
            $domain = ((($TargetOU -split ',') | Where-Object { $_ -like 'DC=*' }) -replace 'DC=','') -join '.'
            $upn = "$($User.sAMAccountName)@$domain"
        }

        # Extended / non-parameter attributes go through -OtherAttributes
        $other = @{}
        foreach ($attr in 'occupation','cluster','clusterCode','facilty','faciltyCode') {
            if ($User.$attr) { $other[$attr] = $User.$attr }
        }
        if ($User.employeeType) { $other['employeeType'] = $User.employeeType }

        $params = @{
            Name                  = $User.name
            SamAccountName        = $User.sAMAccountName
            UserPrincipalName     = $upn
            GivenName             = $User.givenName
            DisplayName           = $User.displayName
            Path                  = $TargetOU
            AccountPassword       = $secure
            ChangePasswordAtLogon = $true      # must change at first login
            Enabled               = $true
            ErrorAction           = 'Stop'
        }
        if ($User.sn)             { $params.Surname        = $User.sn }
        if ($User.mail)           { $params.EmailAddress   = $User.mail }
        if ($User.title)          { $params.Title          = $User.title }
        if ($User.mobile)         { $params.MobilePhone    = $User.mobile }
        if ($User.employeeID)     { $params.EmployeeID     = $User.employeeID }
        if ($User.employeeNumber) { $params.EmployeeNumber = $User.employeeNumber }
        if ($other.Count -gt 0)   { $params.OtherAttributes = $other }
        if ($ADCredential)        { $params.Credential = $ADCredential }
        if ($Config.Servers -and $Config.Servers.Count -gt 0) { $params.Server = $Config.Servers[0] }

        if ($DryRun) {
            # Validate the target OU is reachable, then report intent without writing.
            $ouOk = $true; $ouMsg = ''
            try {
                $ouParams = @{ Identity = $TargetOU; ErrorAction = 'Stop' }
                if ($ADCredential) { $ouParams.Credential = $ADCredential }
                if ($Config.Servers -and $Config.Servers.Count -gt 0) { $ouParams.Server = $Config.Servers[0] }
                Get-ADObject @ouParams | Out-Null
            } catch {
                $ouOk = $false; $ouMsg = " | TARGET OU NOT REACHABLE: $($_.Exception.Message)"
            }
            $result.Password = '(dry run - not generated)'
            $result.Status   = $(if (-not $ouOk) { 'WouldFail' } elseif (-not $dup.Verified) { 'WouldCreate-UNVERIFIED' } else { 'WouldCreate' })
            $unverifiedMsg = $(if (-not $dup.Verified) { ' | WARNING: duplicate check UNVERIFIED - all AD searches failed, this is NOT a confirmed non-duplicate' } else { '' })
            $result.Detail   = "DRY RUN - would create in $TargetOU; UPN=$upn; sn=$($User.sn); cn=$($User.cn)$ouMsg$unverifiedMsg"

            if ($EnableMailbox -and $DbTracker -and $DbTracker.Count -gt 0) {
                $planDb = Select-MailboxDatabase -Tracker $DbTracker -Mode $DbMode
                if ($planDb) {
                    $DbTracker[$planDb].Count++      # simulate placement so distribution is testable
                    $result.Database = $planDb
                    $result.Detail  += " | would use DB $planDb ($($DbTracker[$planDb].Count)/$($DbTracker[$planDb].Max))"
                } else {
                    $result.Detail  += ' | NO DATABASE CAPACITY'
                }
            }
            return $result
        }

        try {
            New-ADUser @params
        } catch {
            # If a custom attribute is not in the schema, retry without it
            if ($_.Exception.Message -match 'attribute' -and $other.Count -gt 0) {
                $failed = $other.Keys -join ','
                $params.Remove('OtherAttributes')
                New-ADUser @params
                $result.Detail = "Created without custom attributes ($failed): $($_.Exception.Message)"
            } else { throw }
        }

        $result.Password = $password
        $result.Status   = 'Created'
        if (-not $result.Detail) { $result.Detail = "OU: $TargetOU" }

        # Optional Exchange mailbox enablement (shared session, capacity-aware DB selection)
        if ($EnableMailbox -and $ExchangeSession) {
            $dbName = $null
            $skipMailbox = $false

            if ($DbTracker -and $DbTracker.Count -gt 0) {
                $dbName = Select-MailboxDatabase -Tracker $DbTracker -Mode $DbMode
                if (-not $dbName) {
                    $result.Detail += ' | Mailbox NOT enabled: all configured databases at max capacity'
                    $skipMailbox = $true
                }
            }

            if (-not $skipMailbox) {
                # Pin Exchange's AD read/write to the SAME domain controller that
                # New-ADUser just wrote the account to (read-your-own-write).
                # Without -DomainController, Exchange's own topology service picks
                # a DC on its own - under a fast bulk-create batch, the object can
                # easily not have replicated there yet, producing
                # "object '' couldn't be found on <DC>" even though the account
                # was created successfully a moment earlier on a different DC.
                $writeDC = $(if ($Config.Servers -and $Config.Servers.Count -gt 0) { $Config.Servers[0] } else { $null })

                $mbxParams = @{ Identity = $User.sAMAccountName; ErrorAction = 'Stop' }
                if ($dbName)  { $mbxParams.Database = $dbName }
                if ($writeDC) { $mbxParams.DomainController = $writeDC }

                $maxAttempts = 4
                $attempt = 0
                $lastErr = $null
                $enabled = $false
                while ($attempt -lt $maxAttempts -and -not $enabled) {
                    $attempt++
                    try {
                        Enable-Mailbox @mbxParams | Out-Null
                        $enabled = $true
                    } catch {
                        $lastErr = $_
                        # "couldn't be found" / "couldn't be performed because object"
                        # is the AD-replication-lag signature - worth a short wait
                        # and retry rather than failing outright. Other errors
                        # (permissions, bad database, etc.) are not retried.
                        if ($attempt -lt $maxAttempts -and $_.Exception.Message -match "couldn't be found|couldn't be performed") {
                            Start-Sleep -Seconds ([math]::Min(3 * $attempt, 10))
                        } else {
                            break
                        }
                    }
                }

                if ($enabled) {
                    if ($dbName) {
                        $DbTracker[$dbName].Count++
                        $result.Database = $dbName
                        $result.Detail  += " | Mailbox in $dbName ($($DbTracker[$dbName].Count)/$($DbTracker[$dbName].Max))"
                    } else {
                        $result.Database = '(auto)'
                        $result.Detail  += ' | Mailbox enabled (automatic database)'
                    }
                    if ($attempt -gt 1) { $result.Detail += " | (succeeded on retry $attempt/$maxAttempts - AD replication lag)" }
                } else {
                    $result.Detail += " | Mailbox enablement FAILED after $attempt attempt(s): $($lastErr.Exception.Message)"
                }
            }
        }
    }
    catch {
        $result.Status = 'Failed'
        $result.Detail = $_.Exception.Message
    }
    return $result
}

# ------------------------------------------------------------------ #
#  Config persistence                                                #
# ------------------------------------------------------------------ #

function Get-UIConfig {
    # If a new password was typed, encrypt it (DPAPI: current user + machine only)
    if ($pwdADPass.SecurePassword.Length -gt 0) {
        $script:ADPassEnc = ConvertFrom-SecureString $pwdADPass.SecurePassword
    }
    [pscustomobject]@{
        ForestName    = $txtForestName.Text.Trim()
        GlobalCatalog = $txtGC.Text.Trim()
        ForestDN      = $txtForestDN.Text.Trim()
        ADUser        = $txtADUser.Text.Trim()
        ADPasswordEnc = $script:ADPassEnc
        DomainDNs     = @($lstDomainDNs.Items)
        Servers       = @($lstServers.Items)
        ExchangeForest= $txtExForest.Text.Trim()
        ExchangeHosts = @($lstExHosts.Items)
        ExchangeUser  = $txtExUser.Text.Trim()
        SearchDNs     = @($lstSearchDNs.Items)
        SearchFilters = @($lstFilters.Items)
        MailboxDatabases = @($lstMbxDbs.Items | ForEach-Object {
            $n, $m = "$_" -split '\|', 2
            [pscustomobject]@{ Name = $n; MaxMailboxes = [int]$m }
        })
        DbDistribution = $(if ($cmbDbMode.SelectedIndex -eq 1) { 'RoundRobin' } else { 'Fill' })
        CountExistingMailboxes = [bool]$chkCountExisting.IsChecked
        TargetOU      = $txtTargetOU.Text.Trim()
    }
}

function Save-UIConfig {
    Get-UIConfig | ConvertTo-Json -Depth 5 | Set-Content -Path $script:ConfigPath -Encoding UTF8
    $lblStatus.Text = "Configuration saved: $script:ConfigPath"
}

function Load-UIConfig {
    if (-not (Test-Path $script:ConfigPath)) { return }
    try {
        $cfg = Get-Content $script:ConfigPath -Raw | ConvertFrom-Json
        $txtForestName.Text = "$($cfg.ForestName)"
        $txtGC.Text         = "$($cfg.GlobalCatalog)"
        $txtForestDN.Text   = "$($cfg.ForestDN)"
        $txtExForest.Text   = "$($cfg.ExchangeForest)"
        $txtExUser.Text     = "$($cfg.ExchangeUser)"
        $txtADUser.Text     = "$($cfg.ADUser)"
        $script:ADPassEnc   = if ("$($cfg.ADPasswordEnc)") { "$($cfg.ADPasswordEnc)" } else { $null }
        if ($cfg.TargetOU) { $txtTargetOU.Text = "$($cfg.TargetOU)" }
        foreach ($pair in @(
            @($lstDomainDNs, $cfg.DomainDNs), @($lstServers, $cfg.Servers),
            @($lstExHosts,  $cfg.ExchangeHosts),
            @($lstSearchDNs,$cfg.SearchDNs), @($lstFilters, $cfg.SearchFilters))) {
            $list, $vals = $pair
            $list.Items.Clear()
            foreach ($v in @($vals)) { if ($v) { [void]$list.Items.Add("$v") } }
        }
        $lstMbxDbs.Items.Clear()
        foreach ($db in @($cfg.MailboxDatabases)) {
            if ($db -and $db.Name) { [void]$lstMbxDbs.Items.Add("$($db.Name)|$($db.MaxMailboxes)") }
        }
        $cmbDbMode.SelectedIndex = $(if ("$($cfg.DbDistribution)" -eq 'RoundRobin') { 1 } else { 0 })
        $chkCountExisting.IsChecked = [bool]$cfg.CountExistingMailboxes
        $lblStatus.Text = 'Configuration loaded.'
    } catch {
        $lblStatus.Text = "Failed to load config: $($_.Exception.Message)"
    }
}

# ------------------------------------------------------------------ #
#  XAML UI                                                           #
# ------------------------------------------------------------------ #

$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="AD User Provisioning Tool" Height="720" Width="980"
        WindowStartupLocation="CenterScreen" FontSize="13">
  <DockPanel>
    <StatusBar DockPanel.Dock="Bottom">
      <TextBlock Name="lblStatus" Text="Ready"/>
    </StatusBar>
    <TabControl Name="tabMain" Margin="8">

      <!-- ===================== CONFIGURATION TAB ===================== -->
      <TabItem Header="Configuration">
        <ScrollViewer VerticalScrollBarVisibility="Auto">
          <StackPanel Margin="10">

            <GroupBox Header="Forest Configuration" Padding="8" Margin="0,0,0,10">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="180"/><ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition/><RowDefinition/><RowDefinition/>
                </Grid.RowDefinitions>
                <Label Content="Forest Name" Grid.Row="0"/>
                <TextBox Name="txtForestName" Grid.Row="0" Grid.Column="1" Margin="2"/>
                <Label Content="Global Catalog Server" Grid.Row="1"/>
                <TextBox Name="txtGC" Grid.Row="1" Grid.Column="1" Margin="2"/>
                <Label Content="Domain DN" Grid.Row="2"/>
                <TextBox Name="txtForestDN" Grid.Row="2" Grid.Column="1" Margin="2"/>
              </Grid>
            </GroupBox>

            <GroupBox Header="Domain Configuration" Padding="8" Margin="0,0,0,10">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/><ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Margin="0,0,8,0">
                  <Label Content="Domain DNs (multiple)"/>
                  <DockPanel>
                    <Button Name="btnAddDomainDN" Content="Add" Width="60" DockPanel.Dock="Right" Margin="4,2,0,2"/>
                    <TextBox Name="txtNewDomainDN" Margin="2"/>
                  </DockPanel>
                  <ListBox Name="lstDomainDNs" Height="90" Margin="2"/>
                  <Button Name="btnDelDomainDN" Content="Remove selected" Width="140" HorizontalAlignment="Left" Margin="2"/>
                </StackPanel>
                <StackPanel Grid.Column="1">
                  <Label Content="Servers / Domain Controllers (multiple)"/>
                  <DockPanel>
                    <Button Name="btnAddServer" Content="Add" Width="60" DockPanel.Dock="Right" Margin="4,2,0,2"/>
                    <TextBox Name="txtNewServer" Margin="2"/>
                  </DockPanel>
                  <ListBox Name="lstServers" Height="90" Margin="2"/>
                  <Button Name="btnDelServer" Content="Remove selected" Width="140" HorizontalAlignment="Left" Margin="2"/>
                </StackPanel>
              </Grid>
            </GroupBox>

            <GroupBox Header="Provisioning Account (used to create users)" Padding="8" Margin="0,0,0,10">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="180"/><ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <Grid.RowDefinitions>
                  <RowDefinition/><RowDefinition/><RowDefinition/>
                </Grid.RowDefinitions>
                <Label Content="Username" Grid.Row="0"/>
                <TextBox Name="txtADUser" Grid.Row="0" Grid.Column="1" Margin="2"
                         ToolTip="DOMAIN\account with rights to create users in the target OU. Leave empty to run as the currently logged-on user."/>
                <Label Content="Password" Grid.Row="1"/>
                <PasswordBox Name="pwdADPass" Grid.Row="1" Grid.Column="1" Margin="2"/>
                <TextBlock Grid.Row="2" Grid.ColumnSpan="2" Margin="4,4,4,0" TextWrapping="Wrap" Foreground="Gray"
                           Text="The password is saved DPAPI-encrypted (only decryptable by this Windows user on this machine) - never in plain text. Leave the box empty on later saves to keep the stored password. If neither is available at run time you will be prompted."/>
              </Grid>
            </GroupBox>

            <GroupBox Header="Exchange Configuration" Padding="8" Margin="0,0,0,10">
              <StackPanel>
                <Grid>
                  <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="180"/><ColumnDefinition Width="*"/>
                  </Grid.ColumnDefinitions>
                  <Grid.RowDefinitions><RowDefinition/><RowDefinition/></Grid.RowDefinitions>
                  <Label Content="Exchange Forest" Grid.Row="0"/>
                  <TextBox Name="txtExForest" Grid.Row="0" Grid.Column="1" Margin="2"/>
                  <Label Content="Exchange User" Grid.Row="1"/>
                  <TextBox Name="txtExUser" Grid.Row="1" Grid.Column="1" Margin="2"
                           ToolTip="DOMAIN\account used to connect to Exchange (prompted for password at run time)"/>
                </Grid>
                <Label Content="Exchange Hosts (multiple)"/>
                <DockPanel>
                  <Button Name="btnAddExHost" Content="Add" Width="60" DockPanel.Dock="Right" Margin="4,2,0,2"/>
                  <TextBox Name="txtNewExHost" Margin="2"/>
                </DockPanel>
                <ListBox Name="lstExHosts" Height="70" Margin="2"/>
                <Button Name="btnDelExHost" Content="Remove selected" Width="140" HorizontalAlignment="Left" Margin="2"/>

                <Label Content="Mailbox Databases (multiple) - database name + max mailboxes per database"/>
                <DockPanel>
                  <Button Name="btnAddMbxDb" Content="Add" Width="60" DockPanel.Dock="Right" Margin="4,2,0,2"/>
                  <TextBox Name="txtNewMbxDbMax" Width="90" DockPanel.Dock="Right" Margin="4,2,0,2"
                           ToolTip="Max mailboxes in this database - a NUMBER, e.g. 500"/>
                  <Label Content="Max:" DockPanel.Dock="Right" Margin="4,2,0,2" VerticalAlignment="Center"/>
                  <TextBox Name="txtNewMbxDbName" Margin="2" ToolTip="Database name, e.g. DAG01-DB001"/>
                </DockPanel>
                <ListBox Name="lstMbxDbs" Height="70" Margin="2"/>
                <DockPanel LastChildFill="False">
                  <Button Name="btnDelMbxDb" Content="Remove selected" Width="140" Margin="2"/>
                  <Label Content="Distribution:" Margin="12,0,0,0"/>
                  <ComboBox Name="cmbDbMode" Width="150" Margin="2" SelectedIndex="0">
                    <ComboBoxItem Content="Fill sequentially"/>
                    <ComboBoxItem Content="Round-robin (least loaded)"/>
                  </ComboBox>
                </DockPanel>
                <CheckBox Name="chkCountExisting" Margin="2,6,2,2" IsChecked="False">
                  <CheckBox.Content>
                    <TextBlock TextWrapping="Wrap" MaxWidth="820"
                               Text="Count pre-existing mailboxes in each database against Max (unchecked = Max applies only to mailboxes created by this run; a batch of 100 always creates 100 regardless of what's already in the database)"/>
                  </CheckBox.Content>
                </CheckBox>
              </StackPanel>
            </GroupBox>

            <GroupBox Header="User Search Scope" Padding="8" Margin="0,0,0,10">
              <Grid>
                <Grid.ColumnDefinitions>
                  <ColumnDefinition Width="*"/><ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Margin="0,0,8,0">
                  <Label Content="Search DNs - container DNs (multiple)"/>
                  <DockPanel>
                    <Button Name="btnAddSearchDN" Content="Add" Width="60" DockPanel.Dock="Right" Margin="4,2,0,2"/>
                    <TextBox Name="txtNewSearchDN" Margin="2"/>
                  </DockPanel>
                  <ListBox Name="lstSearchDNs" Height="90" Margin="2"/>
                  <Button Name="btnDelSearchDN" Content="Remove selected" Width="140" HorizontalAlignment="Left" Margin="2"/>
                </StackPanel>
                <StackPanel Grid.Column="1">
                  <Label Content="Iterated Search Filters (multiple)"/>
                  <DockPanel>
                    <Button Name="btnAddFilter" Content="Add" Width="60" DockPanel.Dock="Right" Margin="4,2,0,2"/>
                    <TextBox Name="txtNewFilter" Margin="2"
                             ToolTip="LDAP filter template. Placeholders: {sAMAccountName} {mail} {employeeID} {employeeNumber}"/>
                  </DockPanel>
                  <ListBox Name="lstFilters" Height="90" Margin="2"/>
                  <Button Name="btnDelFilter" Content="Remove selected" Width="140" HorizontalAlignment="Left" Margin="2"/>
                </StackPanel>
              </Grid>
            </GroupBox>

            <StackPanel Orientation="Horizontal">
              <Button Name="btnSaveConfig" Content="Save Configuration" Width="160" Margin="2" Padding="4"/>
              <Button Name="btnLoadConfig" Content="Reload Configuration" Width="160" Margin="2" Padding="4"/>
              <Button Name="btnTestConn" Content="Test AD Connection" Width="160" Margin="2" Padding="4"/>
            </StackPanel>
          </StackPanel>
        </ScrollViewer>
      </TabItem>

      <!-- ===================== USER CREATION TAB ===================== -->
      <TabItem Header="Create Users">
        <Grid Margin="10">
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <Grid Grid.Row="0">
            <Grid.ColumnDefinitions>
              <ColumnDefinition Width="140"/><ColumnDefinition Width="*"/><ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions><RowDefinition/><RowDefinition/></Grid.RowDefinitions>
            <Label Content="Excel / CSV file" Grid.Row="0"/>
            <TextBox Name="txtExcelPath" Grid.Row="0" Grid.Column="1" Margin="2" IsReadOnly="True"/>
            <Button Name="btnBrowse" Content="Browse..." Grid.Row="0" Grid.Column="2" Width="90" Margin="4,2,0,2"/>
            <Label Content="Target OU (DN)" Grid.Row="1"/>
            <TextBox Name="txtTargetOU" Grid.Row="1" Grid.Column="1" Grid.ColumnSpan="2" Margin="2"
                     Text="OU=Users,OU=YourOrg,DC=contoso,DC=com"/>
          </Grid>

          <StackPanel Grid.Row="1" Orientation="Horizontal" Margin="0,6,0,6">
            <Button Name="btnLoadSheet" Content="Load and Preview" Width="150" Padding="4" Margin="0,0,6,0"/>
            <CheckBox Name="chkMailbox" Content="Enable Exchange mailbox" VerticalAlignment="Center" Margin="0,0,12,0"/>
            <CheckBox Name="chkDryRun" Content="DRY RUN (validate only - create nothing)" VerticalAlignment="Center"
                      Margin="0,0,12,0" Foreground="DarkGreen" FontWeight="Bold" IsChecked="True"/>
            <Button Name="btnCreate" Content="Create Users" Width="150" Padding="4" IsEnabled="False" FontWeight="Bold"/>
          </StackPanel>

          <GroupBox Grid.Row="2" Header="Preview (mapped AD attributes)">
            <DataGrid Name="gridPreview" AutoGenerateColumns="True" IsReadOnly="True"
                      CanUserAddRows="False" HeadersVisibility="Column"/>
          </GroupBox>

          <Label Grid.Row="3" Content="Results" FontWeight="Bold" Margin="0,6,0,0"/>

          <DataGrid Name="gridResults" Grid.Row="4" AutoGenerateColumns="True" IsReadOnly="True"
                    CanUserAddRows="False" HeadersVisibility="Column"/>

          <StackPanel Grid.Row="5" Orientation="Horizontal" Margin="0,6,0,0">
            <Button Name="btnExport" Content="Export Results (CSV)" Width="170" Padding="4" IsEnabled="False"/>
            <TextBlock Text="  Exported file contains initial passwords - store securely and delete after handover."
                       VerticalAlignment="Center" Foreground="DarkRed"/>
          </StackPanel>
        </Grid>
      </TabItem>
    </TabControl>
  </DockPanel>
</Window>
'@

$reader = [System.Xml.XmlNodeReader]::new([xml]$xaml)
$window = [Windows.Markup.XamlReader]::Load($reader)

# Bind named elements to variables
([xml]$xaml).SelectNodes('//*[@Name]') | ForEach-Object {
    Set-Variable -Name $_.Name -Value $window.FindName($_.Name) -Scope Script
}

# ------------------------------------------------------------------ #
#  Wire up list add/remove buttons                                   #
# ------------------------------------------------------------------ #

$listPairs = @(
    @{ Add = $btnAddDomainDN; Del = $btnDelDomainDN; Txt = $txtNewDomainDN; Lst = $lstDomainDNs },
    @{ Add = $btnAddServer;   Del = $btnDelServer;   Txt = $txtNewServer;   Lst = $lstServers },
    @{ Add = $btnAddExHost;   Del = $btnDelExHost;   Txt = $txtNewExHost;   Lst = $lstExHosts },
    @{ Add = $btnAddSearchDN; Del = $btnDelSearchDN; Txt = $txtNewSearchDN; Lst = $lstSearchDNs },
    @{ Add = $btnAddFilter;   Del = $btnDelFilter;   Txt = $txtNewFilter;   Lst = $lstFilters }
)
foreach ($p in $listPairs) {
    $txt = $p.Txt; $lst = $p.Lst
    $p.Add.Add_Click({
        param($s, $e)
        $pair = $listPairs | Where-Object { $_.Add -eq $s }
        $v = $pair.Txt.Text.Trim()
        if (-not $v) { return }

        # Domain DNs and Search DNs must be ONE full distinguished name per entry,
        # e.g. "OU=Users,OU=YourOrg,DC=contoso,DC=com" - not split into
        # DC=contoso / DC=com as separate entries. Catch that here instead of
        # failing later with a confusing "unable to contact the server" error.
        if (($pair.Lst -eq $lstDomainDNs -or $pair.Lst -eq $lstSearchDNs) -and
            $v -match '^(DC|OU|CN)=[^,]+$') {
            $answer = [System.Windows.MessageBox]::Show(
                "`"$v`" looks like ONE PIECE of a distinguished name, not the full DN.`n`n" +
                "This field expects a single complete DN per entry, e.g.:`n" +
                "OU=Users,OU=YourOrg,DC=contoso,DC=com`n`n" +
                "Add it anyway?",
                'This looks incomplete', 'YesNo', 'Warning')
            if ($answer -ne 'Yes') { return }
        }

        if (-not $pair.Lst.Items.Contains($v)) {
            [void]$pair.Lst.Items.Add($v)
            $pair.Txt.Clear()
        }
    }.GetNewClosure())
    $p.Del.Add_Click({
        param($s, $e)
        $pair = $listPairs | Where-Object { $_.Del -eq $s }
        while ($pair.Lst.SelectedItems.Count -gt 0) {
            $pair.Lst.Items.Remove($pair.Lst.SelectedItems[0])
        }
    }.GetNewClosure())
}

# Mailbox database list (two-field entry: name + max mailboxes)
$btnAddMbxDb.Add_Click({
    $name = $txtNewMbxDbName.Text.Trim()
    $maxText = $txtNewMbxDbMax.Text.Trim()
    $max  = 0
    $parsedOk = [int]::TryParse($maxText, [ref]$max)

    if (-not $name) {
        [System.Windows.MessageBox]::Show('Enter a database name in the left box.', 'Missing name', 'OK', 'Warning') | Out-Null
        return
    }
    if (-not $parsedOk -or $max -le 0) {
        [System.Windows.MessageBox]::Show(
            "The Max Mailboxes box (right side) must contain a NUMBER, e.g. 500 - not the database name.`n`nYou entered: '$maxText'",
            'Invalid max mailbox count', 'OK', 'Warning') | Out-Null
        return
    }
    if (@($lstMbxDbs.Items | Where-Object { "$_" -like "$name|*" })) {
        [System.Windows.MessageBox]::Show("`"$name`" is already in the list.", 'Duplicate', 'OK', 'Information') | Out-Null
        return
    }

    [void]$lstMbxDbs.Items.Add("$name|$max")
    $txtNewMbxDbName.Clear(); $txtNewMbxDbMax.Clear()
    $lblStatus.Text = "Added database: $name (max $max)"
})
$btnDelMbxDb.Add_Click({
    while ($lstMbxDbs.SelectedItems.Count -gt 0) {
        $lstMbxDbs.Items.Remove($lstMbxDbs.SelectedItems[0])
    }
})

# ------------------------------------------------------------------ #
#  Config buttons                                                    #
# ------------------------------------------------------------------ #

$btnSaveConfig.Add_Click({ Save-UIConfig })
$btnLoadConfig.Add_Click({ Load-UIConfig })

$btnTestConn.Add_Click({
    try {
        Import-Module ActiveDirectory -ErrorAction Stop
        $cfg = Get-UIConfig
        $p = @{ ErrorAction = 'Stop' }
        if ($cfg.Servers -and $cfg.Servers.Count -gt 0) { $p.Server = $cfg.Servers[0] }
        $cred = Get-ProvisioningCredential -Config $cfg
        if ($cred) { $p.Credential = $cred }
        $d = Get-ADDomain @p
        $lblStatus.Text = "Connected: $($d.DNSRoot) ($($d.DistinguishedName)) as $(if ($cred) { $cred.UserName } else { $env:USERDOMAIN + '\' + $env:USERNAME })"
        [System.Windows.MessageBox]::Show("Connected to domain: $($d.DNSRoot)", 'AD Connection', 'OK', 'Information') | Out-Null
    } catch {
        $lblStatus.Text = "Connection failed: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show($_.Exception.Message, 'AD Connection Failed', 'OK', 'Error') | Out-Null
    }
})

# ------------------------------------------------------------------ #
#  Sheet load / create / export                                      #
# ------------------------------------------------------------------ #

$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Excel/CSV files (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|All files (*.*)|*.*'
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtExcelPath.Text = $dlg.FileName
    }
})

$btnLoadSheet.Add_Click({
    if (-not $txtExcelPath.Text) {
        [System.Windows.MessageBox]::Show('Select an Excel or CSV file first.', 'No file', 'OK', 'Warning') | Out-Null
        return
    }
    try {
        $lblStatus.Text = 'Loading sheet...'
        $rows = Import-UserSheet -Path $txtExcelPath.Text
        $script:Preview = @($rows | ForEach-Object { ConvertTo-UserRecord $_ })
        $gridPreview.ItemsSource = $script:Preview
        $btnCreate.IsEnabled = ($script:Preview.Count -gt 0)
        $lblStatus.Text = "$($script:Preview.Count) user record(s) loaded and mapped."
    } catch {
        $lblStatus.Text = "Load failed: $($_.Exception.Message)"
        [System.Windows.MessageBox]::Show($_.Exception.Message, 'Load Failed', 'OK', 'Error') | Out-Null
    }
})

$btnCreate.Add_Click({
    if ($script:Preview.Count -eq 0) { return }
    $cfg = Get-UIConfig
    $ou  = $txtTargetOU.Text.Trim()
    if (-not $ou) {
        [System.Windows.MessageBox]::Show('Target OU is required.', 'Missing OU', 'OK', 'Warning') | Out-Null
        return
    }
    $dry = [bool]$chkDryRun.IsChecked
    $confirm = [System.Windows.MessageBox]::Show(
        $(if ($dry) {
            "DRY RUN`n`nValidate $($script:Preview.Count) user(s) against:`n$ou`n`nNothing will be created. Continue?"
        } else {
            "LIVE RUN`n`nCreate $($script:Preview.Count) user(s) in:`n$ou`n`nThis WILL create real accounts. Continue?"
        }),
        'Confirm', 'YesNo', 'Question')
    if ($confirm -ne 'Yes') { return }

    try { Import-Module ActiveDirectory -ErrorAction Stop }
    catch {
        [System.Windows.MessageBox]::Show('ActiveDirectory module not available. Install RSAT.', 'Error', 'OK', 'Error') | Out-Null
        return
    }

    # Resolve the provisioning account (UI password > stored encrypted > prompt)
    $adCred = Get-ProvisioningCredential -Config $cfg
    if ($cfg.ADUser -and -not $adCred) { return }   # user cancelled the prompt

    $exSession = $null
    $exModule  = $null
    $dbTracker = $null
    if ($chkMailbox.IsChecked) {
        if (@($cfg.ExchangeHosts).Count -eq 0 -and -not $dry) {
            [System.Windows.MessageBox]::Show('Mailbox enablement is checked but no Exchange Hosts are configured. Add at least one host on the Configuration tab.', 'Exchange', 'OK', 'Warning') | Out-Null
            return
        }

        if ($dry) {
            # Dry run never touches Exchange: simulate the databases from config,
            # starting each at 0 so distribution logic can be verified offline.
            if (@($cfg.MailboxDatabases).Count -gt 0) {
                $dbTracker = [ordered]@{}
                foreach ($db in @($cfg.MailboxDatabases)) {
                    if ($db.Name) { $dbTracker[$db.Name] = [pscustomobject]@{ Max = [int]$db.MaxMailboxes; Count = 0 } }
                }
            }
        } else {
            $exCred = $null
            if ($cfg.ExchangeUser) {
                $exCred = Get-Credential -UserName $cfg.ExchangeUser -Message 'Exchange service account password'
                if (-not $exCred) { return }
            }
            $lblStatus.Text = 'Connecting to Exchange...'
            $window.Dispatcher.Invoke([action]{}, 'Background')
            $exSession = Connect-ExchangeSession -Config $cfg -Credential $exCred
            if (-not $exSession) {
                [System.Windows.MessageBox]::Show('Could not connect to any configured Exchange host.', 'Exchange', 'OK', 'Error') | Out-Null
                return
            }
            try {
                $lblStatus.Text = 'Importing Exchange commands (implicit remoting)...'
                $window.Dispatcher.Invoke([action]{}, 'Background')
                $exModule = Import-ExchangeCommands -Session $exSession
            } catch {
                [System.Windows.MessageBox]::Show("Could not import Exchange commands: $($_.Exception.Message)", 'Exchange', 'OK', 'Error') | Out-Null
                Remove-PSSession $exSession -ErrorAction SilentlyContinue
                return
            }
            if (@($cfg.MailboxDatabases).Count -gt 0) {
                $lblStatus.Text = 'Reading current mailbox counts per database...'
                $window.Dispatcher.Invoke([action]{}, 'Background')
                $countExisting = [bool]$cfg.CountExistingMailboxes
                $dbTracker = Get-MailboxDbTracker -Config $cfg -CountExisting:$countExisting
            }
        }
    }

    $logDir = Join-Path $script:ScriptDir 'Logs'
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $logFile = Join-Path $logDir "Run_$(Get-Date -Format 'yyyyMMdd_HHmmss')_$(if ($dry) {'DRYRUN'} else {'LIVE'}).log"
    $loggingOn = $false
    try { Start-Transcript -Path $logFile -Append -ErrorAction Stop | Out-Null; $loggingOn = $true }
    catch { Write-Warning "Could not start log file: $($_.Exception.Message)" }

    try {
        $script:Results.Clear()
        $gridResults.ItemsSource = $script:Results
        $done = 0
        foreach ($u in $script:Preview) {
            $lblStatus.Text = "Processing $($u.sAMAccountName) ($(++$done)/$($script:Preview.Count))..."
            $window.Dispatcher.Invoke([action]{}, 'Background')  # keep UI responsive
            $r = New-ProvisionedUser -User $u -Config $cfg -TargetOU $ou `
                                     -EnableMailbox:([bool]$chkMailbox.IsChecked) `
                                     -ExchangeSession $exSession -DbTracker $dbTracker `
                                     -DbMode "$($cfg.DbDistribution)" -ADCredential $adCred `
                                     -DryRun:$dry
            $script:Results.Add($r)
        }
    } finally {
        if ($exModule)  { Remove-Module -Name $exModule.Name -ErrorAction SilentlyContinue }
        if ($exSession) { Remove-PSSession $exSession -ErrorAction SilentlyContinue }
        if ($loggingOn) { try { Stop-Transcript | Out-Null } catch { } }
    }
    if ($dry) {
        $would  = @($script:Results | Where-Object Status -eq 'WouldCreate').Count
        $unver  = @($script:Results | Where-Object Status -eq 'WouldCreate-UNVERIFIED').Count
        $wskip  = @($script:Results | Where-Object Status -eq 'Skipped').Count
        $wfail  = @($script:Results | Where-Object { $_.Status -in 'WouldFail','Failed' }).Count
        $lblStatus.Text = "DRY RUN complete (nothing created). Would create: $would | UNVERIFIED (duplicate check failed): $unver | Would skip (existing): $wskip | Would fail: $wfail | Log: $logFile"
    } else {
        $created = @($script:Results | Where-Object Status -eq 'Created').Count
        $skipped = @($script:Results | Where-Object Status -eq 'Skipped').Count
        $failed  = @($script:Results | Where-Object Status -eq 'Failed').Count
        $lblStatus.Text = "Done. Created: $created | Skipped (existing): $skipped | Failed: $failed | Log: $logFile"
    }
    $btnExport.IsEnabled = ($script:Results.Count -gt 0)
})

$btnExport.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'CSV files (*.csv)|*.csv'
    $dlg.FileName = "ADProvisioning_Results_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:Results | Export-Csv -Path $dlg.FileName -NoTypeInformation -Encoding UTF8
        $lblStatus.Text = "Results exported: $($dlg.FileName)"
    }
})

# ------------------------------------------------------------------ #
#  Startup                                                           #
# ------------------------------------------------------------------ #

Load-UIConfig
[void]$window.ShowDialog()
