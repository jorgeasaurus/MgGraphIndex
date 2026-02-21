<#
.SYNOPSIS
    Extracts Microsoft Graph PowerShell cmdlet metadata into JSON for the reference site.

.DESCRIPTION
    Dynamically discovers all installed Microsoft.Graph.* modules, extracts syntax,
    examples, descriptions, and permissions, then writes a two-tier data structure:
      - public/data/manifest.json   (lightweight summary for all cmdlets)
      - public/data/modules/*.json  (full detail per module, loaded lazily)
      - public/data/cmdlets.json    (flat file for backward compatibility)

    Both the permission lookup and metadata extraction phases run in parallel using
    ForEach-Object -Parallel (PowerShell 7+). Use -ThrottleLimit to tune concurrency.

.PARAMETER OutputPath
    Where to write the flat JSON file. Default: ../public/data/cmdlets.json

.PARAMETER IncludeBeta
    Also scan Microsoft.Graph.Beta.* modules.

.PARAMETER MaxCmdlets
    Cap the number of cmdlets processed (0 = no limit). Useful for testing.

.PARAMETER SkipPermissions
    Skip the Find-MgGraphCommand permission lookup entirely. Much faster.

.PARAMETER ThrottleLimit
    Max parallel threads. Default: number of logical processors.

.EXAMPLE
    .\get-graphcmdlets.ps1
    .\get-graphcmdlets.ps1 -IncludeBeta -Verbose
    .\get-graphcmdlets.ps1 -MaxCmdlets 50
    .\get-graphcmdlets.ps1 -SkipPermissions -ThrottleLimit 16
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path $PSScriptRoot '..' 'public' 'data' 'cmdlets.json'),
    [switch]$IncludeBeta,
    [int]$MaxCmdlets = 0,
    [switch]$SkipPermissions,
    [int]$ThrottleLimit = [Environment]::ProcessorCount
)

$ErrorActionPreference = 'Continue'
$ProgressPreference = 'SilentlyContinue'

if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Error "This script requires PowerShell 7+ for ForEach-Object -Parallel."
    exit 1
}

$timer = [System.Diagnostics.Stopwatch]::StartNew()

# -- Dynamic module discovery --
$targetModules = Get-Module -Name 'Microsoft.Graph.*' -ListAvailable |
    Select-Object -ExpandProperty Name -Unique |
    Where-Object { $_ -notmatch '\.internal$|\.private$' }

if (-not $IncludeBeta) {
    $targetModules = $targetModules | Where-Object { $_ -notmatch '\.Beta\.' }
}

Write-Host "`n[*] Discovered $($targetModules.Count) modules" -ForegroundColor Cyan
Write-Host "[*] Parallel threads: $ThrottleLimit" -ForegroundColor Cyan

# -- Pattern-based category mapping --
$categoryPatterns = [ordered]@{
    'DeviceManagement'       = 'Device Management'
    'Identity.Directory'     = 'Identity & Directory'
    'Identity.SignIns'       = 'Security'
    'Identity'               = 'Identity & Directory'
    'Users'                  = 'User Management'
    'Groups'                 = 'Groups'
    'Applications'           = 'Application Management'
    'Security'               = 'Security'
    'Mail'                   = 'Mail'
    'Calendar'               = 'Calendar'
    'Sites'                  = 'SharePoint'
    'Teams'                  = 'Teams'
    'Reports'                = 'Reports'
    'Compliance'             = 'Compliance'
    'Authentication'         = 'Authentication'
    'Files'                  = 'Files'
    'Notes'                  = 'Notes'
    'Planner'                = 'Planner'
    'Education'              = 'Education'
    'Bookings'               = 'Bookings'
    'CrossDeviceExperiences' = 'Cross-Device'
    'PersonalContacts'       = 'Contacts'
    'People'                 = 'People'
    'Search'                 = 'Search'
    'CloudCommunications'    = 'Communications'
}

# Flatten to array of pairs for use inside parallel blocks (ordered hashtable
# doesn't serialize into runspaces cleanly)
$categoryPairs = @($categoryPatterns.GetEnumerator() | ForEach-Object { @{ Key = $_.Key; Value = $_.Value } })

# -- Discover cmdlets --
Write-Host "[*] Scanning modules..." -ForegroundColor Cyan
$allCmdlets = [System.Collections.ArrayList]::new()
foreach ($mod in $targetModules) {
    $cmds = Get-Command -Module $mod -ErrorAction SilentlyContinue
    if ($cmds) {
        Write-Verbose "  $mod : $($cmds.Count) cmdlets"
        foreach ($c in $cmds) { $null = $allCmdlets.Add($c) }
    }
    else {
        Write-Verbose "  $mod : not installed, skipping"
    }
}

# Dedupe by name
$allCmdlets = [System.Collections.ArrayList]@($allCmdlets | Sort-Object Name -Unique)

# Skip -Count cmdlets
$allCmdlets = [System.Collections.ArrayList]@($allCmdlets | Where-Object { $_.Name -notmatch 'Count$' })

if ($MaxCmdlets -gt 0) {
    $allCmdlets = [System.Collections.ArrayList]@($allCmdlets | Select-Object -First $MaxCmdlets)
}

$total = $allCmdlets.Count
Write-Host "[*] Found $total cmdlets to process" -ForegroundColor Cyan
Write-Host "[*] Elapsed: $([math]::Round($timer.Elapsed.TotalSeconds))s`n" -ForegroundColor DarkGray

# -- Build serializable cmdlet info for parallel blocks --
# CmdletInfo objects don't serialize into parallel runspaces, so extract
# the fields we need into plain hashtables.
$cmdletInfoList = $allCmdlets | ForEach-Object {
    @{ Name = $_.Name; ModuleName = $_.ModuleName }
}

# -- Permission lookup (parallel, or skip) --
$permCache = @{}
if (-not $SkipPermissions) {
    Write-Host "[*] Building permission cache (parallel, $ThrottleLimit threads)..." -ForegroundColor Cyan
    $permCounter = [System.Threading.Interlocked]
    $permDone = [ref]0

    $permResults = $cmdletInfoList | ForEach-Object -Parallel {
        $cmdName = $_.Name
        $counter = $using:permDone
        $t = $using:total
        [System.Threading.Interlocked]::Increment($counter) | Out-Null
        try {
            $result = Find-MgGraphCommand -Command $cmdName -ErrorAction SilentlyContinue |
                Select-Object -First 1 -ExpandProperty Permissions |
                Select-Object -ExpandProperty Name -Unique
            if ($result) {
                [PSCustomObject]@{ Name = $cmdName; Permissions = @($result) }
            }
        }
        catch { }
    } -ThrottleLimit $ThrottleLimit

    foreach ($r in $permResults) {
        if ($r) { $permCache[$r.Name] = $r.Permissions }
    }
    Write-Host "[*] Cached permissions for $($permCache.Count) cmdlets" -ForegroundColor Green
}
else {
    Write-Host "[*] Skipping permission lookup (-SkipPermissions)" -ForegroundColor Yellow
}
Write-Host "[*] Elapsed: $([math]::Round($timer.Elapsed.TotalSeconds))s`n" -ForegroundColor DarkGray

# -- Extract metadata (parallel Get-Help) --
Write-Host "[*] Extracting metadata (parallel, $ThrottleLimit threads)..." -ForegroundColor Cyan
$metaDone = [ref]0

$cmdletData = $cmdletInfoList | ForEach-Object -Parallel {
    $cmdName = $_.Name
    $modName = $_.ModuleName
    $catPairs = $using:categoryPairs
    $counter = $using:metaDone
    $t = $using:total
    $done = [System.Threading.Interlocked]::Increment($counter)

    # Category from module name
    $suffix = $modName -replace '^Microsoft\.Graph\.(Beta\.)?', ''
    $category = 'General'
    foreach ($pair in $catPairs) {
        if ($suffix -like "$($pair.Key)*") {
            $category = $pair.Value
            break
        }
    }

    # API version
    $apiVersion = if ($modName -match '\.Beta\.') { 'beta' } else { 'v1.0' }

    # Verb
    $verb = ($cmdName -split '-')[0]

    # Get-Help for description, syntax, examples
    $help = $null
    try { $help = Get-Help $cmdName -Full -ErrorAction SilentlyContinue } catch { }

    $description = ''
    if ($help.description) {
        $description = ($help.description | Out-String).Trim()
        if ($description.Length -gt 300) {
            $description = $description.Substring(0, 297) + '...'
        }
    }

    $syntaxStr = ''
    if ($help.syntax) {
        $syntaxStr = ($help.syntax | Out-String).Trim()
        $lines = @($syntaxStr -split "`n" | Where-Object { $_.Trim() -ne '' })
        if ($lines.Count -gt 0) { $syntaxStr = $lines[0].Trim() }
    }
    if (-not $syntaxStr) { $syntaxStr = "$cmdName [parameters]" }

    $examples = @()
    if ($help.examples.example) {
        $examples = @($help.examples.example | ForEach-Object {
            ($_.code | Out-String).Trim()
        } | Where-Object { $_ -ne '' } | Select-Object -First 3)
    }

    [PSCustomObject]@{
        name        = $cmdName
        verb        = $verb
        category    = $category
        module      = $modName
        apiVersion  = $apiVersion
        description = if ($description) { $description } else { 'Microsoft Graph PowerShell cmdlet.' }
        syntax      = $syntaxStr
        examples    = $examples
    }
} -ThrottleLimit $ThrottleLimit

# Merge permissions from cache and sort
$cmdletData = @($cmdletData | Sort-Object name | ForEach-Object {
    $perms = if ($permCache.ContainsKey($_.name)) { $permCache[$_.name] } else { @() }
    [ordered]@{
        name        = $_.name
        verb        = $_.verb
        category    = $_.category
        module      = $_.module
        apiVersion  = $_.apiVersion
        description = $_.description
        syntax      = $_.syntax
        examples    = @($_.examples)
        permissions = @($perms)
    }
})

Write-Host "[*] Extracted $($cmdletData.Count) cmdlets" -ForegroundColor Green
Write-Host "[*] Elapsed: $([math]::Round($timer.Elapsed.TotalSeconds))s`n" -ForegroundColor DarkGray

# -- Ensure output directories exist --
$dataDir = Split-Path $OutputPath -Parent
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir -Force | Out-Null }

$modulesDir = Join-Path $dataDir 'modules'
if (-not (Test-Path $modulesDir)) { New-Item -ItemType Directory -Path $modulesDir -Force | Out-Null }

# -- Write flat cmdlets.json (backward compat) --
$cmdletData | ConvertTo-Json -Depth 5 -Compress | Set-Content -Path $OutputPath -Encoding UTF8
$flatSize = [math]::Round((Get-Item $OutputPath).Length / 1KB)
Write-Host "[*] Written flat file: $OutputPath (${flatSize}KB)" -ForegroundColor Green

# -- Write manifest.json (lightweight summary) --
$manifest = $cmdletData | ForEach-Object {
    [ordered]@{
        name            = $_.name
        verb            = $_.verb
        category        = $_.category
        module          = $_.module
        apiVersion      = $_.apiVersion
        description     = $_.description
        permissionCount = $_.permissions.Count
        hasExamples     = ($_.examples.Count -gt 0)
    }
}

$manifestPath = Join-Path $dataDir 'manifest.json'
$manifest | ConvertTo-Json -Depth 3 -Compress | Set-Content -Path $manifestPath -Encoding UTF8
$manifestSize = [math]::Round((Get-Item $manifestPath).Length / 1KB)
Write-Host "[*] Written manifest: $manifestPath ($($manifest.Count) entries, ${manifestSize}KB)" -ForegroundColor Green

# -- Write per-module detail files --
$moduleGroups = $cmdletData | Group-Object -Property { $_.module }
foreach ($group in $moduleGroups) {
    $moduleName = $group.Name
    $moduleDetail = [ordered]@{
        module  = $moduleName
        cmdlets = [ordered]@{}
    }
    foreach ($cmd in $group.Group) {
        $moduleDetail.cmdlets[$cmd.name] = [ordered]@{
            syntax      = $cmd.syntax
            examples    = $cmd.examples
            permissions = $cmd.permissions
        }
    }
    $moduleFileName = "$moduleName.json"
    $moduleFilePath = Join-Path $modulesDir $moduleFileName
    $moduleDetail | ConvertTo-Json -Depth 5 -Compress | Set-Content -Path $moduleFilePath -Encoding UTF8
    $fileSize = [math]::Round((Get-Item $moduleFilePath).Length / 1KB)
    Write-Verbose "  $moduleFileName ($($group.Count) cmdlets, ${fileSize}KB)"
}
Write-Host "[*] Written $($moduleGroups.Count) module detail files to $modulesDir" -ForegroundColor Green

$timer.Stop()
Write-Host "`n[*] Done in $([math]::Round($timer.Elapsed.TotalSeconds))s.`n" -ForegroundColor Cyan
