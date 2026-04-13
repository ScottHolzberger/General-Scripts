<#
.SYNOPSIS
  MSP-friendly NinjaOne Blueprint Export (OAuth Client Credentials) with patch compliance.

.DESCRIPTION
  - OAuth token: POST {BaseUrl}/ws/oauth/token (client_credentials) [1](https://app.ninjaone.com/apidocs-beta/authorization/flows/client-credentials-flow)[2](https://www.ninjaone.com/docs/integrations/how-to-set-up-api-oauth-token/)
  - API calls:  GET  {BaseUrl}/v2/... with Authorization: Bearer <token> [1](https://app.ninjaone.com/apidocs-beta/authorization/flows/client-credentials-flow)[2](https://www.ninjaone.com/docs/integrations/how-to-set-up-api-oauth-token/)

  MSP Recommendations implemented:
  - Always export patch compliance via:
      /v2/queries/os-patches
      /v2/queries/software-patches  [3](https://community.postman.com/t/missing-header-help-please/36866)
  - DO NOT export /v2/queries/os-patch-installs by default (it can be extremely slow/large).
    Enable only when needed using -IncludeHeavyOsPatchInstalls with hard caps.

  Reliability:
  - Null-safe and shape-safe (arrays vs wrapped objects)
  - Per-endpoint sanitization fail-safe (won’t kill export)
  - Cursor/page/offset paging strategies
  - Optional HTTP timeout (where supported)

.IMPORTANT
  Do not hardcode secrets. Pass via env vars or parameters.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]$BaseUrl,

  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]$ClientId,

  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()]
  [string]$ClientSecret,

  [Parameter(Mandatory = $false)]
  [ValidateNotNullOrEmpty()]
  [string]$Scope = "monitoring",

  [Parameter(Mandatory = $false)]
  [ValidateNotNullOrEmpty()]
  [string]$OutputPath = (Join-Path $PWD "NinjaOne-Blueprint"),

  # MSP defaults: sanitise + hash org/location names
  [switch]$Sanitize = $true,
  [switch]$HashOrgAndLocationNames = $true,
  [switch]$RedactDeviceIdentifiers = $true,
  [switch]$RedactPublicIPs = $true,
  [switch]$RedactUserEmails = $true,

  # Endpoint export names to skip sanitization (raw export)
  [string[]]$SkipSanitizeExports = @("deviceCustomFields"),

  # Performance controls
  [int]$RequestDelayMs = 150,
  [int]$PageSize = 500,
  [int]$MaxPages = 5000,

  # HTTP timeout (seconds) for Invoke-RestMethod if supported by your PowerShell version.
  [int]$HttpTimeoutSec = 180,

  # Patch compliance controls (MSP recommended)
  [switch]$IncludePatchCompliance = $true,

  # MSP recommended: compliance (missing/pending/failed) — fast and actionable
  [switch]$IncludeOsPatches = $true,
  [switch]$IncludeSoftwarePatches = $true,

  # Optional: patch install history (can be large)
  [switch]$IncludeSoftwarePatchInstalls = $false,

  # HEAVY (often causes long runtimes). Disabled by default.
  [switch]$IncludeHeavyOsPatchInstalls = $false,

  # Patch history window for install history endpoints (days back)
  [ValidateRange(1, 3650)]
  [int]$PatchHistoryDays = 10,

  # Hard caps for heavy endpoints
  [ValidateRange(1, 5000)]
  [int]$HeavyMaxPages = 10
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# -----------------------------
# In-memory token cache
# -----------------------------
$script:TokenCache = $null

function Get-UtcNow { (Get-Date).ToUniversalTime() }

function ConvertTo-UnixEpoch {
  param([Parameter(Mandatory=$true)][datetime]$DateTime)
  $utc = $DateTime.ToUniversalTime()
  $epoch = [datetime]'1970-01-01T00:00:00Z'
  [math]::Floor(($utc - $epoch).TotalSeconds)
}

function New-Hash {
  param([Parameter(Mandatory = $true)][string]$Value)
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($Value)
  $sha = [System.Security.Cryptography.SHA256]::Create()
  ($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString("x2") }) -join ''
}

function Safe-FileName {
  param([Parameter(Mandatory=$true)][string]$Value)
  $invalid = [System.IO.Path]::GetInvalidFileNameChars()
  foreach ($c in $invalid) { $Value = $Value.Replace($c, '_') }
  return $Value
}

function Supports-TimeoutSec {
  try {
    return (Get-Command Invoke-RestMethod).Parameters.ContainsKey("TimeoutSec")
  } catch { return $false }
}

function Invoke-RestMethodSafe {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string]$Method,
    [Parameter(Mandatory=$true)][string]$Uri,
    [Parameter(Mandatory=$false)][hashtable]$Headers,
    [Parameter(Mandatory=$false)][string]$ContentType,
    [Parameter(Mandatory=$false)]$Body
  )

  $params = @{
    Method  = $Method
    Uri     = $Uri
    Headers = $Headers
  }
  if ($ContentType) { $params.ContentType = $ContentType }
  if ($null -ne $Body) { $params.Body = $Body }

  if (Supports-TimeoutSec) {
    $params.TimeoutSec = $HttpTimeoutSec
  }

  return Invoke-RestMethod @params
}

function Get-NinjaOneOAuthToken {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)][string]$BaseUrl,
    [Parameter(Mandatory = $true)][string]$ClientId,
    [Parameter(Mandatory = $true)][string]$ClientSecret,
    [Parameter(Mandatory = $true)][string]$Scope
  )

  if ($script:TokenCache -and $script:TokenCache.ExpiresAtUtc) {
    $now = Get-UtcNow
    if ($now -lt $script:TokenCache.ExpiresAtUtc.AddSeconds(-60)) { return $script:TokenCache }
  }

  $tokenUri = ($BaseUrl.TrimEnd("/") + "/ws/oauth/token")

  $body = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = $Scope
  }

  Start-Sleep -Milliseconds $RequestDelayMs

  $resp = Invoke-RestMethodSafe -Method "POST" -Uri $tokenUri `
    -ContentType "application/x-www-form-urlencoded" -Body $body

  if (-not $resp.access_token) {
    throw "Token response did not include access_token. Raw: $($resp | ConvertTo-Json -Depth 10)"
  }

  $expiresIn = 3600
  if ($resp.expires_in) { $expiresIn = [int]$resp.expires_in }

  $script:TokenCache = [pscustomobject]@{
    AccessToken   = [string]$resp.access_token
    TokenType     = if ($resp.token_type) { [string]$resp.token_type } else { "Bearer" }
    Scope         = if ($resp.scope) { [string]$resp.scope } else { $Scope }
    ExpiresIn     = $expiresIn
    ObtainedAtUtc = Get-UtcNow
    ExpiresAtUtc  = (Get-UtcNow).AddSeconds($expiresIn)
  }

  return $script:TokenCache
}

function Invoke-NinjaApi {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][ValidateSet("GET","POST","PUT","PATCH","DELETE")]
    [string]$Method,

    [Parameter(Mandatory=$true)][string]$Path,

    [Parameter(Mandatory=$false)][hashtable]$Query,

    [Parameter(Mandatory=$false)]$Body
  )

  Start-Sleep -Milliseconds $RequestDelayMs

  $token = Get-NinjaOneOAuthToken -BaseUrl $BaseUrl -ClientId $ClientId -ClientSecret $ClientSecret -Scope $Scope

  $p = $Path.Trim()
  if (-not $p.StartsWith("/")) { $p = "/$p" }
  if (-not $p.StartsWith("/v2/")) { $p = "/v2" + $p }

  $uriBuilder = [System.UriBuilder]::new(($BaseUrl.TrimEnd("/") + $p))

  if ($Query) {
    $qs = $Query.GetEnumerator() | ForEach-Object {
      "{0}={1}" -f [System.Web.HttpUtility]::UrlEncode($_.Key),
                [System.Web.HttpUtility]::UrlEncode([string]$_.Value)
    }
    $uriBuilder.Query = ($qs -join '&')
  }

  $headers = @{
    "Accept"        = "application/json"
    "Authorization" = "Bearer $($token.AccessToken)"
  }

  if ($Method -eq "GET" -or $null -eq $Body) {
    return Invoke-RestMethodSafe -Method $Method -Uri $uriBuilder.Uri.AbsoluteUri -Headers $headers
  } else {
    return Invoke-RestMethodSafe -Method $Method -Uri $uriBuilder.Uri.AbsoluteUri -Headers $headers `
      -ContentType "application/json" -Body ($Body | ConvertTo-Json -Depth 12)
  }
}

# -----------------------------
# Response helpers (shape-safe)
# -----------------------------
function Get-ItemsFromResponse {
  param($resp)
  if ($null -eq $resp) { return @() }
  if ($resp -is [System.Array]) { return $resp }

  foreach ($k in @("data","items","results","records")) {
    if ($resp.PSObject.Properties.Name -contains $k) {
      $v = $resp.$k
      if ($null -eq $v) { return @() }
      if ($v -is [System.Array]) { return $v }
      return @($v)
    }
  }

  return @($resp)
}

function Get-NextCursorFromResponse {
  param($resp)
  if ($null -eq $resp) { return $null }
  foreach ($name in @("nextCursor","next_cursor","cursor","next","after")) {
    if ($resp.PSObject.Properties.Name -contains $name) {
      $v = $resp.$name
      if ($null -ne $v -and "$v" -ne "") { return "$v" }
    }
  }
  return $null
}

function Get-AllPages {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string]$Path,
    [hashtable]$BaseQuery = @{},
    [int]$LocalMaxPages = $MaxPages
  )

  $results = New-Object System.Collections.Generic.List[object]

  # Strategy A: page/pageSize
  $page = 1
  for ($i = 0; $i -lt $LocalMaxPages; $i++) {
    $q = @{}
    $BaseQuery.GetEnumerator() | ForEach-Object { $q[$_.Key] = $_.Value }
    $q["page"] = $page
    $q["pageSize"] = $PageSize

    $resp = Invoke-NinjaApi -Method "GET" -Path $Path -Query $q
    $items = @(Get-ItemsFromResponse $resp)
    if ($items.Count -eq 0) { break }
    foreach ($it in $items) { $results.Add($it) }
    if ($items.Count -lt $PageSize) { break }
    $page++
  }
  if ($results.Count -gt 0) { return $results }

  # Strategy B: limit/offset
  $offset = 0
  for ($i = 0; $i -lt $LocalMaxPages; $i++) {
    $q = @{}
    $BaseQuery.GetEnumerator() | ForEach-Object { $q[$_.Key] = $_.Value }
    $q["limit"] = $PageSize
    $q["offset"] = $offset

    $resp = Invoke-NinjaApi -Method "GET" -Path $Path -Query $q
    $items = @(Get-ItemsFromResponse $resp)
    if ($items.Count -eq 0) { break }
    foreach ($it in $items) { $results.Add($it) }
    if ($items.Count -lt $PageSize) { break }
    $offset += $PageSize
  }
  if ($results.Count -gt 0) { return $results }

  # Strategy C: cursor paging (common for query/report endpoints)
  $cursor = $null
  for ($i = 0; $i -lt $LocalMaxPages; $i++) {
    $q = @{}
    $BaseQuery.GetEnumerator() | ForEach-Object { $q[$_.Key] = $_.Value }
    $q["pageSize"] = $PageSize
    if ($null -ne $cursor) { $q["cursor"] = $cursor }

    $resp = Invoke-NinjaApi -Method "GET" -Path $Path -Query $q
    $items = @(Get-ItemsFromResponse $resp)
    if ($items.Count -eq 0) { break }
    foreach ($it in $items) { $results.Add($it) }

    $nextCursor = Get-NextCursorFromResponse $resp
    if ([string]::IsNullOrWhiteSpace($nextCursor) -or $nextCursor -eq $cursor) {
      if ($items.Count -lt $PageSize) { break }
      break
    }
    $cursor = $nextCursor
  }

  return $results
}

# -----------------------------
# Sanitization (per-endpoint fail-safe)
# -----------------------------
function Sanitize-Object {
  param($Obj)

  if ($null -eq $Obj) { return @() }
  if (-not $Sanitize) { return $Obj }

  $clone = $Obj | ConvertTo-Json -Depth 80 | ConvertFrom-Json

  function Walk($node) {
    if ($null -eq $node) { return }

    if ($node -is [System.Array]) {
      foreach ($n in $node) { Walk $n }
      return
    }

    if ($node -is [pscustomobject]) {

      if ($HashOrgAndLocationNames) {
        foreach ($propName in @("name","Name","label","Label")) {
          if ($node.PSObject.Properties.Name -contains $propName) {
            $n = [string]$node.$propName
            if (-not [string]::IsNullOrWhiteSpace($n)) {
              $node.$propName = "HASH_" + (New-Hash -Value $n).Substring(0, 12)
            }
          }
        }
      }

      if ($RedactDeviceIdentifiers) {
        foreach ($p in @("serialNumber","serial","macAddress","mac","uuid","biosSerial","deviceId","agentId","assetTag")) {
          if ($node.PSObject.Properties.Name -contains $p) { $node.$p = "[REDACTED]" }
        }
      }

      if ($RedactPublicIPs) {
        foreach ($p in @("publicIp","publicIP","externalIp","externalIP","wanIp","wanIP")) {
          if ($node.PSObject.Properties.Name -contains $p) { $node.$p = "[REDACTED]" }
        }
      }

      if ($RedactUserEmails) {
        foreach ($p in @("email","userEmail","primaryEmail","username","login")) {
          if ($node.PSObject.Properties.Name -contains $p) {
            $v = [string]$node.$p
            if ($v -match "@") { $node.$p = "[REDACTED_EMAIL]" }
          }
        }
      }

      foreach ($prop in $node.PSObject.Properties) { Walk $prop.Value }
    }
  }

  Walk $clone
  return $clone
}

function Write-JsonFile {
  param([Parameter(Mandatory=$true)][string]$Path, [Parameter(Mandatory=$true)]$Data)
  $dir = Split-Path -Parent $Path
  if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
  $Data | ConvertTo-Json -Depth 80 | Out-File -FilePath $Path -Encoding UTF8
}

function Get-RecordCount {
  param($Data)
  if ($null -eq $Data) { return 0 }
  if ($Data -is [System.Array]) { return $Data.Count }
  foreach ($k in @("data","items","results","records")) {
    if ($Data.PSObject.Properties.Name -contains $k) {
      $v = $Data.$k
      if ($null -eq $v) { return 0 }
      if ($v -is [System.Array]) { return $v.Count }
      return 1
    }
  }
  return 1
}

function Export-Endpoint {
  param(
    [Parameter(Mandatory=$true)][string]$Name,
    [Parameter(Mandatory=$true)][string]$Path,
    [Parameter(Mandatory=$true)][string]$OutDir,
    [hashtable]$Query = @{},
    [switch]$Paged = $true,
    [int]$LocalMaxPages = $MaxPages
  )

  Write-Host "Pulling: $Name ($Path)" -ForegroundColor Yellow
  $sw = [System.Diagnostics.Stopwatch]::StartNew()

  try {
    $raw = if ($Paged) { Get-AllPages -Path $Path -BaseQuery $Query -LocalMaxPages $LocalMaxPages } else { Invoke-NinjaApi -Method "GET" -Path $Path -Query $Query }
    if ($null -eq $raw) { $raw = @() }

    $data = $null
    if ($SkipSanitizeExports -contains $Name) {
      $data = $raw
    } else {
      try { $data = Sanitize-Object -Obj $raw }
      catch {
        $script:Manifest.warnings += "Sanitize failed for '$Name' :: $($_.Exception.Message) (exported raw instead)"
        $data = $raw
      }
    }

    $file = Join-Path $OutDir ("{0}.json" -f $Name)
    Write-JsonFile -Path $file -Data $data

    $count = Get-RecordCount -Data $data
    $sw.Stop()

    $script:Manifest.files[$Name] = [ordered]@{ path = (Split-Path $file -Leaf); count = $count }
    $script:Manifest.timings[$Name] = [ordered]@{ seconds = [math]::Round($sw.Elapsed.TotalSeconds, 2) }

    Write-Host "  -> Saved $count records" -ForegroundColor Green
    return $data
  }
  catch {
    $sw.Stop()
    $msg = $_.Exception.Message
    Write-Warning "Failed export '$Name' at '$Path': $msg"
    $script:Manifest.files[$Name] = [ordered]@{ path = $null; count = 0; error = $msg }
    $script:Manifest.warnings += "Export failed: $Name ($Path) :: $msg"
    $script:Manifest.timings[$Name] = [ordered]@{ seconds = [math]::Round($sw.Elapsed.TotalSeconds, 2); failed = $true }
    return $null
  }
}

function Build-PatchComplianceSummary {
  param(
    $OsPatches,
    $SoftwarePatches,
    $SoftwarePatchInstalls,
    $OsPatchInstalls
  )

  # Create a lightweight summary that’s stable even when payload shapes vary
  $summary = [ordered]@{
    generatedAtUtc = (Get-UtcNow).ToString("yyyy-MM-dd HH:mm:ss 'UTC'")
    windowsOs = [ordered]@{}
    thirdParty = [ordered]@{}
    notes = @()
  }

  $os = @($OsPatches)
  $sw = @($SoftwarePatches)
  $swInst = @($SoftwarePatchInstalls)
  $osInst = @($OsPatchInstalls)

  $summary.windowsOs.totalRows = $os.Count
  $summary.thirdParty.totalRows = $sw.Count

  # Count common fields if present
  function Group-Counts($rows, $field) {
    if (-not $rows -or $rows.Count -eq 0) { return @{} }
    $has = $false
    foreach ($r in $rows) { if ($r.PSObject.Properties.Name -contains $field) { $has = $true; break } }
    if (-not $has) { return @{} }

    $g = $rows | Group-Object -Property $field
    $o = [ordered]@{}
    foreach ($x in $g) { $o[$x.Name] = $x.Count }
    return $o
  }

  $summary.windowsOs.byStatus   = Group-Counts $os "status"
  $summary.windowsOs.bySeverity = Group-Counts $os "severity"
  $summary.windowsOs.byType     = Group-Counts $os "type"

  $summary.thirdParty.byStatus   = Group-Counts $sw "status"
  $summary.thirdParty.bySeverity = Group-Counts $sw "severity"
  $summary.thirdParty.byProduct  = Group-Counts $sw "productName"

  if ($swInst.Count -gt 0) {
    $summary.thirdParty.installHistoryRows = $swInst.Count
    $summary.thirdParty.installHistoryByStatus = Group-Counts $swInst "status"
  }

  if ($osInst.Count -gt 0) {
    $summary.windowsOs.installHistoryRows = $osInst.Count
    $summary.windowsOs.installHistoryByStatus = Group-Counts $osInst "status"
  } else {
    $summary.notes += "OS patch install history not included by default (heavy endpoint). Enable with -IncludeHeavyOsPatchInstalls if required."
  }

  return $summary
}

# -----------------------------
# Main
# -----------------------------
$timestamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
$root = Join-Path $OutputPath ("Blueprint-$timestamp")
New-Item -ItemType Directory -Path $root -Force | Out-Null

$patchRoot = Join-Path $root "patch"
New-Item -ItemType Directory -Path $patchRoot -Force | Out-Null

$script:Manifest = [ordered]@{
  createdAtUtc = (Get-UtcNow).ToString("yyyy-MM-dd HH:mm:ss 'UTC'")
  baseUrl      = $BaseUrl
  auth         = "oauth_client_credentials"
  scope        = $Scope
  sanitize     = [bool]$Sanitize
  files        = [ordered]@{}
  warnings     = @()
  timings      = [ordered]@{}
  patch        = [ordered]@{
    includePatchCompliance = [bool]$IncludePatchCompliance
    includeOsPatches = [bool]$IncludeOsPatches
    includeSoftwarePatches = [bool]$IncludeSoftwarePatches
    includeSoftwarePatchInstalls = [bool]$IncludeSoftwarePatchInstalls
    includeHeavyOsPatchInstalls = [bool]$IncludeHeavyOsPatchInstalls
    patchHistoryDays = $PatchHistoryDays
    heavyMaxPages = $HeavyMaxPages
  }
}

Write-Host "Exporting NinjaOne blueprint to: $root" -ForegroundColor Cyan

# ---- Core MSP blueprint exports
$null = Export-Endpoint -Name "organizations"         -Path "/v2/organizations"          -OutDir $root
$null = Export-Endpoint -Name "organizationsDetailed" -Path "/v2/organizations-detailed" -OutDir $root
$null = Export-Endpoint -Name "locations"             -Path "/v2/locations"              -OutDir $root
$devices = Export-Endpoint -Name "devices"            -Path "/v2/devices"                -OutDir $root
$null = Export-Endpoint -Name "devicesDetailed"       -Path "/v2/devices-detailed"       -OutDir $root
$null = Export-Endpoint -Name "policies"              -Path "/v2/policies"               -OutDir $root
$null = Export-Endpoint -Name "alerts"                -Path "/v2/alerts"                 -OutDir $root
$null = Export-Endpoint -Name "automationScripts"     -Path "/v2/automation/scripts"     -OutDir $root
$null = Export-Endpoint -Name "deviceCustomFields"    -Path "/v2/device-custom-fields"   -OutDir $root
$null = Export-Endpoint -Name "users"                 -Path "/v2/users"                  -OutDir $root
$null = Export-Endpoint -Name "userRoles"             -Path "/v2/user/roles"             -OutDir $root
$null = Export-Endpoint -Name "contacts"              -Path "/v2/contacts"               -OutDir $root
$null = Export-Endpoint -Name "notificationChannels"  -Path "/v2/notification-channels"  -OutDir $root
$null = Export-Endpoint -Name "groups"                -Path "/v2/groups"                 -OutDir $root

# ---- Patch compliance exports (MSP-friendly defaults)
$osPatches = @()
$swPatches = @()
$osInstalls = @()
$swInstalls = @()

if ($IncludePatchCompliance) {

  if ($IncludeOsPatches) {
    # Compliance view (missing/pending/failed/rejected) [3](https://community.postman.com/t/missing-header-help-please/36866)
    $osPatches = Export-Endpoint -Name "queries_os_patches" -Path "/v2/queries/os-patches" -OutDir $patchRoot
    $script:Manifest.files["queries_os_patches"].path = "patch\queries_os_patches.json"
  }

  if ($IncludeSoftwarePatches) {
    # Compliance view (pending/failed/rejected software patches) [3](https://community.postman.com/t/missing-header-help-please/36866)
    $swPatches = Export-Endpoint -Name "queries_software_patches" -Path "/v2/queries/software-patches" -OutDir $patchRoot
    $script:Manifest.files["queries_software_patches"].path = "patch\queries_software_patches.json"
  }

  $installedAfterEpoch  = ConvertTo-UnixEpoch -DateTime (Get-Date).AddDays(-1 * $PatchHistoryDays)
  $installedBeforeEpoch = ConvertTo-UnixEpoch -DateTime (Get-Date)

  if ($IncludeSoftwarePatchInstalls) {
    # Optional: software patch install history (can still be sizable)
    $swInstalls = Export-Endpoint -Name "queries_software_patch_installs" `
      -Path "/v2/queries/software-patch-installs" -OutDir $patchRoot `
      -Query @{ installedAfter = $installedAfterEpoch; installedBefore = $installedBeforeEpoch }
    if ($script:Manifest.files.Contains("queries_software_patch_installs")) {
      $script:Manifest.files["queries_software_patch_installs"].path = "patch\queries_software_patch_installs.json"
    }
  }

  if ($IncludeHeavyOsPatchInstalls) {
    # HEAVY endpoint: hard cap paging to avoid "hang" in MSP environments
    $osInstalls = Export-Endpoint -Name "queries_os_patch_installs" `
      -Path "/v2/queries/os-patch-installs" -OutDir $patchRoot `
      -Query @{ installedAfter = $installedAfterEpoch; installedBefore = $installedBeforeEpoch } `
      -LocalMaxPages $HeavyMaxPages
    if ($script:Manifest.files.Contains("queries_os_patch_installs")) {
      $script:Manifest.files["queries_os_patch_installs"].path = "patch\queries_os_patch_installs.json"
    }
  } else {
    $script:Manifest.warnings += "Skipped /v2/queries/os-patch-installs by default (heavy). Enable with -IncludeHeavyOsPatchInstalls if required."
  }

  # Build a quick summary for MSP reporting
  $summary = Build-PatchComplianceSummary -OsPatches $osPatches -SoftwarePatches $swPatches -SoftwarePatchInstalls $swInstalls -OsPatchInstalls $osInstalls
  $summaryFile = Join-Path $patchRoot "patch_compliance_summary.json"
  Write-JsonFile -Path $summaryFile -Data $summary
  $script:Manifest.files["patch_compliance_summary"] = [ordered]@{ path = "patch\patch_compliance_summary.json"; count = 1 }
}

# ---- Write manifest
$manifestPath = Join-Path $root "manifest.json"
Write-JsonFile -Path $manifestPath -Data $script:Manifest

# ---- Zip output
$zipPath = Join-Path $OutputPath ("Blueprint-$timestamp.zip")
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path (Join-Path $root "*") -DestinationPath $zipPath -Force

Write-Host ""
Write-Host "DONE. Blueprint bundle created:" -ForegroundColor Cyan
Write-Host "  Folder: $root"
Write-Host "  ZIP:    $zipPath"
Write-Host ""
Write-Host "MSP Patch Compliance: patch\queries_os_patches.json + patch\queries_software_patches.json + patch\patch_compliance_summary.json" -ForegroundColor Cyan
Write-Host "Heavy OS patch install history is OFF by default (enable with -IncludeHeavyOsPatchInstalls -HeavyMaxPages <n>)." -ForegroundColor Cyan