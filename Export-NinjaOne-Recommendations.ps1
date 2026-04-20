
<#
Export-NinjaOne-Recommendations.ps1

Exports NinjaOne datasets required to make recommendations for:
  1) Patch management (security + minimize user impact)
  2) Monitoring & alerting noise reduction
  3) Automation improvements

Auth:
  OAuth2 client credentials. Token endpoint differs by region/instance; script tries:
    {BaseUrl}/oauth/token
    {BaseUrl}/ws/oauth/token

API:
  NinjaOne Public API v2 endpoints under:
    {BaseUrl}/api/v2/...

Outputs:
  - Raw JSON exports in OutDir
  - Summary CSVs:
      patch_device_policy_summary.csv
      alerts_noise_summary.csv
      automation_inventory_summary.csv
      export_manifest.json
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$false)]
  [string]$BaseUrl = "https://oc.ninjarmm.com",

  [Parameter(Mandatory=$false)]
  [string]$ClientId = $env:NINJAONE_CLIENT_ID,

  [Parameter(Mandatory=$false)]
  [SecureString]$ClientSecret,

  [Parameter(Mandatory=$false)]
  [ValidateSet("monitoring","management","control","monitoring management","monitoring management control")]
  [string]$Scope = "monitoring",

  [Parameter(Mandatory=$false)]
  [string]$OutDir = (Join-Path (Get-Location) ("ninjaone-export-" + (Get-Date -Format "yyyyMMdd-HHmmss"))),

  [Parameter(Mandatory=$false)]
  [int]$PageSize = 250,

  [Parameter(Mandatory=$false)]
  [switch]$Sanitize,

  [Parameter(Mandatory=$false)]
  [switch]$IncludeDeviceOsPatches,

  [Parameter(Mandatory=$false)]
  [switch]$IncludeDeviceOsPatchInstalls,

  [Parameter(Mandatory=$false)]
  [int]$PatchHistoryDays = 30,

  [Parameter(Mandatory=$false)]
  [switch]$IncludeHeavyOsPatchInstalls,

  [Parameter(Mandatory=$false)]
  [int]$HeavyMaxPages = 10
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# -----------------------------
# Helpers
# -----------------------------
function New-Directory([string]$Path) {
  if (-not (Test-Path -LiteralPath $Path)) {
    New-Item -ItemType Directory -Path $Path | Out-Null
  }
}

function ConvertFrom-SecureStringToPlainText([SecureString]$Secure) {
  $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
  try { [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
  finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

function Sanitize-Object([object]$obj) {
  $blocked = @("secret","password","token","apikey","api_key","clientsecret","client_secret","access_token","refresh_token","privatekey","private_key","key")
  if ($null -eq $obj) { return $null }

  if ($obj -is [System.Collections.IDictionary]) {
    $clone = @{}
    foreach ($k in $obj.Keys) {
      $keyName = [string]$k
      if ($blocked -contains ($keyName.ToLowerInvariant())) { continue }
      $clone[$k] = Sanitize-Object $obj[$k]
    }
    return $clone
  }

  if ($obj -is [System.Collections.IEnumerable] -and -not ($obj -is [string])) {
    $list = @()
    foreach ($item in $obj) { $list += (Sanitize-Object $item) }
    return $list
  }

  if ($obj -is [psobject]) {
    $pso = [pscustomobject]@{}
    foreach ($p in $obj.PSObject.Properties) {
      $name = $p.Name
      if ($blocked -contains ($name.ToLowerInvariant())) { continue }
      $pso | Add-Member -NotePropertyName $name -NotePropertyValue (Sanitize-Object $p.Value) -Force
    }
    return $pso
  }

  return $obj
}

function Save-Json([string]$Path, [object]$Data, [switch]$DoSanitize) {
  $payload = if ($DoSanitize) { Sanitize-Object $Data } else { $Data }
  $json = $payload | ConvertTo-Json -Depth 50
  $dir = Split-Path -Parent $Path
  New-Directory $dir
  [IO.File]::WriteAllText($Path, $json, [Text.Encoding]::UTF8)
}

function Save-Csv([string]$Path, [object[]]$Rows) {
  $dir = Split-Path -Parent $Path
  New-Directory $dir
  $Rows | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $Path
}

function Build-QueryString([hashtable]$Query) {
  if (-not $Query -or $Query.Count -eq 0) { return "" }

  $pairs = foreach ($kv in $Query.GetEnumerator()) {
    $k = [uri]::EscapeDataString([string]$kv.Key)
    $v = [uri]::EscapeDataString([string]$kv.Value)
    "$k=$v"
  }

  return ($pairs -join "&")
}

function Invoke-NinjaRequest {
  param(
    [Parameter(Mandatory=$true)][ValidateSet("GET","POST")][string]$Method,
    [Parameter(Mandatory=$true)][string]$Url,
    [Parameter(Mandatory=$false)][hashtable]$Headers,
    [Parameter(Mandatory=$false)][hashtable]$Body,
    [Parameter(Mandatory=$false)][string]$ContentType
  )

  if ($Method -eq "GET") {
    return Invoke-RestMethod -Method Get -Uri $Url -Headers $Headers -ContentType "application/json"
  } else {
    $ct = if ($ContentType) { $ContentType } else { "application/x-www-form-urlencoded" }
    return Invoke-RestMethod -Method Post -Uri $Url -Headers $Headers -Body $Body -ContentType $ct
  }
}

function Get-NinjaAccessToken {
  param(
    [Parameter(Mandatory=$true)][string]$BaseUrl,
    [Parameter(Mandatory=$true)][string]$ClientId,
    [Parameter(Mandatory=$true)][SecureString]$ClientSecret,
    [Parameter(Mandatory=$true)][string]$Scope
  )

  $secretPlain = ConvertFrom-SecureStringToPlainText $ClientSecret

  $tokenBody = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $secretPlain
    scope         = $Scope
  }

  $candidates = @(
    "$BaseUrl/oauth/token",
    "$BaseUrl/ws/oauth/token"
  )

  $lastErr = $null
  foreach ($tokenUrl in $candidates) {
    try {
      $resp = Invoke-NinjaRequest -Method POST -Url $tokenUrl -Body $tokenBody -ContentType "application/x-www-form-urlencoded"
      if ($resp.access_token) { return $resp.access_token }
    } catch {
      $lastErr = $_
    }
  }

  throw "Failed to obtain OAuth token from known token endpoints. Last error: $($lastErr.Exception.Message)"
}

function Build-ApiUrl([string]$BaseUrl, [string]$PathAndQuery) {
  $apiBase = "$BaseUrl/api"
  if ($PathAndQuery.StartsWith("/")) { return "$apiBase$PathAndQuery" }
  return "$apiBase/$PathAndQuery"
}

function Invoke-NinjaGet([string]$BaseUrl, [string]$AccessToken, [string]$PathAndQuery) {
  $url = Build-ApiUrl -BaseUrl $BaseUrl -PathAndQuery $PathAndQuery
  $headers = @{ Authorization = "Bearer $AccessToken"; Accept = "application/json" }
  return Invoke-NinjaRequest -Method GET -Url $url -Headers $headers
}

# Safe property access under StrictMode
function Has-Prop([object]$Obj, [string]$Name) {
  if ($null -eq $Obj) { return $false }
  return $null -ne ($Obj.PSObject.Properties | Where-Object { $_.Name -eq $Name } | Select-Object -First 1)
}

function Get-FirstPropValue([object]$Obj, [string[]]$Names, [object]$Default = $null) {
  foreach ($n in $Names) {
    if (Has-Prop $Obj $n) {
      $v = $Obj.$n
      if ($null -ne $v -and [string]$v -ne "") { return $v }
    }
  }
  return $Default
}

# -----------------------------
# Validate inputs
# -----------------------------
if ([string]::IsNullOrWhiteSpace($ClientId)) {
  throw "ClientId not provided. Set -ClientId or environment variable NINJAONE_CLIENT_ID."
}

if (-not $ClientSecret) {
  if ($env:NINJAONE_CLIENT_SECRET) {
    $ClientSecret = ConvertTo-SecureString -String $env:NINJAONE_CLIENT_SECRET -AsPlainText -Force
  } else {
    $ClientSecret = Read-Host -AsSecureString "Enter NinjaOne OAuth Client Secret"
  }
}

New-Directory $OutDir

# -----------------------------
# Auth
# -----------------------------
Write-Host "Authenticating to NinjaOne..."
$accessToken = Get-NinjaAccessToken -BaseUrl $BaseUrl -ClientId $ClientId -ClientSecret $ClientSecret -Scope $Scope

# -----------------------------
# Manifest
# -----------------------------
$manifest = [ordered]@{
  createdAtUtc = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd HH:mm:ss 'UTC'")
  baseUrl      = $BaseUrl
  auth         = "oauth_client_credentials"
  scope        = $Scope
  sanitize     = [bool]$Sanitize
  files        = [ordered]@{}
  warnings     = @()
  patch        = [ordered]@{
    includeDeviceOsPatches       = [bool]$IncludeDeviceOsPatches
    includeDeviceOsPatchInstalls = [bool]$IncludeDeviceOsPatchInstalls
    includeHeavyOsPatchInstalls  = [bool]$IncludeHeavyOsPatchInstalls
    patchHistoryDays             = $PatchHistoryDays
    heavyMaxPages                = $HeavyMaxPages
  }
}

function Export-Endpoint([string]$Name, [string]$PathAndQuery, [string]$RelativeFile) {
  Write-Host ("Exporting {0}..." -f $Name)
  $start = Get-Date
  try {
    $data = Invoke-NinjaGet -BaseUrl $BaseUrl -AccessToken $accessToken -PathAndQuery $PathAndQuery
    $outPath = Join-Path $OutDir $RelativeFile
    Save-Json -Path $outPath -Data $data -DoSanitize:$Sanitize

    $count = 1
    if ($data -is [System.Collections.IEnumerable] -and -not ($data -is [string])) { $count = @($data).Count }

    $manifest.files[$Name] = [ordered]@{
      path    = $RelativeFile
      count   = $count
      seconds = [math]::Round(((Get-Date) - $start).TotalSeconds, 2)
    }

    return $data
  } catch {
    $manifest.files[$Name] = [ordered]@{
      path    = $null
      count   = 0
      failed  = $true
      error   = $_.Exception.Message
      seconds = [math]::Round(((Get-Date) - $start).TotalSeconds, 2)
    }
    $manifest.warnings += ("Export failed: {0} ({1}) :: {2}" -f $Name, $PathAndQuery, $_.Exception.Message)
    return $null
  }
}

# -----------------------------
# Export core datasets
# -----------------------------
$organizations         = Export-Endpoint -Name "organizations"         -PathAndQuery "/v2/organizations"          -RelativeFile "organizations.json"
$organizationsDetailed = Export-Endpoint -Name "organizationsDetailed" -PathAndQuery "/v2/organizations-detailed" -RelativeFile "organizationsDetailed.json"
$locations             = Export-Endpoint -Name "locations"             -PathAndQuery "/v2/locations"              -RelativeFile "locations.json"
$devices               = Export-Endpoint -Name "devices"               -PathAndQuery "/v2/devices"                -RelativeFile "devices.json"
$devicesDetailed       = Export-Endpoint -Name "devicesDetailed"       -PathAndQuery "/v2/devices-detailed"       -RelativeFile "devicesDetailed.json"
$policies              = Export-Endpoint -Name "policies"              -PathAndQuery "/v2/policies"               -RelativeFile "policies.json"
$alerts                = Export-Endpoint -Name "alerts"                -PathAndQuery "/v2/alerts"                 -RelativeFile "alerts.json"
$automationScripts     = Export-Endpoint -Name "automationScripts"     -PathAndQuery "/v2/automation/scripts"     -RelativeFile "automationScripts.json"
$deviceCustomFields    = Export-Endpoint -Name "deviceCustomFields"    -PathAndQuery "/v2/device-custom-fields"   -RelativeFile "deviceCustomFields.json"
$users                 = Export-Endpoint -Name "users"                 -PathAndQuery "/v2/users"                  -RelativeFile "users.json"
$userRoles             = Export-Endpoint -Name "userRoles"             -PathAndQuery "/v2/user/roles"             -RelativeFile "userRoles.json"
$notificationChannels  = Export-Endpoint -Name "notificationChannels"  -PathAndQuery "/v2/notification-channels"  -RelativeFile "notificationChannels.json"
$notificationEnabled   = Export-Endpoint -Name "notificationChannelsEnabled" -PathAndQuery "/v2/notification-channels/enabled" -RelativeFile "notificationChannelsEnabled.json"
$groups                = Export-Endpoint -Name "groups"                -PathAndQuery "/v2/groups"                 -RelativeFile "groups.json"
$tasks                 = Export-Endpoint -Name "tasks"                 -PathAndQuery "/v2/tasks"                  -RelativeFile "tasks.json"

# -----------------------------
# Optional patch datasets
# -----------------------------
if ($IncludeDeviceOsPatches -and $devices) {
  $patchDir = Join-Path $OutDir "patch"
  New-Directory $patchDir
  $patchRows = @()

  $deviceList = @($devices)
  $i = 0

  foreach ($d in $deviceList) {
    $i++
    $deviceId = Get-FirstPropValue $d @("id") $null
    if (-not $deviceId) { continue }

    Write-Host ("[{0}/{1}] OS patches for device {2}..." -f $i, $deviceList.Count, $deviceId)
    try {
      $pending = Invoke-NinjaGet -BaseUrl $BaseUrl -AccessToken $accessToken -PathAndQuery ("/v2/device/{0}/os-patches" -f $deviceId)
      Save-Json -Path (Join-Path $patchDir ("device_{0}_os_patches.json" -f $deviceId)) -Data $pending -DoSanitize:$Sanitize

      $pendingCount = 0
      if ($pending -is [System.Collections.IEnumerable] -and -not ($pending -is [string])) { $pendingCount = @($pending).Count }

      $patchRows += [pscustomobject]@{
        deviceId     = $deviceId
        pendingCount = $pendingCount
      }
    } catch {
      $manifest.warnings += ("OS patches failed for device {0} :: {1}" -f $deviceId, $_.Exception.Message)
    }
  }

  if ($patchRows.Count -gt 0) {
    Save-Csv -Path (Join-Path $patchDir "device_os_patches_summary.csv") -Rows $patchRows
  }
}

if ($IncludeDeviceOsPatchInstalls -and $devices) {
  $patchDir = Join-Path $OutDir "patch"
  New-Directory $patchDir

  $since = (Get-Date).AddDays(-1 * $PatchHistoryDays).ToString("yyyy-MM-dd")
  $deviceList = @($devices)
  $i = 0

  foreach ($d in $deviceList) {
    $i++
    $deviceId = Get-FirstPropValue $d @("id") $null
    if (-not $deviceId) { continue }

    Write-Host ("[{0}/{1}] OS patch installs for device {2} (since {3})..." -f $i, $deviceList.Count, $deviceId, $since)
    try {
      $q = Build-QueryString @{ installedAfter = $since }
      $path = ("/v2/device/{0}/os-patch-installs?{1}" -f $deviceId, $q)
      $installs = Invoke-NinjaGet -BaseUrl $BaseUrl -AccessToken $accessToken -PathAndQuery $path
      Save-Json -Path (Join-Path $patchDir ("device_{0}_os_patch_installs.json" -f $deviceId)) -Data $installs -DoSanitize:$Sanitize
    } catch {
      $manifest.warnings += ("OS patch installs failed for device {0} :: {1}" -f $deviceId, $_.Exception.Message)
    }
  }
}

if ($IncludeHeavyOsPatchInstalls) {
  $patchDir = Join-Path $OutDir "patch"
  New-Directory $patchDir

  Write-Host "Running heavy query export: /v2/queries/os-patch-installs ..."
  $all = @()
  $page = 0

  while ($page -lt $HeavyMaxPages) {
    $page++
    try {
      $q = Build-QueryString @{ pageSize = $PageSize; page = $page }
      $resp = Invoke-NinjaGet -BaseUrl $BaseUrl -AccessToken $accessToken -PathAndQuery ("/v2/queries/os-patch-installs?{0}" -f $q)

      if ($resp -is [System.Collections.IEnumerable] -and -not ($resp -is [string])) {
        $chunk = @($resp)
        $all += $chunk
        if ($chunk.Count -lt $PageSize) { break }
      } else {
        $all += $resp
        break
      }
    } catch {
      $manifest.warnings += ("Heavy query export failed on page {0}: /v2/queries/os-patch-installs :: {1}" -f $page, $_.Exception.Message)
      break
    }
  }

  if ($all.Count -gt 0) {
    Save-Json -Path (Join-Path $patchDir "queries_os_patch_installs.json") -Data $all -DoSanitize:$Sanitize
    $manifest.files["queries_os_patch_installs"] = [ordered]@{
      path  = "patch\queries_os_patch_installs.json"
      count = $all.Count
    }
  }
}

# -----------------------------
# Create recommendation-ready summaries
# -----------------------------
# 1) Patch/device-policy summary (safe for schema differences)
$policyMap = @{}
if ($policies) {
  foreach ($p in @($policies)) {
    $policyIdValue = Get-FirstPropValue $p @("id") $null
    if ($policyIdValue) { $policyMap[[string]$policyIdValue] = $p }
  }
}

$devicePolicyRows = @()
if ($devicesDetailed) {
  foreach ($d in @($devicesDetailed)) {

    $deviceId = Get-FirstPropValue $d @("id") $null
    if (-not $deviceId) { continue }

    # policyId might be "policyId" or nested in "policy"
    $policyIdValue = Get-FirstPropValue $d @("policyId") $null
    if (-not $policyIdValue -and (Has-Prop $d "policy") -and $d.policy) {
      $policyIdValue = Get-FirstPropValue $d.policy @("id") $null
    }

    $policyNameValue = $null
    if ($policyIdValue -and $policyMap.ContainsKey([string]$policyIdValue)) {
      $policyNameValue = Get-FirstPropValue $policyMap[[string]$policyIdValue] @("name") $null
    }

    # Device name varies; prefer displayName, then systemName, etc.
    $deviceNameValue = Get-FirstPropValue $d @("displayName","systemName","dnsName","netbiosName","hostname","name") ("device_" + $deviceId)

    # Class varies; API examples show nodeClass; some exports use class
    $deviceClassValue = Get-FirstPropValue $d @("class","nodeClass") $null

    # Seen/contact time varies; API examples show lastContact/lastUpdate
    $lastSeenValue = Get-FirstPropValue $d @("lastSeen","lastContact","lastUpdate") $null

    $orgIdValue = Get-FirstPropValue $d @("organizationId") $null
    $onlineValue = $null
    if (Has-Prop $d "online") { $onlineValue = $d.online }
    elseif (Has-Prop $d "offline") { $onlineValue = -not [bool]$d.offline }

    $devicePolicyRows += [pscustomobject]@{
      deviceId       = $deviceId
      deviceName     = $deviceNameValue
      organizationId = $orgIdValue
      deviceClass    = $deviceClassValue
      online         = $onlineValue
      lastSeen       = $lastSeenValue
      policyId       = $policyIdValue
      policyName     = $policyNameValue
    }
  }
}

if ($devicePolicyRows.Count -gt 0) {
  Save-Csv -Path (Join-Path $OutDir "patch_device_policy_summary.csv") -Rows $devicePolicyRows
}

# 2) Alert noise summary
$alertRows = @()
if ($alerts) {
  foreach ($a in @($alerts)) {
    $alertRows += [pscustomobject]@{
      alertId        = Get-FirstPropValue $a @("id") $null
      deviceId       = Get-FirstPropValue $a @("deviceId") $null
      organizationId = Get-FirstPropValue $a @("organizationId") $null
      severity       = Get-FirstPropValue $a @("severity") $null
      type           = Get-FirstPropValue $a @("type") $null
      createdAt      = Get-FirstPropValue $a @("createdAt") $null
      message        = Get-FirstPropValue $a @("message") $null
    }
  }

  $grouped = $alertRows | Group-Object -Property type, severity | Sort-Object Count -Descending
  $noiseSummary = @()
  foreach ($g in $grouped) {
    $parts = $g.Name -split ",\s*"
    $noiseSummary += [pscustomobject]@{
      type     = $parts[0]
      severity = if ($parts.Count -gt 1) { $parts[1] } else { $null }
      count    = $g.Count
    }
  }
  Save-Csv -Path (Join-Path $OutDir "alerts_noise_summary.csv") -Rows $noiseSummary
}

# 3) Automation inventory summary
$autoSummary = @()
if ($automationScripts) {
  $byLang = @($automationScripts) | Group-Object -Property language | Sort-Object Count -Descending
  foreach ($g in $byLang) {
    $autoSummary += [pscustomobject]@{
      dimension = "language"
      value     = $g.Name
      count     = $g.Count
    }
  }

  $byOS = @()
  foreach ($s in @($automationScripts)) {
    if (Has-Prop $s "operatingSystems" -and $s.operatingSystems) {
      foreach ($os in @($s.operatingSystems)) {
        $byOS += [pscustomobject]@{ os = $os }
      }
    }
  }
  $byOSGrouped = $byOS | Group-Object -Property os | Sort-Object Count -Descending
  foreach ($g in $byOSGrouped) {
    $autoSummary += [pscustomobject]@{
      dimension = "operatingSystem"
      value     = $g.Name
      count     = $g.Count
    }
  }

  Save-Csv -Path (Join-Path $OutDir "automation_inventory_summary.csv") -Rows $autoSummary
}

# -----------------------------
# Write manifest
# -----------------------------
Save-Json -Path (Join-Path $OutDir "export_manifest.json") -Data $manifest -DoSanitize:$false

Write-Host ""
Write-Host "DONE. Export written to: $OutDir"
Write-Host "Key summaries:"
Write-Host "  - patch_device_policy_summary.csv"
Write-Host "  - alerts_noise_summary.csv"
Write-Host "  - automation_inventory_summary.csv"
Write-Host "  - export_manifest.json"
