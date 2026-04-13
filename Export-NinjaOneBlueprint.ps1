<#
.SYNOPSIS
  Exports a NinjaOne "tenant blueprint" (policies, patching, alerts, scripts, roles, org structure, device metadata)
  with optional sanitization.

.NOTES
  - You run this locally.
  - Token is supplied via parameter or environment variable.
  - Output is JSON files + a ZIP bundle.
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$true)]
  [string]$BaseUrl,

  [Parameter(Mandatory=$false)]
  [string]$Token,

  [Parameter(Mandatory=$false)]
  [string]$OutputPath = (Join-Path $PWD "NinjaOne-Blueprint"),

  # Sanitization / privacy controls
  [switch]$Sanitize = $true,
  [switch]$HashOrgAndSiteNames = $false,
  [switch]$RedactDeviceIdentifiers = $true,
  [switch]$RedactPublicIPs = $true,
  [switch]$RedactUserEmails = $true,

  # If your tenant has strict rate limits, increase delay
  [int]$RequestDelayMs = 150
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-EnvToken {
  $envToken = $env:NINJAONE_TOKEN
  if ([string]::IsNullOrWhiteSpace($envToken)) { return $null }
  return $envToken
}

function New-Hash {
  param([Parameter(Mandatory=$true)][string]$Value)
  $bytes = [System.Text.Encoding]::UTF8.GetBytes($Value)
  $sha = [System.Security.Cryptography.SHA256]::Create()
  ($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString("x2") }) -join ''
}

function Invoke-NinjaApi {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory=$true)][string]$Method,
    [Parameter(Mandatory=$true)][string]$Path,
    [hashtable]$Query,
    $Body
  )

  Start-Sleep -Milliseconds $RequestDelayMs

  $uriBuilder = [System.UriBuilder]::new(($BaseUrl.TrimEnd('/') + $Path))
  if ($Query) {
    $qs = $Query.GetEnumerator() | ForEach-Object {
      "{0}={1}" -f [System.Web.HttpUtility]::UrlEncode($_.Key),
                [System.Web.HttpUtility]::UrlEncode([string]$_.Value)
    }
    $uriBuilder.Query = ($qs -join '&')
  }

  $headers = @{
  "Accept"    = "application/json"
  "X-API-KEY" = $script:AuthToken
}

  $params = @{
    Method      = $Method
    Uri         = $uriBuilder.Uri.AbsoluteUri
    Headers     = $headers
  }

  if ($null -ne $Body) {
    $params["ContentType"] = "application/json"
    $params["Body"] = ($Body | ConvertTo-Json -Depth 10)
  }

  try {
    return Invoke-RestMethod @params
  }
  catch {
    throw "API call failed: $Method $($uriBuilder.Uri.AbsoluteUri)`n$($_.Exception.Message)"
  }
}

function Get-AllPages {
  <#
    Best-effort pagination handler.
    NinjaOne endpoints vary; some support page/pageSize, some use limit/offset.
    We try page/pageSize first then fall back to limit/offset.
  #>
  param(
    [Parameter(Mandatory=$true)][string]$Path,
    [hashtable]$BaseQuery = @{},
    [int]$PageSize = 200
  )

  $results = New-Object System.Collections.Generic.List[object]

  # Attempt page/pageSize
  $page = 1
  while ($true) {
    $q = @{}
    $BaseQuery.GetEnumerator() | ForEach-Object { $q[$_.Key] = $_.Value }
    $q["page"] = $page
    $q["pageSize"] = $PageSize

    $resp = Invoke-NinjaApi -Method "GET" -Path $Path -Query $q
    if ($null -eq $resp) { break }

    # Many endpoints return arrays directly; some wrap in "data"
    $items =
      if ($resp.PSObject.Properties.Name -contains "data") { $resp.data }
      elseif ($resp -is [System.Array]) { $resp }
      else { @($resp) }

    if ($items.Count -eq 0) { break }

    $items | ForEach-Object { $results.Add($_) }

    # If response indicates last page, stop. Otherwise stop when less than pageSize.
    $isLast =
      ($resp.PSObject.Properties.Name -contains "totalPages" -and $page -ge [int]$resp.totalPages) -or
      ($items.Count -lt $PageSize)

    if ($isLast) { break }
    $page++
  }

  # If we got nothing, try limit/offset approach
  if ($results.Count -eq 0) {
    $offset = 0
    while ($true) {
      $q = @{}
      $BaseQuery.GetEnumerator() | ForEach-Object { $q[$_.Key] = $_.Value }
      $q["limit"] = $PageSize
      $q["offset"] = $offset

      $resp = Invoke-NinjaApi -Method "GET" -Path $Path -Query $q
      if ($null -eq $resp) { break }

      $items =
        if ($resp.PSObject.Properties.Name -contains "data") { $resp.data }
        elseif ($resp -is [System.Array]) { $resp }
        else { @($resp) }

      if ($items.Count -eq 0) { break }

      $items | ForEach-Object { $results.Add($_) }

      if ($items.Count -lt $PageSize) { break }
      $offset += $PageSize
    }
  }

  return $results
}

function Sanitize-Object {
  param([Parameter(Mandatory=$true)]$Obj)

  if (-not $Sanitize) { return $Obj }

  # Convert to JSON and back to get a mutable PSCustomObject graph
  $clone = $Obj | ConvertTo-Json -Depth 50 | ConvertFrom-Json

  # Helper to redact property if it exists
  function Redact-IfPresent($o, $propName, $replacement = "[REDACTED]") {
    if ($null -ne $o -and $o.PSObject.Properties.Name -contains $propName) {
      $o.$propName = $replacement
    }
  }

  # Recursive walk
  function Walk($node) {
    if ($null -eq $node) { return }
    if ($node -is [System.Array]) {
      foreach ($n in $node) { Walk $n }
      return
    }

    # If it's an object, handle redactions
    if ($node -is [pscustomobject]) {

      if ($HashOrgAndSiteNames) {
        if ($node.PSObject.Properties.Name -contains "name" -and $node.PSObject.Properties.Name -contains "id") {
          # Hash only when it looks like an org/site style object
          $node.name = ("HASH_" + (New-Hash -Value [string]$node.name).Substring(0,12))
        }
      }

      if ($RedactDeviceIdentifiers) {
        foreach ($p in @("serialNumber","serial","macAddress","mac","uuid","biosSerial","deviceId","agentId")) {
          Redact-IfPresent $node $p
        }
      }

      if ($RedactPublicIPs) {
        foreach ($p in @("publicIp","publicIP","externalIp","externalIP","wanIp","wanIP")) {
          Redact-IfPresent $node $p
        }
      }

      if ($RedactUserEmails) {
        foreach ($p in @("email","userEmail","primaryEmail","username")) {
          if ($node.PSObject.Properties.Name -contains $p) {
            $v = [string]$node.$p
            if ($v -match '@') { $node.$p = "[REDACTED_EMAIL]" }
          }
        }
      }

      # Walk child properties
      foreach ($prop in $node.PSObject.Properties) {
        Walk $prop.Value
      }
    }
  }

  Walk $clone
  return $clone
}

function Write-JsonFile {
  param(
    [Parameter(Mandatory=$true)][string]$Path,
    [Parameter(Mandatory=$true)]$Data
  )

  $dir = Split-Path -Parent $Path
  if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

  $Data | ConvertTo-Json -Depth 50 | Out-File -FilePath $Path -Encoding UTF8
}

# --- Auth ---
if ([string]::IsNullOrWhiteSpace($Token)) {
  $Token = Get-EnvToken
}
if ([string]::IsNullOrWhiteSpace($Token)) {
  throw "No token provided. Use -Token or set environment variable NINJAONE_TOKEN."
}
$script:AuthToken = $Token

# --- Output folder ---
$timestamp = (Get-Date).ToString("yyyyMMdd-HHmmss")
$root = Join-Path $OutputPath ("Blueprint-$timestamp")
New-Item -ItemType Directory -Path $root -Force | Out-Null

$manifest = [ordered]@{
  createdAt = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
  baseUrl   = $BaseUrl
  sanitize  = [bool]$Sanitize
  files     = @{}
}

Write-Host "Exporting NinjaOne blueprint to: $root" -ForegroundColor Cyan

# --- Define endpoints to attempt ---
# NOTE: Endpoint paths can vary by NinjaOne API version/tenant.
# The script will fail fast if a path is invalid; you can comment out the failing section and re-run,
# or tell me the error and I’ll tailor paths to your tenant.
$exports = @(
  @{ Name = "organizations"; Path = "/api/v2/organizations"; Paged = $true },
  @{ Name = "sites";         Path = "/api/v2/sites";         Paged = $true },
  @{ Name = "devices";       Path = "/api/v2/devices";       Paged = $true },
  @{ Name = "policies";      Path = "/api/v2/policies";      Paged = $true },
  @{ Name = "patchPolicies"; Path = "/api/v2/patch/policies";Paged = $true },
  @{ Name = "alerts";        Path = "/api/v2/alerts";        Paged = $true },
  @{ Name = "alertRules";    Path = "/api/v2/alert-rules";   Paged = $true },
  @{ Name = "scripts";       Path = "/api/v2/scripts";       Paged = $true },
  @{ Name = "roles";         Path = "/api/v2/roles";         Paged = $true },
  @{ Name = "users";         Path = "/api/v2/users";         Paged = $true },
  @{ Name = "customFields";  Path = "/api/v2/custom-fields"; Paged = $true }
)

foreach ($e in $exports) {
  $name = $e.Name
  $path = $e.Path
  Write-Host "Pulling: $name ($path)" -ForegroundColor Yellow

  try {
    $raw =
      if ($e.Paged) { Get-AllPages -Path $path }
      else { Invoke-NinjaApi -Method "GET" -Path $path }

    $data = Sanitize-Object -Obj $raw

    $file = Join-Path $root ("{0}.json" -f $name)
    Write-JsonFile -Path $file -Data $data

    $count =
      if ($data -is [System.Array]) { $data.Count }
      elseif ($data.PSObject.Properties.Name -contains "data") { @($data.data).Count }
      else { 1 }

    $manifest.files[$name] = [ordered]@{ path = (Split-Path $file -Leaf); count = $count }
    Write-Host "  -> Saved $count records" -ForegroundColor Green
  }
  catch {
    Write-Warning "Failed export '$name' at '$path': $($_.Exception.Message)"
    $manifest.files[$name] = [ordered]@{ path = $null; count = 0; error = $($_.Exception.Message) }
  }
}

# Write manifest
$manifestPath = Join-Path $root "manifest.json"
Write-JsonFile -Path $manifestPath -Data $manifest

# Zip it
$zipPath = Join-Path $OutputPath ("Blueprint-$timestamp.zip")
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path (Join-Path $root "*") -DestinationPath $zipPath -Force

Write-Host ""
Write-Host "DONE. Blueprint bundle created:" -ForegroundColor Cyan
Write-Host "  Folder: $root"
Write-Host "  ZIP:    $zipPath"
Write-Host ""
Write-Host "Tip: You can share the ZIP, or paste selected JSON sections here." -ForegroundColor Cyan