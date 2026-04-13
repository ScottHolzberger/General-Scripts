#requires -Version 5.1
<#
.SYNOPSIS
  Sync Mosyle-exported iOS devices JSON into HaloPSA/HaloITSM as "Mobile Device" assets.

.DESCRIPTION
  - Reads Mosyle JSON export (array of objects).
  - Connects to Halo via OAuth2 client_credentials (Bearer token).
  - Resolves the Asset Type ID for "Mobile Device" (strict match; errors if ambiguous).
  - Pulls ONLY assets for the specified CustomerId AND the resolved AssetTypeId.
  - For each Mosyle device (serial_number):
      * If no match -> create asset
      * If match -> update mapped fields if different
  - Deactivates (inactive=true) only those Halo assets that BOTH:
      * belong to the resolved AssetTypeId
      * belong to the CustomerId
      * AND do not exist in Mosyle export

  Field mapping (per your requirements):
    - Serial Number -> key_field (optionally inventory_number)
    - Manufacturer  -> manufacturer_name = "Apple"
    - Model         -> key_field2 = osversion
    - Notes         -> notes = device_name

  Safety guards included:
    - Strict asset type resolution: will not "pick first" if name doesn't match.
    - Deactivation guard: will NEVER deactivate an asset whose assettype_id != resolved AssetTypeId.

.NOTES
  Recommended save locations:
    - Script:  C:\Users\ScottHolzberger\OneDrive - ZaheZone\Scripts\MosyleToHaloMobileAssetsSync_v3.ps1
    - Halo API credential (DPAPI protected):  %APPDATA%\ZaheZone\Halo\halo_api_app.xml

#>

Set-StrictMode -Version Latest

# -----------------------------
# DPAPI credential storage (Halo API Application)
# -----------------------------
function Get-HaloCredentialPath {
    [CmdletBinding()]
    param([string]$Name = 'halo_api_app')

    $base = Join-Path $env:APPDATA 'ZaheZone\Halo'
    if (-not (Test-Path $base)) { New-Item -ItemType Directory -Path $base -Force | Out-Null }
    Join-Path $base "$Name.xml"
}

function Save-HaloApiAppCredential {
    <# Stores Halo API ClientID/ClientSecret as PSCredential (UserName=ClientID, Password=ClientSecret). #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$ClientId,
        [Parameter(Mandatory)][Security.SecureString]$ClientSecret,
        [string]$Name = 'halo_api_app'
    )

    $path = Get-HaloCredentialPath -Name $Name
    $cred = New-Object System.Management.Automation.PSCredential($ClientId, $ClientSecret)
    $cred | Export-Clixml -Path $path
    Write-Host "Saved Halo API app credential to: $path" -ForegroundColor Green
    $path
}

function Get-HaloApiAppCredential {
    [CmdletBinding()]
    param([string]$Name = 'halo_api_app')

    $path = Get-HaloCredentialPath -Name $Name
    if (-not (Test-Path $path)) { return $null }
    Import-Clixml -Path $path
}

function ConvertTo-PlainText {
    [CmdletBinding()]
    param([Parameter(Mandatory)][Security.SecureString]$SecureString)

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try { [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

# -----------------------------
# Halo API auth + request helpers
# -----------------------------
function Connect-HaloApi {
    [CmdletBinding()]
    param(
        # Resource Server base URL. Example: https://zahezone.halopsa.com/api
        [Parameter(Mandatory)][string]$HaloResourceServer,

        # Authorisation Server base URL. Example: https://auth.haloservicedesk.com (hosted) OR same as instance.
        # If omitted, defaults to the same host as HaloResourceServer (without trailing /api).
        [string]$HaloAuthServer,

        # Tenant name (optional) appended to token request as ?tenant=<tenant>.
        [string]$Tenant,

        [string]$CredentialName = 'halo_api_app',
        [string]$Scope = 'all',
        [switch]$ForceTls12
    )

    if ($ForceTls12) {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    }

    $cred = Get-HaloApiAppCredential -Name $CredentialName
    if (-not $cred) {
        throw "No stored Halo API credential found. Run Save-HaloApiAppCredential first (CredentialName=$CredentialName)."
    }

    # Normalise resource server
    $rs = $HaloResourceServer.TrimEnd('/')
    if ($rs -notmatch '/api$') { $rs = "$rs/api" }

    # Default auth server to same host as resource server (without /api)
    if (-not $HaloAuthServer) {
        $HaloAuthServer = ($rs -replace '/api$','')
    }
    $as = $HaloAuthServer.TrimEnd('/')

    $tokenUri = "$as/auth/token"
    if ($Tenant) { $tokenUri = "$tokenUri?tenant=$([uri]::EscapeDataString($Tenant))" }

    $clientId = $cred.UserName
    $clientSecretPlain = ConvertTo-PlainText -SecureString $cred.Password

    $body = "grant_type=client_credentials&client_id=$([uri]::EscapeDataString($clientId))&client_secret=$([uri]::EscapeDataString($clientSecretPlain))&scope=$([uri]::EscapeDataString($Scope))"

    $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -ContentType 'application/x-www-form-urlencoded' -Body $body

    if (-not ($resp.PSObject.Properties.Name -contains 'access_token')) {
        throw "Halo token response did not include access_token. Raw: $($resp | ConvertTo-Json -Depth 5 -Compress)"
    }

    $token = $resp.access_token

    [pscustomobject]@{
        ResourceServer = $rs
        AuthServer     = $as
        Tenant         = $Tenant
        Token          = $token
        Headers        = @{ 'Authorization' = "Bearer $token"; 'Content-Type' = 'application/json'; 'Accept'='application/json' }
        ConnectedAt    = (Get-Date)
    }
}

function Invoke-HaloApi {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][ValidateSet('GET','POST','DELETE')][string]$Method,
        [Parameter(Mandatory)][string]$Path,
        [hashtable]$Query,
        $Body,
        [int]$TimeoutSeconds = 120,
        [int]$MaxRetries = 3,
        [int]$RetryDelaySeconds = 2
    )

    $uri = "$($Session.ResourceServer)/$($Path.TrimStart('/'))"

    if ($Query) {
        $qs = ($Query.GetEnumerator() | ForEach-Object {
            "$([uri]::EscapeDataString([string]$_.Key))=$([uri]::EscapeDataString([string]$_.Value))"
        }) -join '&'

        if ($qs) {
            # StrictMode-safe string concatenation (avoid "$uri?$qs" which PowerShell parses as $uri? variable)
            $uri = $uri + '?' + $qs
        }
    }

    $payload = $null
    if ($null -ne $Body) {
        if ($Body -is [string]) { $payload = $Body }
        else { $payload = ($Body | ConvertTo-Json -Depth 25) }
    }

    for ($attempt = 0; $attempt -le $MaxRetries; $attempt++) {
        try {
            if ($Method -eq 'GET') {
                return Invoke-RestMethod -Method Get -Uri $uri -Headers $Session.Headers -TimeoutSec $TimeoutSeconds
            }
            elseif ($Method -eq 'DELETE') {
                return Invoke-RestMethod -Method Delete -Uri $uri -Headers $Session.Headers -TimeoutSec $TimeoutSeconds
            }
            else {
                return Invoke-RestMethod -Method Post -Uri $uri -Headers $Session.Headers -Body $payload -TimeoutSec $TimeoutSeconds
            }
        }
        catch {
            $ex = $_.Exception
            $statusCode = $null
            $respText = $null

            if ($ex -and ($ex.PSObject.Properties.Name -contains 'Response') -and $ex.Response) {
                try { $statusCode = [int]$ex.Response.StatusCode } catch { }
                try {
                    $stream = $ex.Response.GetResponseStream()
                    if ($stream) {
                        $reader = New-Object System.IO.StreamReader($stream)
                        $respText = $reader.ReadToEnd()
                        $reader.Close()
                    }
                } catch { }
            }

            $isTimeout = ($ex.Message -match 'timed out|timeout')
            $isRetryableStatus = ($statusCode -in 429,500,502,503,504)
            $shouldRetry = ($attempt -lt $MaxRetries) -and ($isTimeout -or $isRetryableStatus)

            if ($shouldRetry) {
                $sleep = [math]::Min(60, ($RetryDelaySeconds * [math]::Pow(2, $attempt)))
                Write-Warning "Halo API call failed (attempt $($attempt+1)/$($MaxRetries+1)) $Method $uri HTTP=$statusCode. Retrying in ${sleep}s..."
                Start-Sleep -Seconds $sleep
                continue
            }

            $msg = "Halo API call failed: $Method $uri"
            if ($statusCode) { $msg += " HTTP=$statusCode" }
            if ($respText) { $msg += " Body=$respText" } else { $msg += " Error=$($ex.Message)" }
            throw $msg
        }
    }
}

# -----------------------------
# Halo lookups
# -----------------------------
function Resolve-HaloAssetTypeId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$Name
    )

    $want = ($Name -replace '\s+',' ').Trim().ToLowerInvariant()

    $resp = Invoke-HaloApi -Session $Session -Method GET -Path 'assettype' -Query @{ search = $Name; includeinactive = 'true'; includeactive='true'; pageinate='true'; page_size=200; page_no=1 }

    $types = @()
    foreach ($prop in @('assettypes','AssetTypes','types','results')) {
        if ($resp.PSObject.Properties.Name -contains $prop) { $types = $resp.$prop; break }
    }
    if (-not $types -and $resp -is [System.Collections.IEnumerable] -and -not ($resp -is [string])) { $types = $resp }

    if (-not $types -or $types.Count -eq 0) {
        throw "No asset types returned when searching for '$Name'. Provide -AssetTypeId explicitly."
    }

    $matches = @()
    foreach ($t in $types) {
        if ($t.PSObject.Properties.Name -contains 'name') {
            $n = ([string]$t.name -replace '\s+',' ').Trim().ToLowerInvariant()
            if ($n -eq $want) { $matches += $t }
        }
    }

    if ($matches.Count -eq 1) {
        return [int]$matches[0].id
    }
    elseif ($matches.Count -gt 1) {
        $ids = ($matches | ForEach-Object { $_.id }) -join ','
        throw "Multiple asset types matched name '$Name' (ids: $ids). Provide -AssetTypeId explicitly."
    }

    # No exact match; do NOT guess
    $names = ($types | Where-Object { $_.PSObject.Properties.Name -contains 'name' } | Select-Object -ExpandProperty name) -join '; '
    throw "No exact asset type name match for '$Name'. Returned types: $names. Provide -AssetTypeId explicitly."
}

function Get-HaloDefaultSiteId {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][int]$CustomerId
    )

    $resp = Invoke-HaloApi -Session $Session -Method GET -Path 'Site' -Query @{ client_id = $CustomerId; includeinactive = 'false'; includeactive = 'true'; pageinate='true'; page_size=200; page_no=1 }

    $sites = @()
    foreach ($prop in @('sites','Sites','response','results')) {
        if ($resp.PSObject.Properties.Name -contains $prop) {
            $sites = $resp.$prop
            if ($sites -and ($sites.PSObject.Properties.Name -contains 'sites')) { $sites = $sites.sites }
            break
        }
    }
    if (-not $sites -and $resp -is [System.Collections.IEnumerable] -and -not ($resp -is [string])) { $sites = $resp }

    if (-not $sites -or $sites.Count -eq 0) {
        throw "No active sites found for client_id=$CustomerId. Provide -DefaultSiteId explicitly."
    }

    $active = $sites | Where-Object { -not ($_.PSObject.Properties.Name -contains 'inactive') -or $_.inactive -eq $false } | Select-Object -First 1
    if ($active) { return [int]$active.id }

    return [int]($sites | Select-Object -First 1).id
}

function Get-HaloCustomerAssetsByType {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][int]$CustomerId,
        [Parameter(Mandatory)][int]$AssetTypeId
    )

    $page = 1
    $pageSize = 200
    $all = New-Object System.Collections.Generic.List[object]

    while ($true) {
        $resp = Invoke-HaloApi -Session $Session -Method GET -Path 'Asset' -Query @{ client_id = $CustomerId; assettype_id = $AssetTypeId; includeinactive='true'; includeactive='true'; pageinate='true'; page_no=$page; page_size=$pageSize }

        $assets = @()
        if ($resp.PSObject.Properties.Name -contains 'assets') { $assets = $resp.assets }
        elseif ($resp.PSObject.Properties.Name -contains 'Devices') { $assets = $resp.Devices }
        elseif ($resp.PSObject.Properties.Name -contains 'response') { $assets = $resp.response }

        if ($assets) { foreach ($a in $assets) { [void]$all.Add($a) } }

        $recordCount = $null
        if ($resp.PSObject.Properties.Name -contains 'record_count') { $recordCount = [int64]$resp.record_count }

        if ($null -ne $recordCount) {
            if ($all.Count -ge $recordCount) { break }
        } else {
            if (-not $assets -or $assets.Count -lt $pageSize) { break }
        }

        $page++
    }

    $all.ToArray()
}

function Get-AssetSerialFromHalo {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Asset)

    foreach ($k in @('key_field','inventory_number','key_field3')) {
        if ($Asset.PSObject.Properties.Name -contains $k) {
            $v = $Asset.$k
            if ($v -and -not [string]::IsNullOrWhiteSpace([string]$v)) {
                return ([string]$v).Trim()
            }
        }
    }
    return $null
}

function Get-AssetTypeIdFromHaloAsset {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Asset)

    if ($Asset.PSObject.Properties.Name -contains 'assettype_id') { return [int]$Asset.assettype_id }
    if ($Asset.PSObject.Properties.Name -contains 'assettype') { return [int]$Asset.assettype }
    if ($Asset.PSObject.Properties.Name -contains 'assettypeid') { return [int]$Asset.assettypeid }
    return $null
}

# -----------------------------
# Main sync
# -----------------------------
function Sync-MosyleJsonToHaloMobileAssets {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)][string]$MosyleJsonPath,
        [Parameter(Mandatory)][string]$HaloResourceServer,
        [string]$HaloAuthServer,
        [string]$Tenant,
        [Parameter(Mandatory)][int]$CustomerId,
        [int]$DefaultSiteId,
        [int]$AssetTypeId,
        [string]$AssetTypeName = 'Mobile Device',
        [string]$ManufacturerName = 'Apple',
        [switch]$WriteSerialToInventoryNumber,
        [switch]$ForceTls12
    )

    if (-not (Test-Path $MosyleJsonPath)) {
        throw "Mosyle JSON file not found: $MosyleJsonPath"
    }

    $mosyle = Get-Content -Path $MosyleJsonPath -Raw | ConvertFrom-Json
    if (-not $mosyle) { throw "Mosyle JSON parsed to empty. Check file content: $MosyleJsonPath" }

    if ($mosyle -isnot [System.Collections.IEnumerable] -or $mosyle -is [string]) { $mosyle = @($mosyle) }

    $mosyle = $mosyle | Where-Object { $_.PSObject.Properties.Name -contains 'serial_number' -and -not [string]::IsNullOrWhiteSpace([string]$_.serial_number) }

    $session = Connect-HaloApi -HaloResourceServer $HaloResourceServer -HaloAuthServer $HaloAuthServer -Tenant $Tenant -ForceTls12:$ForceTls12

    if (-not $AssetTypeId) {
        $AssetTypeId = Resolve-HaloAssetTypeId -Session $session -Name $AssetTypeName
    }

    if (-not $DefaultSiteId) {
        $DefaultSiteId = Get-HaloDefaultSiteId -Session $session -CustomerId $CustomerId
    }

    $haloAssets = Get-HaloCustomerAssetsByType -Session $session -CustomerId $CustomerId -AssetTypeId $AssetTypeId

    # Index existing Halo assets by serial
    $haloBySerial = @{}
    foreach ($a in $haloAssets) {
        $serial = Get-AssetSerialFromHalo -Asset $a
        if ($serial -and -not $haloBySerial.ContainsKey($serial)) { $haloBySerial[$serial] = $a }
    }

    $mosyleSerials = New-Object 'System.Collections.Generic.HashSet[string]'
    foreach ($m in $mosyle) { [void]$mosyleSerials.Add(([string]$m.serial_number).Trim()) }

    $created = 0; $updated = 0; $deactivated = 0; $unchanged = 0

    foreach ($m in $mosyle) {
        $serial = ([string]$m.serial_number).Trim()
        $deviceName = if ($m.PSObject.Properties.Name -contains 'device_name') { [string]$m.device_name } else { '' }
        $osversion = if ($m.PSObject.Properties.Name -contains 'osversion') { [string]$m.osversion } else { '' }

        if ($haloBySerial.ContainsKey($serial)) {
            $existing = $haloBySerial[$serial]

            $desired = [ordered]@{
                id          = $existing.id
                client_id   = $CustomerId
                site_id     = $DefaultSiteId
                assettype_id= $AssetTypeId
                inactive    = $false
                key_field   = $serial
                key_field2  = $osversion
                manufacturer_name = $ManufacturerName
                notes       = $deviceName
            }
            if ($WriteSerialToInventoryNumber) { $desired.inventory_number = $serial }

            $needsUpdate = $false
            foreach ($p in $desired.Keys) {
                if (-not ($existing.PSObject.Properties.Name -contains $p)) { $needsUpdate = $true; break }
                $cur = $existing.$p
                $new = $desired[$p]
                if (([string]$cur) -ne ([string]$new)) { $needsUpdate = $true; break }
            }

            if ($needsUpdate) {
                if ($PSCmdlet.ShouldProcess("Asset id=$($existing.id) serial=$serial", 'Update')) {
                    Invoke-HaloApi -Session $session -Method POST -Path 'Asset' -Body @($desired) | Out-Null
                }
                $updated++
            } else {
                $unchanged++
            }
        }
        else {
            $newAsset = [ordered]@{
                client_id   = $CustomerId
                site_id     = $DefaultSiteId
                assettype_id= $AssetTypeId
                inactive    = $false
                key_field   = $serial
                key_field2  = $osversion
                manufacturer_name = $ManufacturerName
                notes       = $deviceName
            }
            if ($WriteSerialToInventoryNumber) { $newAsset.inventory_number = $serial }

            if ($PSCmdlet.ShouldProcess("serial=$serial", 'Create')) {
                Invoke-HaloApi -Session $session -Method POST -Path 'Asset' -Body @($newAsset) | Out-Null
            }
            $created++
        }
    }

    # Deactivate ONLY assets that truly are of the Mobile Device asset type
    foreach ($a in $haloAssets) {
        $atype = Get-AssetTypeIdFromHaloAsset -Asset $a
        if ($null -ne $atype -and $atype -ne $AssetTypeId) {
            continue
        }

        $serial = Get-AssetSerialFromHalo -Asset $a
        if (-not $serial) { continue }

        if (-not $mosyleSerials.Contains($serial)) {
            $isInactive = $false
            if ($a.PSObject.Properties.Name -contains 'inactive') { $isInactive = [bool]$a.inactive }

            if (-not $isInactive) {
                $patch = [ordered]@{ id = $a.id; inactive = $true }
                if ($PSCmdlet.ShouldProcess("Asset id=$($a.id) serial=$serial", 'Set inactive=true')) {
                    Invoke-HaloApi -Session $session -Method POST -Path 'Asset' -Body @($patch) | Out-Null
                }
                $deactivated++
            }
        }
    }

    [pscustomobject]@{
        CustomerId    = $CustomerId
        AssetTypeId   = $AssetTypeId
        DefaultSiteId = $DefaultSiteId
        MosyleCount   = $mosyle.Count
        HaloCount     = $haloAssets.Count
        Created       = $created
        Updated       = $updated
        Deactivated   = $deactivated
        Unchanged     = $unchanged
        CompletedAt   = (Get-Date)
    }
}

<#
.EXAMPLE
  . .\MosyleToHaloMobileAssetsSync_v3.ps1

  # One-time: store Halo API app credentials
  $clientId = 'YOUR_HALO_CLIENT_ID'
  $clientSecret = Read-Host -AsSecureString 'Halo Client Secret'
  Save-HaloApiAppCredential -ClientId $clientId -ClientSecret $clientSecret

  # Dry-run
  Sync-MosyleJsonToHaloMobileAssets -MosyleJsonPath C:\Temp\mosyle_ios_devices.json `
    -HaloResourceServer https://zahezone.halopsa.com/api `
    -CustomerId 235 -WhatIf

  # Real run
  Sync-MosyleJsonToHaloMobileAssets -MosyleJsonPath C:\Temp\mosyle_ios_devices.json `
    -HaloResourceServer https://zahezone.halopsa.com/api `
    -CustomerId 235
#>
