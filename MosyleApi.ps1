Set-StrictMode -Version Latest

# -----------------------------
# Credential Storage (DPAPI)
# -----------------------------
function Get-MosyleCredentialPath {
    [CmdletBinding()]
    param(
        [Parameter()][string]$Name = "mosyle_cred"
    )
    $base = Join-Path $env:APPDATA "ZaheZone\Mosyle"
    if (-not (Test-Path $base)) { New-Item -ItemType Directory -Path $base -Force | Out-Null }
    Join-Path $base "$Name.xml"
}

function Save-MosyleCredential {
    <#
      Saves a PSCredential (email + secure password) to disk using Export-Clixml.
      The password is protected by DPAPI for the current user on this machine.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Email,
        [Parameter(Mandatory)][Security.SecureString]$Password,
        [Parameter()][string]$Name = "mosyle_cred"
    )

    $path = Get-MosyleCredentialPath -Name $Name
    $cred = New-Object System.Management.Automation.PSCredential($Email, $Password)
    $cred | Export-Clixml -Path $path
    Write-Host "Saved Mosyle credential to: $path" -ForegroundColor Green
    $path
}

function Get-MosyleCredential {
    <#
      Loads a PSCredential from disk (created by Save-MosyleCredential).
      Returns $null if not found.
    #>
    [CmdletBinding()]
    param(
        [Parameter()][string]$Name = "mosyle_cred"
    )

    $path = Get-MosyleCredentialPath -Name $Name
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
# Auth Helpers
# -----------------------------
function Get-MosyleJwtFromAuthHeader {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$AuthorizationHeaderValue)

    $v = $AuthorizationHeaderValue.Trim()
    while ($v -match '^(?i)Bearer\s+') {
        $v = ($v -replace '^(?i)Bearer\s+', '').Trim()
    }
    if ([string]::IsNullOrWhiteSpace($v)) { throw "Authorization header did not contain a usable JWT." }
    $v
}

function Connect-MosyleApi {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$AccessToken,

        # Option A: Supply email + secure password directly
        [Parameter()][string]$Email,
        [Parameter()][Security.SecureString]$Password,

        # Option B: Use stored credential name (default mosyle_cred)
        [Parameter()][string]$CredentialName = "mosyle_cred",

        [Parameter()][string]$BaseUri = "https://businessapi.mosyle.com/v1",
        [Parameter()][switch]$ForceTls12
    )

    if ($ForceTls12) {
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    }

    # If email/password not supplied, try stored credential
    if (-not $Email -or -not $Password) {
        $stored = Get-MosyleCredential -Name $CredentialName
        if ($stored) {
            $Email = $stored.UserName
            $Password = $stored.Password
        }
    }

    if (-not $Email -or -not $Password) {
        throw "No credentials provided and no stored credential found. Use Save-MosyleCredential first or pass -Email/-Password."
    }

    $loginUri = "$BaseUri/login"
    $pwPlain  = ConvertTo-PlainText -SecureString $Password

    $headers = @{
        "accessToken"  = $AccessToken
        "Content-Type" = "application/json"
    }

    $bodyJson = @{ email = $Email; password = $pwPlain } | ConvertTo-Json -Depth 5

    # Use Invoke-WebRequest so we can read response headers for Authorization. [3](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/invoke-webrequest?view=powershell-7.6)
    $resp = Invoke-WebRequest -Method Post -Uri $loginUri -Headers $headers -Body $bodyJson -UseBasicParsing

    $authHeader = $resp.Headers["Authorization"]
    if (-not $authHeader) { throw "Login did not return Authorization header. Check API token / email / password." }

    $jwt = Get-MosyleJwtFromAuthHeader -AuthorizationHeaderValue $authHeader

    [pscustomobject]@{
        BaseUri  = $BaseUri
        Headers  = @{
            "Content-Type"  = "application/json"
            "accessToken"   = $AccessToken
            "Authorization" = "Bearer $jwt"
        }
        Jwt = $jwt
        ConnectedAt = (Get-Date)
    }
}

# -----------------------------
# Robust POST helper
# -----------------------------
function Invoke-MosyleApi {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][hashtable]$Body,
        [Parameter()][int]$TimeoutSeconds = 120
    )

    $uri  = "$($Session.BaseUri)/$($Path.TrimStart('/'))"
    $json = $Body | ConvertTo-Json -Depth 12

    try {
        Invoke-RestMethod -Method Post -Uri $uri -Headers $Session.Headers -Body $json -TimeoutSec $TimeoutSeconds
    }
    catch {
        # StrictMode-safe error extraction. StrictMode errors on missing properties are expected behavior. [1](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/set-strictmode?view=powershell-7.5)[2](https://github.com/PowerShell/PowerShell/issues/10875)
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

        $msg = "Mosyle API request failed. Endpoint=$uri"
        if ($statusCode) { $msg += " HTTP=$statusCode" }
        if ($respText) { $msg += " Body=$respText" } else { $msg += " Error=$($ex.Message)" }
        throw $msg
    }
}

# -----------------------------
# Device Listing (iOS)
# -----------------------------
function Get-MosyleIOSDevices {
    <#
      Returns a flat list of iOS devices with selected columns.
      Default columns: device_name, device_type, serial_number, os
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter()][ValidateRange(1, 500)][int]$PageSize = 50,
        [Parameter()][switch]$Supervised,
        [Parameter()][string[]]$SpecificColumns = @("device_name","device_type","serial_number","os")
    )

    $all = New-Object System.Collections.Generic.List[object]
    $page = 1
    $rows = $null

    while ($true) {
        $options = @{
            os = "ios"
            page = $page
            page_size = $PageSize
            specific_columns = $SpecificColumns
        }
        if ($Supervised) { $options.supervised = "true" }

        $result = Invoke-MosyleApi -Session $Session -Path "devices" -Body @{
            operation = "list"
            options   = $options
        }

        if (-not ($result.PSObject.Properties.Name -contains 'status')) {
            throw "Unexpected Mosyle response (missing top-level 'status'). Raw: $($result | ConvertTo-Json -Depth 10 -Compress)"
        }
        if ($result.status -ne "OK") {
            throw "Mosyle returned status=$($result.status). Raw: $($result | ConvertTo-Json -Depth 10 -Compress)"
        }

        $entry = $null
        if (($result.PSObject.Properties.Name -contains 'response') -and $result.response) {
            $entry = $result.response | Select-Object -First 1
        }
        if (-not $entry) { break }

        # If entry has a status (only present for some error responses), handle it safely
        if ($entry.PSObject.Properties.Name -contains 'status') {
            if ($entry.status -eq "DEVICES_NOTFOUND") { return @() }
            if ($entry.status -ne "OK") {
                $info = if ($entry.PSObject.Properties.Name -contains 'info') { $entry.info } else { "" }
                throw "Mosyle response entry status=$($entry.status) $info"
            }
        }

        $devices = @()
        if (($entry.PSObject.Properties.Name -contains 'devices') -and $entry.devices) {
            $devices = $entry.devices
        }

        foreach ($d in $devices) {
            # Emit ONLY the requested columns
            $o = [ordered]@{}
            foreach ($col in $SpecificColumns) {
                $o[$col] = if ($d.PSObject.Properties.Name -contains $col) { $d.$col } else { $null }
            }
            [void]$all.Add([pscustomobject]$o)
        }

        if ($rows -eq $null -and ($entry.PSObject.Properties.Name -contains 'rows')) {
            $rows = [int]$entry.rows
        }

        if ($rows -ne $null) {
            if ($all.Count -ge $rows) { break }
        } else {
            if ($devices.Count -lt $PageSize) { break }
        }

        $page++
    }

    $all.ToArray()
}

# -----------------------------
# Export
# -----------------------------
function Export-MosyleDevicesToJson {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$Path,
        [Parameter()][ValidateRange(1, 500)][int]$PageSize = 50,
        [Parameter()][switch]$Supervised,
        [Parameter()][string[]]$SpecificColumns = @("device_name","device_type","serial_number","os")
    )

    $devices = Get-MosyleIOSDevices -Session $Session -PageSize $PageSize -Supervised:$Supervised -SpecificColumns $SpecificColumns

    $dir = Split-Path -Parent $Path
    if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }

    $devices | ConvertTo-Json -Depth 6 | Set-Content -Path $Path -Encoding UTF8
    Write-Host ("Exported {0} iOS devices to {1}" -f $devices.Count, $Path) -ForegroundColor Green
    $Path
}