<#
.SYNOPSIS
  NinjaOne OAuth 2.0 (Client Credentials) connector + API helper

.DESCRIPTION
  - Requests an OAuth access token from /ws/oauth/token using client_credentials
  - Caches token in-memory until expiry
  - Provides Invoke-NinjaOneApi to call /v2 endpoints with Bearer token

.NOTES
  Token endpoint: https://{region}.ninjarmm.com/ws/oauth/token
  API endpoints:  https://{region}.ninjarmm.com/v2/...

  Per NinjaOne docs, request token with grant_type=client_credentials and use returned access_token as Bearer. [1](https://app.ninjaone.com/apidocs-beta/authorization/flows/client-credentials-flow)[2](https://oc.ninjarmm.com/apidocs/?links.active=authorization)
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# In-memory token cache
$script:NinjaOneTokenCache = $null

function Get-NinjaOneAccessToken {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$BaseUrl,  # e.g. https://oc.ninjarmm.com

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ClientSecret,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$Scope = "monitoring"
  )

  # Reuse cached token if still valid (add 60s safety buffer)
  if ($script:NinjaOneTokenCache -and $script:NinjaOneTokenCache.ExpiresAtUtc) {
    $now = [DateTime]::UtcNow
    if ($now -lt $script:NinjaOneTokenCache.ExpiresAtUtc.AddSeconds(-60)) {
      return $script:NinjaOneTokenCache
    }
  }

  $tokenUri = ($BaseUrl.TrimEnd("/") + "/ws/oauth/token")

  # OAuth token request is form-urlencoded: grant_type, client_id, client_secret, scope [1](https://app.ninjaone.com/apidocs-beta/authorization/flows/client-credentials-flow)[2](https://oc.ninjarmm.com/apidocs/?links.active=authorization)
  $body = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = $Scope
  }

  try {
    $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -ContentType "application/x-www-form-urlencoded" -Body $body
  }
  catch {
    throw "Failed to obtain NinjaOne OAuth token from $tokenUri. $($_.Exception.Message)"
  }

  if (-not $resp.access_token) {
    throw "Token response did not include access_token. Raw response: $($resp | ConvertTo-Json -Depth 10)"
  }

  # expires_in is typically 3600 seconds in example responses [1](https://app.ninjaone.com/apidocs-beta/authorization/flows/client-credentials-flow)
  $expiresIn = 3600
  if ($resp.expires_in) { $expiresIn = [int]$resp.expires_in }

  $cache = [pscustomobject]@{
    AccessToken  = [string]$resp.access_token
    TokenType    = if ($resp.token_type) { [string]$resp.token_type } else { "Bearer" }
    Scope        = if ($resp.scope) { [string]$resp.scope } else { $Scope }
    ExpiresIn    = $expiresIn
    ObtainedAtUtc = [DateTime]::UtcNow
    ExpiresAtUtc  = ([DateTime]::UtcNow).AddSeconds($expiresIn)
  }

  $script:NinjaOneTokenCache = $cache
  return $cache
}

function Invoke-NinjaOneApi {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$BaseUrl,      # e.g. https://oc.ninjarmm.com

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ClientId,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$ClientSecret,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$Path,         # e.g. /v2/organizations  OR organizations (we normalize)

    [Parameter(Mandatory = $false)]
    [ValidateSet("GET","POST","PUT","PATCH","DELETE")]
    [string]$Method = "GET",

    [Parameter(Mandatory = $false)]
    [hashtable]$Query,

    [Parameter(Mandatory = $false)]
    $Body,

    [Parameter(Mandatory = $false)]
    [string]$Scope = "monitoring"
  )

  # Normalize Path to /v2/...
  $p = $Path.Trim()
  if (-not $p.StartsWith("/")) { $p = "/$p" }
  if (-not $p.StartsWith("/v2/")) {
    if ($p -eq "/v2") { $p = "/v2/" }
    elseif ($p.StartsWith("/v2")) { } # ok
    else { $p = "/v2" + $p }         # e.g. /organizations -> /v2/organizations
  }

  $token = Get-NinjaOneAccessToken -BaseUrl $BaseUrl -ClientId $ClientId -ClientSecret $ClientSecret -Scope $Scope

  $uriBuilder = [System.UriBuilder]::new(($BaseUrl.TrimEnd("/") + $p))

  if ($Query) {
    $qs = $Query.GetEnumerator() | ForEach-Object {
      "{0}={1}" -f [System.Web.HttpUtility]::UrlEncode($_.Key),
                [System.Web.HttpUtility]::UrlEncode([string]$_.Value)
    }
    $uriBuilder.Query = ($qs -join '&')
  }

  # Use Authorization: Bearer <token> for API calls [1](https://app.ninjaone.com/apidocs-beta/authorization/flows/client-credentials-flow)[2](https://oc.ninjarmm.com/apidocs/?links.active=authorization)
  $headers = @{
    "Accept"        = "application/json"
    "Authorization" = "Bearer $($token.AccessToken)"
  }

  $irmParams = @{
    Method  = $Method
    Uri     = $uriBuilder.Uri.AbsoluteUri
    Headers = $headers
  }

  if ($null -ne $Body) {
    # If caller provides a hashtable/object, send JSON for non-GET methods
    if ($Method -ne "GET") {
      $irmParams["ContentType"] = "application/json"
      $irmParams["Body"] = ($Body | ConvertTo-Json -Depth 10)
    }
  }

  try {
    return Invoke-RestMethod @irmParams
  }
  catch {
    $msg = $_.Exception.Message
    throw "NinjaOne API call failed: $Method $($uriBuilder.Uri.AbsoluteUri) :: $msg"
  }
}

function Clear-NinjaOneTokenCache {
  [CmdletBinding()]
  param()
  $script:NinjaOneTokenCache = $null
}
