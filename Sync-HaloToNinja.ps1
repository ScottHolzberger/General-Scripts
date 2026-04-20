<#
.SYNOPSIS
Sync HaloPSA customer number into NinjaOne Organization custom field by matching organization names.

.DESCRIPTION
- Dry-run by default (prints actions only). Use -Commit to write changes.
- Exact match first using normalized names.
- Optional fuzzy match uses similarity score (Levenshtein-based) with threshold + ambiguity guard.

REQUIREMENTS
- HaloPSA OAuth2 Client Credentials (Client ID/Secret) and correct Auth/Resource URLs.
- NinjaOne OAuth2 Client Credentials (Client ID/Secret) and correct Base URL for your region.
- NinjaOne org custom field must exist and allow API read/write.

REFERENCES
- NinjaOne OAuth client credentials flow and token endpoint (/ws/oauth/token). (NinjaOne API docs) [2](https://app.ninjaone.com/apidocs-beta/authorization/flows/client-credentials-flow)
- NinjaOne org custom fields endpoint PATCH /api/v2/organization/{id}/custom-fields. (NinjaOne API docs / PS Gallery example) [3](https://app.ninjaone.com/apidocs/)[4](https://www.powershellgallery.com/packages/NinjaOne/1.10.1/Content/Public%5CSet%5CSet-NinjaOneOrganisationCustomFields.ps1)
- HaloPSA API is OAuth2 token based (Bearer). [5](https://haloitsm.com/apidoc/)[6](https://www.stitchflow.com/user-management/halopsa/api)
#>

[CmdletBinding()]
param(
    # --- Mode ---
    [switch]$Commit,                        # default is Dry-Run
    [switch]$FuzzyMatch,                    # enable fuzzy matching if exact match fails
    [ValidateRange(0.50, 0.99)]
    [double]$FuzzyThreshold = 0.92,         # safe starting point
    [ValidateRange(0.00, 0.50)]
    [double]$FuzzyAmbiguityGap = 0.05,      # skip if best-vs-second-best too close

    # --- HaloPSA ---
    [Parameter(Mandatory=$true)]
    [string]$HaloAuthServer,                # e.g. https://zahezone.halopsa.com/auth
    [Parameter(Mandatory=$true)]
    [string]$HaloResourceServer,            # e.g. https://zahezone.halopsa.com/api
    [string]$HaloTenant = "",               # if your instance requires tenant query param
    [Parameter(Mandatory=$true)]
    [string]$HaloClientId,
    [Parameter(Mandatory=$true)]
    [string]$HaloClientSecret,
    [string]$HaloScope = "all",

    # Which Halo field to store in Ninja (defaults to "id")
    [string]$HaloCustomerNumberField = "id",

    # --- NinjaOne ---
    [Parameter(Mandatory=$true)]
    [string]$NinjaBaseUrl = "https://app.ninjaone.com",  # change to your region if needed
    [Parameter(Mandatory=$true)]
    [string]$NinjaClientId,
    [Parameter(Mandatory=$true)]
    [string]$NinjaClientSecret,
    [string]$NinjaScope = "management monitoring",
    [Parameter(Mandatory=$true)]
    [string]$NinjaCustomFieldName
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# -------------------------
# Helpers
# -------------------------
function Normalize-Name {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return "" }

    $n = $Name.ToLowerInvariant()
    $n = $n -replace '&', ' and '
    $n = $n -replace '[^a-z0-9 ]', ''   # drop punctuation
    $n = $n -replace '\s+', ' '
    return $n.Trim()
}

function Get-LevenshteinDistance {
    param([string]$a, [string]$b)

    if ($null -eq $a) { $a = "" }
    if ($null -eq $b) { $b = "" }

    $lenA = $a.Length
    $lenB = $b.Length
    if ($lenA -eq 0) { return $lenB }
    if ($lenB -eq 0) { return $lenA }

    $d = New-Object 'int[,]' ($lenA + 1), ($lenB + 1)
    for ($i = 0; $i -le $lenA; $i++) { $d[$i,0] = $i }
    for ($j = 0; $j -le $lenB; $j++) { $d[0,$j] = $j }

    for ($i = 1; $i -le $lenA; $i++) {
        for ($j = 1; $j -le $lenB; $j++) {
            $cost = if ($a[$i-1] -eq $b[$j-1]) { 0 } else { 1 }
            $del  = $d[$i-1,$j] + 1
            $ins  = $d[$i,$j-1] + 1
            $sub  = $d[$i-1,$j-1] + $cost
            $d[$i,$j] = [Math]::Min([Math]::Min($del, $ins), $sub)
        }
    }
    return $d[$lenA,$lenB]
}

function Get-Similarity {
    param([string]$a, [string]$b)

    $aN = Normalize-Name $a
    $bN = Normalize-Name $b

    if ($aN.Length -eq 0 -or $bN.Length -eq 0) { return 0.0 }
    if ($aN -eq $bN) { return 1.0 }

    $dist = Get-LevenshteinDistance $aN $bN
    $maxLen = [Math]::Max($aN.Length, $bN.Length)
    return [Math]::Round((1.0 - ($dist / $maxLen)), 4)
}

function Invoke-Json {
    param(
        [Parameter(Mandatory=$true)][string]$Method,
        [Parameter(Mandatory=$true)][string]$Uri,
        [Parameter(Mandatory=$true)][hashtable]$Headers,
        [object]$Body = $null,
        [string]$ContentType = "application/json"
    )

    if ($Body -ne $null) {
        $json = $Body | ConvertTo-Json -Depth 10
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $Headers -ContentType $ContentType -Body $json
    } else {
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $Headers
    }
}

# -------------------------
# Auth: HaloPSA (OAuth2 client_credentials)
# -------------------------
function Get-HaloAccessToken {
    $tokenUrl = ($HaloAuthServer.TrimEnd('/') + "/token")
    if (-not [string]::IsNullOrWhiteSpace($HaloTenant)) {
        $tokenUrl = $tokenUrl + "?tenant=$([uri]::EscapeDataString($HaloTenant))"
    }

    $form = @{
        grant_type    = "client_credentials"
        client_id     = $HaloClientId
        client_secret = $HaloClientSecret
        scope         = $HaloScope
    }

    Write-Verbose "Halo token URL: $tokenUrl"
    $resp = Invoke-RestMethod -Method POST -Uri $tokenUrl -ContentType "application/x-www-form-urlencoded" -Body $form
    if (-not $resp.access_token) { throw "Halo token response did not include access_token." }
    return $resp.access_token
}

# -------------------------
# Auth: NinjaOne (OAuth2 client_credentials)
# -------------------------
function Get-NinjaAccessToken {
    $tokenUrl = ($NinjaBaseUrl.TrimEnd('/') + "/ws/oauth/token")  # documented token endpoint [2](https://app.ninjaone.com/apidocs-beta/authorization/flows/client-credentials-flow)

    $form = @{
        grant_type    = "client_credentials"
        client_id     = $NinjaClientId
        client_secret = $NinjaClientSecret
        scope         = $NinjaScope
    }

    Write-Verbose "Ninja token URL: $tokenUrl"
    $resp = Invoke-RestMethod -Method POST -Uri $tokenUrl -ContentType "application/x-www-form-urlencoded" -Body $form
    if (-not $resp.access_token) { throw "Ninja token response did not include access_token." }
    return $resp.access_token
}

# -------------------------
# Fetch Halo Clients (pagination)
# -------------------------
function Get-HaloClients {
    $all = @()
    $pageNo = 1
    $pageSize = 200

    while ($true) {
        # This query style is common in Halo modules/docs; if your API doesn’t support it, we’ll see it immediately. ent/)
        $url = $HaloResourceServer.TrimEnd('/') + "/Client?pageinate=true&page_size=$pageSize&page_no=$pageNo&includeactive=true&includeinactive=true"
        Write-Verbose "Halo clients URL: $url"

        $resp = Invoke-Json -Method GET -Uri $url -Headers $script:HaloHeaders

        # Halo responses vary; try a few shapes safely
        $items = $null
        if ($resp -is [System.Array]) { $items = $resp }
        elseif ($resp.clients) { $items = $resp.clients }
        elseif ($resp.result) { $items = $resp.result }
        elseif ($resp.data) { $items = $resp.data }
        else { $items = @() }

        if (-not $items -or $items.Count -eq 0) { break }

        $all += $items
        if ($items.Count -lt $pageSize) { break }
        $pageNo++
    }

    return $all
}

# -------------------------
# Ninja endpoints
# -------------------------
function Get-NinjaOrganizations {
    $url = $NinjaBaseUrl.TrimEnd('/') + "/api/v2/organizations"
    Write-Verbose "Ninja orgs URL: $url"
    return Invoke-Json -Method GET -Uri $url -Headers $script:NinjaHeaders
}

function Get-NinjaOrgCustomFields {
    param([int]$OrgId)
    $url = $NinjaBaseUrl.TrimEnd('/') + "/api/v2/organization/$OrgId/custom-fields"  # org custom fields endpoint [3](https://app.ninjaone.com/apidocs/)
    return Invoke-Json -Method GET -Uri $url -Headers $script:NinjaHeaders
}

function Set-NinjaOrgCustomField {
    param(
        [int]$OrgId,
        [string]$FieldName,
        [string]$Value
    )

    $url = $NinjaBaseUrl.TrimEnd('/') + "/api/v2/organization/$OrgId/custom-fields"  # PATCH update field values [3](https://app.ninjaone.com/apidocs/)[4](https://www.powershellgallery.com/packages/NinjaOne/1.10.1/Content/Public%5CSet%5CSet-NinjaOneOrganisationCustomFields.ps1)
    $body = @{ $FieldName = $Value }

    if ($Commit) {
        Invoke-Json -Method PATCH -Uri $url -Headers $script:NinjaHeaders -Body $body | Out-Null
        return "UPDATED"
    } else {
        return "DRYRUN"
    }
}

# -------------------------
# Main
# -------------------------
$mode = if ($Commit) { "COMMIT" } else { "DRY-RUN" }
Write-Host "Mode: $mode"
Write-Host "FuzzyMatch: $FuzzyMatch (Threshold: $FuzzyThreshold, Gap: $FuzzyAmbiguityGap)"
Write-Host "Ninja Org Field: $NinjaCustomFieldName"
Write-Host ""

# Tokens
Write-Host "Auth: Halo..."
$haloToken  = Get-HaloAccessToken
Write-Host "Auth: Ninja..."
$ninjaToken = Get-NinjaAccessToken

$script:HaloHeaders  = @{ Authorization = "Bearer $haloToken";  Accept = "application/json" }   # Halo bearer tokens [5](https://haloitsm.com/apidoc/)[6](https://www.stitchflow.com/user-management/halopsa/api)
$script:NinjaHeaders = @{ Authorization = "Bearer $ninjaToken"; Accept = "application/json" }  # Ninja bearer tokens [2](https://app.ninjaone.com/apidocs-beta/authorization/flows/client-credentials-flow)

# Fetch Halo clients
Write-Host "Fetching Halo clients..."
$haloClientsRaw = Get-HaloClients
Write-Host ("Halo returned: " + ($haloClientsRaw | Measure-Object).Count)

# Build lookup
$haloByName = @{}
$haloList = New-Object System.Collections.Generic.List[object]

foreach ($c in $haloClientsRaw) {
    $name =
        if ($c.name) { $c.name }
        elseif ($c.client_name) { $c.client_name }
        elseif ($c.Name) { $c.Name }
        else { $null }

    if ([string]::IsNullOrWhiteSpace($name)) { continue }
    $norm = Normalize-Name $name
    if ([string]::IsNullOrWhiteSpace($norm)) { continue }

    $custNo = $null
    try { $custNo = $c.$HaloCustomerNumberField } catch { $custNo = $null }
    if ($null -eq $custNo -and $c.id) { $custNo = $c.id }
    if ($null -eq $custNo) { continue }

    if (-not $haloByName.ContainsKey($norm)) {
        $obj = [pscustomobject]@{
            Name = $name
            Norm = $norm
            CustomerNo = "$custNo"
        }
        $haloByName[$norm] = $obj
        $haloList.Add($obj) | Out-Null
    }
}

Write-Host ("Halo unique-name map: " + $haloByName.Count)

# Fetch Ninja orgs
Write-Host "Fetching NinjaOne organizations..."
$ninjaOrgs = Get-NinjaOrganizations
Write-Host ("Ninja returned: " + ($ninjaOrgs | Measure-Object).Count)
Write-Host ""

$matched = 0
$changed = 0
$skippedNoMatch = 0
$skippedAmbiguous = 0

foreach ($org in $ninjaOrgs) {
    $orgName = $org.name
    if ([string]::IsNullOrWhiteSpace($orgName)) { continue }

    $orgId = [int]$org.id
    $orgNorm = Normalize-Name $orgName

    $hit = $null
    $matchType = "none"
    $score = 0.0

    # Exact match
    if ($haloByName.ContainsKey($orgNorm)) {
        $hit = $haloByName[$orgNorm]
        $matchType = "exact"
        $score = 1.0
    }
    elseif ($FuzzyMatch) {
        $best = $null
        $second = $null

        foreach ($h in $haloList) {
            $s = Get-Similarity $orgName $h.Name
            if ($null -eq $best -or $s -gt $best.Score) {
                $second = $best
                $best = [pscustomobject]@{ Score = $s; Hit = $h }
            } elseif ($null -eq $second -or $s -gt $second.Score) {
                $second = [pscustomobject]@{ Score = $s; Hit = $h }
            }
        }

        if ($best -and $best.Score -ge $FuzzyThreshold) {
            $gap = if ($second) { $best.Score - $second.Score } else { 1.0 }
            if ($gap -lt $FuzzyAmbiguityGap) {
                $skippedAmbiguous++
                continue
            }
            $hit = $best.Hit
            $matchType = "fuzzy"
            $score = $best.Score
        }
    }

    if (-not $hit) {
        $skippedNoMatch++
        Write-Host ("SKIPPED-NOMATCH | Org='{0}' (#{1})" -f $orgName, $orgId) -ForegroundColor Yellow
    continue
    }
    

    $matched++

    # current fields
    $currentFields = Get-NinjaOrgCustomFields -OrgId $orgId
    $currentValue = $null
    try { $currentValue = $currentFields.$NinjaCustomFieldName } catch { $currentValue = $null }

    $targetValue = $hit.CustomerNo
    if ($currentValue -eq $targetValue) { continue }

    $result = Set-NinjaOrgCustomField -OrgId $orgId -FieldName $NinjaCustomFieldName -Value $targetValue
    $changed++

    Write-Host ("{0} | Org='{1}' (#{2}) <= Halo='{3}' [{4}:{5}] | Field='{6}' Old='{7}' New='{8}'" -f `
        $result, $orgName, $orgId, $hit.Name, $matchType, $score, $NinjaCustomFieldName, $currentValue, $targetValue)
}

Write-Host ""
Write-Host "Summary:"
Write-Host "  Matched:           $matched"
Write-Host "  Updated/DryRun:    $changed"
Write-Host "  Skipped NoMatch:   $skippedNoMatch"
Write-Host "  Skipped Ambiguous: $skippedAmbiguous"