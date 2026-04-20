<#
  Halo ticket create + resolve (Automation/AI charge type)

  Uses:
   - OAuth client_credentials (token endpoint)
   - POST /api/tickets (ARRAY payload)
   - POST /api/Actions (ARRAY payload)

  Runtime params:
   - CustomerId, Summary, Detail, TimeMinutes, ResolutionSummary
#>

param(
    [Parameter(Mandatory=$true)] [int]    $CustomerId,
    [Parameter(Mandatory=$true)] [string] $Summary,
    [Parameter(Mandatory=$true)] [string] $Detail,
    [Parameter(Mandatory=$true)] [int]    $TimeMinutes,
    [Parameter(Mandatory=$true)] [string] $ResolutionSummary
)

$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.Web

# =========================
# CONFIG (EDIT THESE)
# =========================
$BaseHost          = "https://zahezone.halopsa.com"

# OAuth Client Credentials (set within script)
$OAuthClientId     = "d98fd052-79ba-4b4f-a0c3-20fbd6b01245"
$OAuthClientSecret = "zMBDJgJuFrDPPHlaJNgrcgWuMjXc9iuQfYpqYIv48EQ"

# Tenant query param (blank if not required)
$Tenant            = "zahezone"   # or ""

# Debug switch: prints JSON bodies when failing
$DebugRequests     = $true

# =========================
# STATIC DEFAULTS (ticket create)
# Keep minimal but include mandatory fields for your tenant
# =========================
$TicketDefaults = @{
    tickettype_id              = 34
    priority_id                = 4
    sla_id                     = 1

    # Mandatory in your tenant
    category_1                 = "RMM Alert"
    impact                     = "3"
    urgency                    = "2"

    dont_do_rules              = $true
    donotapplytemplateintheapi = $true
    return_this                = $true
    utcoffset                  = -600
}

# =========================
# RESOLVE DEFAULTS (matches your working payload pattern)
# =========================
$ResolveDefaults = @{
    # Required resolution mechanics
    outcome_id          = "4"
    new_status          = 9
    new_matched_rule_id = 8
    new_rule_ids        = "8"
    new_slastatus       = -1

    # Charge Type: Automation/AI (your requirement)
    chargerate          = "8"
    travel_chargerate   = 3

    # Common “defaults” your working payload includes
    dont_do_rules       = $true
    timerinuse          = $false
    from_mailbox_id     = 0
    follow              = $false
    hiddenfromuser      = $false
    important           = $false
    sendemail           = $false
    sendsms             = $false
    send_survey         = $false
    run_ai_insights     = $false
    sync_to_halo_api    = 0
    utcoffset           = -600

    # These flags appear in your working payloads
    _validate_travel    = $true
    _ignore_ai          = $false
    _sendtweet          = $false

    # Optional but harmless (kept consistent with your payload)
    emailcc             = ""
    emailbcc            = ""
    smsto               = ""
    emailsubject        = $null
    emailto             = $null
    files               = $null
    attachments         = @()
}

# =========================
# ENDPOINTS (from your Postman collection)
# =========================
$TokenUrl        = if ([string]::IsNullOrWhiteSpace($Tenant)) { "$BaseHost/auth/token" } else { "$BaseHost/auth/token?tenant=$Tenant" }  # 
$CreateTicketUrl = "$BaseHost/api/tickets"   # 
$CreateActionUrl = "$BaseHost/api/Actions"

function Get-DetailsHtml([string]$plainText) {
    # Correct: encode just the text and wrap in entity-encoded <p> tags like your payload
    $encodedText = [System.Web.HttpUtility]::HtmlEncode($plainText)
    return "&lt;p&gt;$encodedText&lt;/p&gt;"
}

function Get-NoteHtml([string]$plainText) {
    # Same pattern as your manual payload: &lt;p&gt;...&lt;/p&gt;
    $encodedText = [System.Web.HttpUtility]::HtmlEncode($plainText)
    return "&lt;p&gt;$encodedText&lt;/p&gt;"
}

function Get-BearerToken {
    if ([string]::IsNullOrWhiteSpace($OAuthClientId) -or [string]::IsNullOrWhiteSpace($OAuthClientSecret)) {
        throw "OAuthClientId/OAuthClientSecret are not set in the script."
    }

    $body = @{
        grant_type    = "client_credentials"
        client_id     = $OAuthClientId
        client_secret = $OAuthClientSecret
        scope         = "all"
    }

    $resp = Invoke-RestMethod -Method Post -Uri $TokenUrl -ContentType "application/x-www-form-urlencoded" -Body $body
    if (-not $resp.access_token) { throw "Token response missing access_token." }
    return $resp.access_token
}

function Invoke-HaloWeb {
    param(
        [Parameter(Mandatory=$true)] [string] $Uri,
        [Parameter(Mandatory=$true)] [string] $Method,
        [Parameter(Mandatory=$true)] [string] $JsonBody,
        [Parameter(Mandatory=$true)] [hashtable] $Headers
    )

    # Use Invoke-WebRequest to reliably capture response bodies on failure
    if ($PSVersionTable.PSVersion.Major -ge 7) {
        $status = $null
        $wr = Invoke-WebRequest -Method $Method -Uri $Uri -Headers $Headers -Body $JsonBody `
            -ContentType "application/json" -SkipHttpErrorCheck -StatusCodeVariable status
        return [pscustomobject]@{
            StatusCode = $status
            RawContent = $wr.Content
        }
    } else {
        try {
            $wr = Invoke-WebRequest -Method $Method -Uri $Uri -Headers $Headers -Body $JsonBody -ContentType "application/json"
            return [pscustomobject]@{
                StatusCode = $wr.StatusCode
                RawContent = $wr.Content
            }
        } catch {
            $statusCode = $null
            $raw = $null
            try {
                $r = $_.Exception.Response
                if ($r) {
                    $statusCode = [int]$r.StatusCode
                    $reader = New-Object System.IO.StreamReader($r.GetResponseStream())
                    $raw = $reader.ReadToEnd()
                }
            } catch {}
            return [pscustomobject]@{
                StatusCode = $statusCode
                RawContent = $raw
            }
        }
    }
}

# =========================
# AUTH
# =========================
$BearerToken = Get-BearerToken
$Headers = @{
    "Authorization" = "Bearer $BearerToken"
    "Accept"        = "application/json"
    "Content-Type"  = "application/json"
}

# =========================
# 1) CREATE TICKET (MINIMUM REQUIRED FIELDS)
# =========================
$ticketObj = @{
    client_id    = $CustomerId
    summary      = $Summary
    details_html = (Get-DetailsHtml $Detail)
} + $TicketDefaults

# MUST be array payload for /api/tickets 
$ticketJson = ConvertTo-Json -InputObject @($ticketObj) -Depth 50

$create = Invoke-HaloWeb -Uri $CreateTicketUrl -Method "POST" -JsonBody $ticketJson -Headers $Headers
Write-Host "CREATE /api/tickets Status:" $create.StatusCode

if ($create.StatusCode -lt 200 -or $create.StatusCode -ge 300) {
    Write-Host "CREATE Response Body:"
    Write-Host $create.RawContent
    if ($DebugRequests) {
        Write-Host "`nCREATE Request JSON:"
        Write-Host $ticketJson
    }
    throw "Create ticket failed (HTTP $($create.StatusCode))"
}

$ticketResp = $create.RawContent | ConvertFrom-Json
$ticketId = $null

if ($ticketResp -is [System.Array] -and $ticketResp.Count -gt 0) {
    $ticketId = $ticketResp[0].id
    if (-not $ticketId) { $ticketId = $ticketResp[0].ticket_id }
    if (-not $ticketId) { $ticketId = $ticketResp[0].faultid }
} else {
    $ticketId = $ticketResp.id
    if (-not $ticketId) { $ticketId = $ticketResp.ticket_id }
    if (-not $ticketId) { $ticketId = $ticketResp.faultid }
}

if (-not $ticketId) {
    if ($DebugRequests) {
        Write-Host "`nCREATE Raw Response:"
        Write-Host $create.RawContent
    }
    throw "Ticket created but could not determine Ticket ID from response."
}

Write-Host "Created ticket ID:" $ticketId

# =========================
# 2) RESOLVE TICKET (MINIMUM REQUIRED FIELDS + CHARGE TYPE)
# =========================
$timeHours = [Math]::Round(($TimeMinutes / 60.0), 6)

$actionObj = @{
    ticket_id    = "$ticketId"
    itsm_summary = $Summary
    note_html    = (Get-NoteHtml $ResolutionSummary)
    timetaken    = $timeHours
} + $ResolveDefaults

$actionJson = ConvertTo-Json -InputObject @($actionObj) -Depth 50

$resolve = Invoke-HaloWeb -Uri $CreateActionUrl -Method "POST" -JsonBody $actionJson -Headers $Headers
Write-Host "POST /api/Actions Status:" $resolve.StatusCode

if ($resolve.StatusCode -lt 200 -or $resolve.StatusCode -ge 300) {
    Write-Host "RESOLVE Response Body:"
    Write-Host $resolve.RawContent
    if ($DebugRequests) {
        Write-Host "`nRESOLVE Request JSON:"
        Write-Host $actionJson
    }
    throw "Resolve action failed (HTTP $($resolve.StatusCode))"
}

Write-Host "Resolved ticket ID: $ticketId (TimeMinutes=$TimeMinutes => timetaken=$timeHours hours; ChargeRate=8 Automation/AI)"