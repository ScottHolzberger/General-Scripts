#Requires -Version 5.1
<#
Export-Halo-RecommendationData.ps1 (v6)

Exports HaloPSA ticket + knowledge-related data to recommend:
- Ticket category rationalisation
- KB articles to reduce ticket volume

Compatibility:
- Windows PowerShell 5.1 compatible (no ??, no ?., no PS7-only syntax)
- Uses SecureString for ClientSecret

Behaviour:
- Tickets: GET /Tickets wrapper { record_count, tickets, include_children } -> unwrap .tickets
- FAQ Lists: exports GET /FAQLists (knowledge structure)
- KB Entries: tries GET /KBEntry; if 404, continues (creates KB recommendations from ticket trends anyway)

Outputs (in OutDir):
- tickets.raw.jsonl
- tickets.summary.csv
- tickets.category-usage.csv
- tickets.top-summaries.csv
- faqlists.raw.jsonl
- faqlists.summary.csv
- kb.raw.jsonl (only if KBEntry works)
- kb.summary.csv (only if KBEntry works)
- kb.endpoint.notfound.txt (if KBEntry 404)
- kb.recommendations.csv
- tickets.schema.fields.txt (if -DumpTicketSchema)
#>

[CmdletBinding()]
param(
  [Parameter(Mandatory=$false)]
  [string]$ApiBase = "https://zahezone.halopsa.com/api",

  [Parameter(Mandatory=$false)]
  [string]$AuthUrl = "https://zahezone.halopsa.com/auth/token",

  [Parameter(Mandatory=$true)]
  [string]$ClientId,

  [Parameter(Mandatory=$true)]
  [SecureString]$ClientSecret,

  [Parameter(Mandatory=$false)]
  [string]$Scope = "all",

  [Parameter(Mandatory=$false)]
  [int]$LookbackDays = 90,

  [Parameter(Mandatory=$false)]
  [string]$OutDir = (Join-Path -Path (Get-Location) -ChildPath ("halo-export_" + (Get-Date -Format "yyyyMMdd_HHmmss"))),

  [Parameter(Mandatory=$false)]
  [int]$PageSize = 500,

  [Parameter(Mandatory=$false)]
  [switch]$DebugResponseShape,

  [Parameter(Mandatory=$false)]
  [switch]$DumpTicketSchema
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ----------------
# Helper functions
# ----------------
function Ensure-Tls12 {
  try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }
}

function New-Dir([string]$Path) {
  if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path | Out-Null }
}

function ConvertFrom-SecureStringToPlain([SecureString]$Secure) {
  $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
  try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
  finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

function Invoke-HaloAuthToken {
  param(
    [string]$AuthUrl,
    [string]$ClientId,
    [SecureString]$ClientSecret,
    [string]$Scope
  )

  $plainSecret = ConvertFrom-SecureStringToPlain $ClientSecret

  $body = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $plainSecret
    scope         = $Scope
  }

  $resp = Invoke-RestMethod -Method Post -Uri $AuthUrl -Body $body -ContentType "application/x-www-form-urlencoded"

  if (-not ($resp -and ($resp.PSObject.Properties.Name -contains "access_token") -and $resp.access_token)) {
    throw "Auth failed or access_token missing."
  }
  return $resp.access_token
}

function Invoke-HaloGet {
  param(
    [string]$Url,
    [string]$Token,
    [hashtable]$Query = $null
  )

  if ($Query -and $Query.Count -gt 0) {
    $pairs = foreach ($kv in $Query.GetEnumerator()) {
      $k = [System.Uri]::EscapeDataString([string]$kv.Key)
      $v = [System.Uri]::EscapeDataString([string]$kv.Value)
      "$k=$v"
    }
    $qs = ($pairs -join "&")
    if ($Url -notmatch "\?") { $Url = "$Url`?$qs" } else { $Url = "$Url`&$qs" }
  }

  $headers = @{
    Authorization = "Bearer $Token"
    Accept        = "application/json"
  }

  return Invoke-RestMethod -Method Get -Uri $Url -Headers $headers
}

function Get-HttpStatusCodeFromError($_err) {
  try {
    if ($_err -and $_err.Exception -and $_err.Exception.Response) {
      return [int]$_.Exception.Response.StatusCode
    }
  } catch { }
  return $null
}

function Unwrap-HaloListResponse {
  <#
    ALWAYS returns an array so .Count is safe under StrictMode.
    Handles your tickets wrapper (.tickets) + common list wrappers.
  #>
  param([object]$Resp)

  if ($null -eq $Resp) { return @() }

  if ($Resp -is [System.Array]) { return @($Resp) }

  $props = @($Resp.PSObject.Properties.Name)

  if ($props -contains "tickets") { return @($Resp.tickets) }
  if ($props -contains "records") { return @($Resp.records) }
  if ($props -contains "items")   { return @($Resp.items) }
  if ($props -contains "data")    { return @($Resp.data) }

  return @($Resp)
}

function Get-PropValue {
  <#
    StrictMode-safe getter.
    Returns the first matching existing property value (or $null).
  #>
  param(
    [Parameter(Mandatory=$true)][object]$Obj,
    [Parameter(Mandatory=$true)][string[]]$Names
  )

  if ($null -eq $Obj) { return $null }
  $props = $Obj.PSObject.Properties
  foreach ($n in $Names) {
    $p = $props[$n]
    if ($p) { return $p.Value }
  }
  return $null
}

function Normalize-Text([string]$s) {
  if ([string]::IsNullOrWhiteSpace($s)) { return "" }
  $s = $s -replace "<[^>]+>", " "
  $s = $s -replace "\s+", " "
  return $s.Trim()
}

function Write-Jsonl {
  param([string]$Path, [object[]]$Items)
  foreach ($i in $Items) {
    ($i | ConvertTo-Json -Depth 80 -Compress) | Add-Content -Path $Path -Encoding UTF8
  }
}

function Get-ResponseShapeString([object]$Resp) {
  if ($null -eq $Resp) { return "null" }
  $type = $Resp.GetType().FullName
  $props = @($Resp.PSObject.Properties.Name)
  return ("Type={0}; Props={1}" -f $type, ($props -join ", "))
}

function Guess-KBTopic {
  param([string]$Summary)

  # PowerShell 5.1-safe null handling
  $safe = ""
  if ($null -ne $Summary) { $safe = [string]$Summary }
  $s = $safe.ToLowerInvariant()

  $tags = New-Object System.Collections.Generic.List[string]

  if ($s -match "password|reset|mfa|auth|login|sign") { $tags.Add("authentication") }
  if ($s -match "outlook|email|mailbox")               { $tags.Add("email") }
  if ($s -match "onedrive|sharepoint|teams")           { $tags.Add("m365") }
  if ($s -match "vpn|firewall|dns|wifi|wireless|internet") { $tags.Add("network") }
  if ($s -match "printer|print|scan")                  { $tags.Add("printing") }
  if ($s -match "3cx|voip|phone|yealink|poly")         { $tags.Add("telecoms") }

  $title = $safe
  if ($title.Length -gt 90) { $title = $title.Substring(0, 90).Trim() + "…" }
  if ([string]::IsNullOrWhiteSpace($title)) { $title = "Common support issue" }

  return [pscustomobject]@{
    suggested_title = ("How to: {0}" -f $title)
    suggested_tags  = (($tags | Select-Object -Unique) -join "; ")
  }
}

function Test-HaloEndpoint404 {
  <#
    Returns $true if endpoint appears to exist, $false if definite 404.
    (Some endpoints may 400/401 for bad params/permissions; those are treated as "exists".)
  #>
  param(
    [string]$ApiBase,
    [string]$Endpoint,
    [string]$Token
  )

  $testUrl = "$ApiBase/$Endpoint"
  try {
    $null = Invoke-HaloGet -Url $testUrl -Token $Token -Query @{ count = 1; page = 1 }
    return $true
  } catch {
    $code = $null
    try {
      if ($_.Exception -and $_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode }
    } catch { }
    if ($code -eq 404) { return $false }
    return $true
  }
}

# -----------
# Start script
# -----------
Ensure-Tls12
New-Dir $OutDir
Write-Host "Output: $OutDir"

$token = Invoke-HaloAuthToken -AuthUrl $AuthUrl -ClientId $ClientId -ClientSecret $ClientSecret -Scope $Scope

$start = (Get-Date).AddDays(-1 * $LookbackDays)
$startIso = $start.ToString("yyyy-MM-dd")
$endIso   = (Get-Date).ToString("yyyy-MM-dd")

# -----------------------------
# Export Tickets (GET /Tickets)
# -----------------------------
$ticketsRawPath = Join-Path $OutDir "tickets.raw.jsonl"
$ticketsCsvPath = Join-Path $OutDir "tickets.summary.csv"

$tickets = New-Object System.Collections.Generic.List[object]
$page = 1
$printedTicketShape = $false
$schemaDumped = $false

Write-Host "Exporting tickets (last $LookbackDays days)..."

while ($true) {
  $query = @{
    count        = $PageSize
    page         = $page
    datesearch   = "dateoccured"
    startdate    = $startIso
    enddate      = $endIso
    includeagent = "true"
    includeuser  = "true"
  }

  $url = "$ApiBase/Tickets"
  $resp = Invoke-HaloGet -Url $url -Token $token -Query $query

  if ($DebugResponseShape -and -not $printedTicketShape) {
    Write-Host ("Tickets response shape: " + (Get-ResponseShapeString $resp))
    $printedTicketShape = $true
  }

  $batch = Unwrap-HaloListResponse -Resp $resp
  if ($batch.Count -eq 0) { break }

  Write-Host ("  page {0}: {1} tickets" -f $page, $batch.Count)
  $tickets.AddRange($batch)
  Write-Jsonl -Path $ticketsRawPath -Items $batch

  if ($DumpTicketSchema -and -not $schemaDumped -and $batch.Count -gt 0) {
    $schemaPath = Join-Path $OutDir "tickets.schema.fields.txt"
    @($batch[0].PSObject.Properties.Name) | Sort-Object | Set-Content -Encoding UTF8 -Path $schemaPath
    Write-Host "Wrote ticket field list to: $schemaPath"
    $schemaDumped = $true
  }

  if ($batch.Count -lt $PageSize) { break }
  $page++
}

Write-Host ("Total tickets exported: {0}" -f $tickets.Count)

$ticketRows = $tickets | ForEach-Object {
  $t = $_
  [pscustomobject]@{
    id           = Get-PropValue $t @("id","fault_id","ticket_id")
    summary      = Normalize-Text (Get-PropValue $t @("summary","title","subject"))
    details      = Normalize-Text (Get-PropValue $t @("details","description","body"))
    client_id    = Get-PropValue $t @("client_id","area_id","organisation_id")
    client_name  = Get-PropValue $t @("client_name","area_name","organisation_name")
    site_id      = Get-PropValue $t @("site_id")
    site_name    = Get-PropValue $t @("site_name")
    tickettype_id= Get-PropValue $t @("tickettype_id","ticket_type_id","requesttype_id")
    status_id    = Get-PropValue $t @("status_id","statusid")
    outcome      = Get-PropValue $t @("outcome","outcome_name","resolution")
    priority     = Get-PropValue $t @("priority","priority_name","priority_id")
    category_1   = Get-PropValue $t @("category_1","category1","cat1")
    category_2   = Get-PropValue $t @("category_2","category2","cat2")
    category_3   = Get-PropValue $t @("category_3","category3","cat3")
    category_4   = Get-PropValue $t @("category_4","category4","cat4")
    dateoccurred = Get-PropValue $t @("dateoccurred","date_occurred","opened","created","datecreated")
    datecleared  = Get-PropValue $t @("datecleared","date_cleared","closed","dateclosed")
    agent_id     = Get-PropValue $t @("agent_id","assigned_agent_id")
    agent_name   = Get-PropValue $t @("agent_name","assigned_agent_name","agent")
    user_id      = Get-PropValue $t @("user_id","requester_id","contact_id")
    user_name    = Get-PropValue $t @("user_name","requester_name","contact_name","user")
    source       = Get-PropValue $t @("source","channel","origin")
  }
}

$ticketRows | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $ticketsCsvPath

$catUsagePath = Join-Path $OutDir "tickets.category-usage.csv"
$ticketRows |
  Group-Object -Property category_1,category_2,category_3,category_4 |
  Sort-Object Count -Descending |
  ForEach-Object {
    $n = $_.Name -split ","
    [pscustomobject]@{
      count      = $_.Count
      category_1 = ($n[0] -replace '^"|"$','')
      category_2 = ($n[1] -replace '^"|"$','')
      category_3 = ($n[2] -replace '^"|"$','')
      category_4 = ($n[3] -replace '^"|"$','')
    }
  } | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $catUsagePath

$topSummariesPath = Join-Path $OutDir "tickets.top-summaries.csv"
$ticketRows |
  Where-Object { -not [string]::IsNullOrWhiteSpace($_.summary) } |
  Group-Object -Property summary |
  Sort-Object Count -Descending |
  Select-Object -First 200 |
  ForEach-Object {
    [pscustomobject]@{ count = $_.Count; summary = $_.Name }
  } | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $topSummariesPath

# -----------------------------
# Export FAQ Lists (GET /FAQLists)
# -----------------------------
$faqRawPath = Join-Path $OutDir "faqlists.raw.jsonl"
$faqCsvPath = Join-Path $OutDir "faqlists.summary.csv"

Write-Host "Exporting FAQ Lists (knowledge structure)..."

try {
  $faqResp = Invoke-HaloGet -Url "$ApiBase/FAQLists" -Token $token -Query @{ showcounts = "true" }
  $faqBatch = Unwrap-HaloListResponse -Resp $faqResp

  Write-Jsonl -Path $faqRawPath -Items $faqBatch

  $faqBatch | ForEach-Object {
    $f = $_
    [pscustomobject]@{
      id              = Get-PropValue $f @("id")
      name            = Get-PropValue $f @("name","title")
      parent_id       = Get-PropValue $f @("parent_id")
      level           = Get-PropValue $f @("level")
      type            = Get-PropValue $f @("type")
      item_count      = Get-PropValue $f @("count","item_count","article_count","showcounts")
      organisation_id = Get-PropValue $f @("organisation_id")
    }
  } | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $faqCsvPath

  Write-Host ("FAQ Lists exported: {0}" -f $faqBatch.Count)
} catch {
  $code = $null
  try { if ($_.Exception -and $_.Exception.Response) { $code = [int]$_.Exception.Response.StatusCode } } catch { }
  $note = ("FAQLists export failed. HTTP={0}. {1}" -f $code, $_.Exception.Message)
  Set-Content -Encoding UTF8 -Path (Join-Path $OutDir "faqlists.error.txt") -Value $note
  Write-Host $note
}

# -----------------------------
# Export KB Entries (GET /KBEntry) if available
# -----------------------------
$kbRawPath = Join-Path $OutDir "kb.raw.jsonl"
$kbCsvPath = Join-Path $OutDir "kb.summary.csv"

$kbAvailable = Test-HaloEndpoint404 -ApiBase $ApiBase -Endpoint "KBEntry" -Token $token

if ($kbAvailable) {
  Write-Host "Exporting KB entries (KBEntry)..."
  $kb = New-Object System.Collections.Generic.List[object]
  $page = 1
  $printedKbShape = $false

  while ($true) {
    $query = @{
      count          = $PageSize
      page           = $page
      paginate       = "true"
      includedetails = "false"
    }

    $resp = Invoke-HaloGet -Url "$ApiBase/KBEntry" -Token $token -Query $query

    if ($DebugResponseShape -and -not $printedKbShape) {
      Write-Host ("KB response shape: " + (Get-ResponseShapeString $resp))
      $printedKbShape = $true
    }

    $batch = Unwrap-HaloListResponse -Resp $resp
    if ($batch.Count -eq 0) { break }

    Write-Host ("  page {0}: {1} KB entries" -f $page, $batch.Count)
    $kb.AddRange($batch)
    Write-Jsonl -Path $kbRawPath -Items $batch

    if ($batch.Count -lt $PageSize) { break }
    $page++
  }

  $kbRows = $kb | ForEach-Object {
    $k = $_
    [pscustomobject]@{
      id      = Get-PropValue $k @("id")
      title   = Normalize-Text (Get-PropValue $k @("title","name"))
      summary = Normalize-Text (Get-PropValue $k @("summary","description"))
      tags    = Normalize-Text (Get-PropValue $k @("tags","tag_string"))
      type    = Get-PropValue $k @("type")
      updated = Get-PropValue $k @("datelastmodified","date_edited","updated")
    }
  }

  $kbRows | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $kbCsvPath
  Write-Host ("KB entries exported: {0}" -f $kbRows.Count)
}
else {
  $note = "KBEntry endpoint returned 404 in this tenant. Skipping existing KB export."
  Set-Content -Encoding UTF8 -Path (Join-Path $OutDir "kb.endpoint.notfound.txt") -Value $note
  Write-Host $note
}

# -----------------------------
# KB recommendations from tickets (always generated)
# -----------------------------
$kbRecsPath = Join-Path $OutDir "kb.recommendations.csv"

$top = Import-Csv $topSummariesPath

$top | ForEach-Object {
  $g = Guess-KBTopic -Summary $_.summary
  [pscustomobject]@{
    count           = [int]$_.count
    ticket_summary  = $_.summary
    suggested_title = $g.suggested_title
    suggested_tags  = $g.suggested_tags
  }
} | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $kbRecsPath

Write-Host "Done."
Write-Host "Key outputs:"
Write-Host " - $ticketsCsvPath"
Write-Host " - $catUsagePath"
Write-Host " - $topSummariesPath"
Write-Host " - $faqCsvPath"
Write-Host " - $kbCsvPath"
Write-Host " - $kbRecsPath"