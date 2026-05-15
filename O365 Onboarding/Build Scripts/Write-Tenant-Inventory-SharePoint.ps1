param(
  [Parameter(Mandatory)][string]$SiteUrl,
  [Parameter(Mandatory)][string]$Tenant,
  [Parameter(Mandatory)][string]$ClientId,
  [Parameter(Mandatory)][string]$PfxPath,
  [Parameter(Mandatory)][string]$PfxPasswordPlain,

  [Parameter(Mandatory)][string]$ClientTenancyTitle,
  [Parameter(Mandatory)][string]$InventoryJsonPath
)

$ErrorActionPreference = "Stop"
Import-Module PnP.PowerShell -ErrorAction Stop

Connect-PnPOnline `
  -Url $SiteUrl `
  -Tenant $Tenant `
  -ClientId $ClientId `
  -CertificatePath $PfxPath `
  -CertificatePassword (ConvertTo-SecureString $PfxPasswordPlain -AsPlainText -Force)

# -------------------------------
# Helper: resolve internal field name safely
# -------------------------------
function Resolve-FieldName {
  param(
    [Parameter(Mandatory)][string]$List,
    [Parameter(Mandatory)][string[]]$Candidates
  )
  foreach ($c in $Candidates) {
    $f = Get-PnPField -List $List -Identity $c -ErrorAction SilentlyContinue
    if ($f) { return $c }
  }
  return $null
}

# -------------------------------
# Load Graph payload
# -------------------------------
$payload = Get-Content $InventoryJsonPath -Raw | ConvertFrom-Json

# Normalize “unknown” states to be explicit rather than blank
if ($null -eq $payload.SecurityDefaultsEnabled) { $payload.SecurityDefaultsEnabled = $null }
if ($null -eq $payload.SSPREnabled) { $payload.SSPREnabled = $null }
if (-not $payload.DKIMStatus) { $payload | Add-Member -NotePropertyName DKIMStatus -NotePropertyValue "Unknown" -Force }

# -------------------------------
# Find Client Tenancy list item
# -------------------------------
$clientItem = Get-PnPListItem -List "Client Tenancies" -Query @"
<View><Query><Where>
  <Eq><FieldRef Name='Title'/><Value Type='Text'>$ClientTenancyTitle</Value></Eq>
</Where></Query></View>
"@

if (-not $clientItem) {
  throw "Client Tenancy '$ClientTenancyTitle' not found in 'Client Tenancies' list."
}

# -------------------------------
# Resolve internal field names (lookup + core fields)
# -------------------------------
$InvList = "Tenant Inventory"
$RecList = "Recommendations Register"

$Inv_ClientLookup = Resolve-FieldName -List $InvList -Candidates @("Client_x0020_Tenancy","ClientTenancy")
$Inv_InventoryDate= Resolve-FieldName -List $InvList -Candidates @("InventoryDate","Inventory_x0020_Date")
$Inv_SecDefaults  = Resolve-FieldName -List $InvList -Candidates @("SecurityDefaultsEnabled","SecurityDefaults")
$Inv_SSPR         = Resolve-FieldName -List $InvList -Candidates @("SSPREnabled","SSPR")
$Inv_DKIM         = Resolve-FieldName -List $InvList -Candidates @("DKIMStatus","DKIM")

if (-not $Inv_ClientLookup) { throw "Tenant Inventory lookup field not found (expected Client_x0020_Tenancy)." }
if (-not $Inv_InventoryDate){ throw "Tenant Inventory field InventoryDate not found." }

$Rec_ClientLookup = Resolve-FieldName -List $RecList -Candidates @("Client_x0020_Tenancy","ClientTenancy")
$Rec_InvLookup    = Resolve-FieldName -List $RecList -Candidates @("Inventory_x0020_Reference","InventoryReference")
$Rec_Category      = Resolve-FieldName -List $RecList -Candidates @("Category")
$Rec_Severity      = Resolve-FieldName -List $RecList -Candidates @("Severity")
$Rec_Status        = Resolve-FieldName -List $RecList -Candidates @("RecommendationStatus","Status")
$Rec_ServiceTier   = Resolve-FieldName -List $RecList -Candidates @("ServiceTier")
$Rec_AutoApply     = Resolve-FieldName -List $RecList -Candidates @("AutoApplyEligible")
$Rec_Source        = Resolve-FieldName -List $RecList -Candidates @("Source")

# -------------------------------
# Write Inventory row
# -------------------------------
$invTitle = "Inventory - " + (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
$invValues = @{
  "Title"                 = $invTitle
  $Inv_ClientLookup       = $clientItem.Id
  $Inv_InventoryDate      = (Get-Date)
}

if ($Inv_SecDefaults) { $invValues[$Inv_SecDefaults] = $payload.SecurityDefaultsEnabled }
if ($Inv_SSPR)        { $invValues[$Inv_SSPR]        = $payload.SSPREnabled }
if ($Inv_DKIM)        { $invValues[$Inv_DKIM]        = $payload.DKIMStatus }

$invItem = Add-PnPListItem -List $InvList -Values $invValues
Write-Host "✅ Inventory record written: $invTitle (ID=$($invItem.Id))"

# ==========================================================
# RULE ENGINE (Scalable “B”)
# Aligned to PRJ-004-001 governance model
# ==========================================================

function New-ZZRule {
  param(
    [Parameter(Mandatory)][string]$RuleId,
    [Parameter(Mandatory)][string]$Title,
    [Parameter(Mandatory)][string]$Category,   # Identity/Email/Security/Licensing/Governance
    [Parameter(Mandatory)][scriptblock]$When,
    [Parameter(Mandatory)][hashtable]$Defaults
  )
  [pscustomobject]@{ RuleId=$RuleId; Title=$Title; Category=$Category; When=$When; Defaults=$Defaults }
}

function Get-ZZRulePack {
  $rules = @()

  # SEC-0001: Security Defaults disabled -> High, Included, Auto-apply allowed (per your model) [3](https://outlook.office365.com/owa/?ItemID=AAMkADhlMTM2M2M5LTMwMDAtNDhmMy05NWEzLTdiNmE3M2FhZjc4MQBGAAAAAADb8e2aAMxXSpYvmbR2C5nJBwDMTeUeLbukSoYFIhN4KOryAAIgy27eAADMTeUeLbukSoYFIhN4KOryAALm6xaYAAA%3d&exvsurl=1&viewmodel=ReadMessageItem)
  $rules += New-ZZRule -RuleId "SEC-0001" -Title "Enable Security Defaults" -Category "Security" -When {
    param($inv) $inv.SecurityDefaultsEnabled -eq $false
  } -Defaults @{
    Severity="High"; AutoApplyEligible=$true; ServiceTier="Included"; RecommendationStatus="Deferred"; Source="Onboarding"
  }

  # ID-0002: SSPR unknown -> Medium review
  $rules += New-ZZRule -RuleId "ID-0002" -Title "Review SSPR Configuration" -Category "Identity" -When {
    param($inv) $null -eq $inv.SSPREnabled
  } -Defaults @{
    Severity="Medium"; AutoApplyEligible=$false; ServiceTier="Included"; RecommendationStatus="Deferred"; Source="Onboarding"
  }

  # ID-0003: SSPR disabled -> Medium
  $rules += New-ZZRule -RuleId "ID-0003" -Title "Enable Self-Service Password Reset (SSPR)" -Category "Identity" -When {
    param($inv) $inv.SSPREnabled -eq $false
  } -Defaults @{
    Severity="Medium"; AutoApplyEligible=$false; ServiceTier="Included"; RecommendationStatus="Deferred"; Source="Onboarding"
  }

  # MAIL-0004: DKIM Unknown/Disabled -> Medium
  $rules += New-ZZRule -RuleId "MAIL-0004" -Title "Review and Enable DKIM" -Category "Email" -When {
    param($inv) ($inv.DKIMStatus -eq "Unknown") -or ($inv.DKIMStatus -eq "Disabled")
  } -Defaults @{
    Severity="Medium"; AutoApplyEligible=$false; ServiceTier="Included"; RecommendationStatus="Deferred"; Source="Onboarding"
  }

  # LIC-0005: No SKU data -> Low (informational)
  $rules += New-ZZRule -RuleId "LIC-0005" -Title "Review Tenant Licensing Inventory" -Category "Licensing" -When {
    param($inv) (-not $inv.SKUs) -or ($inv.SKUs.Count -eq 0)
  } -Defaults @{
    Severity="Low"; AutoApplyEligible=$false; ServiceTier="Included"; RecommendationStatus="Deferred"; Source="Onboarding"
  }

  return $rules
}

function Ensure-ZZRecommendation {
  param(
    [Parameter(Mandatory)][int]$ClientTenancyItemId,
    [Parameter(Mandatory)][int]$InventoryItemId,
    [Parameter(Mandatory)][pscustomobject]$Rule,
    [Parameter(Mandatory)][pscustomobject]$InventoryPayload
  )

  # Dedupe title embeds RuleId (no schema changes required)
  $dedupeTitle = "$($Rule.Title) [$($Rule.RuleId)]"

  $existing = Get-PnPListItem -List $RecList -PageSize 20 -Query @"
<View><Query><Where><And>
  <Eq><FieldRef Name='Title'/><Value Type='Text'>$dedupeTitle</Value></Eq>
  <Eq><FieldRef Name='$Rec_ClientLookup'/><Value Type='Lookup'>$ClientTenancyItemId</Value></Eq>
</And></Where></Query></View>
"@

  if ($existing) {
    Write-Host "  = Recommendation exists (skip): $dedupeTitle"
    return
  }

  $vals = @{
    "Title"                = $dedupeTitle
    $Rec_ClientLookup      = $ClientTenancyItemId
  }

  if ($Rec_InvLookup)    { $vals[$Rec_InvLookup] = $InventoryItemId }
  if ($Rec_Category)     { $vals[$Rec_Category]  = $Rule.Category }
  if ($Rec_Severity)     { $vals[$Rec_Severity]  = $Rule.Defaults.Severity }
  if ($Rec_Status)       { $vals[$Rec_Status]    = $Rule.Defaults.RecommendationStatus }
  if ($Rec_ServiceTier)  { $vals[$Rec_ServiceTier]= $Rule.Defaults.ServiceTier }
  if ($Rec_AutoApply)    { $vals[$Rec_AutoApply] = [bool]$Rule.Defaults.AutoApplyEligible }
  if ($Rec_Source)       { $vals[$Rec_Source]    = $Rule.Defaults.Source }

  Add-PnPListItem -List $RecList -Values $vals | Out-Null
  Write-Host "  + Recommendation created: $dedupeTitle"
}

function Invoke-ZZRuleEngine {
  param(
    [Parameter(Mandatory)][int]$ClientTenancyItemId,
    [Parameter(Mandatory)][int]$InventoryItemId,
    [Parameter(Mandatory)][pscustomobject]$InventoryPayload
  )

  $rules = Get-ZZRulePack
  foreach ($r in $rules) {
    $match = $false
    try { $match = & $r.When $InventoryPayload } catch { $match = $false }
    if ($match) {
      Ensure-ZZRecommendation -ClientTenancyItemId $ClientTenancyItemId -InventoryItemId $InventoryItemId -Rule $r -InventoryPayload $InventoryPayload
    }
  }
}

Invoke-ZZRuleEngine -ClientTenancyItemId $clientItem.Id -InventoryItemId $invItem.Id -InventoryPayload $payload
Write-Host "✅ Rule engine complete"