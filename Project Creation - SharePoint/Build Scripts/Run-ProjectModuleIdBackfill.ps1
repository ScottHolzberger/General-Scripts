Import-Module PnP.PowerShell -Force
Disconnect-PnPOnline -ErrorAction SilentlyContinue

# ===== AUTH (APP-ONLY CERT) =====
$SiteUrl  = "https://zahe.sharepoint.com/sites/Projects"
$Tenant   = "zahe.onmicrosoft.com"
$ClientId = "01e1b71f-cbcb-48df-a076-871aa4ba10d9"

$PfxPath = ".\ZaheZone-PnP-Projects.pfx"
$PfxPwd  = ConvertTo-SecureString "UseA-LongRandomPasswordHere" -AsPlainText -Force

Connect-PnPOnline -Url $SiteUrl -Tenant $Tenant -ClientId $ClientId -CertificatePath $PfxPath -CertificatePassword $PfxPwd

# ===== LISTS =====
$ProjectsList = "Project Register"
$ModulesList  = "Project Modules"

function Pad3([int]$n) { return ("{0:000}" -f $n) }

# -------- 1) Build Project map (SharePoint Item ID -> ProjectID) --------
$projectItems = Get-PnPListItem -List $ProjectsList -PageSize 2000 -Fields "ID","ProjectID"
$projectIdByItemId = @{}

foreach ($p in $projectItems) {
  $spId = [int]$p.Id
  $projId = [string]$p["ProjectID"]

  if ([string]::IsNullOrWhiteSpace($projId)) {
    $projId = "PRJ-" + (Pad3 $spId)
    Set-PnPListItem -List $ProjectsList -Identity $spId -Values @{ "ProjectID" = $projId } | Out-Null  # [8](https://pnp.github.io/powershell/cmdlets/Set-PnPListItem.html)
  }

  $projectIdByItemId[$spId] = $projId
}

# -------- 2) Load Modules, compute sequences per ParentProject --------
$moduleItems = Get-PnPListItem -List $ModulesList -PageSize 2000 -Fields "ID","ParentProject","ModuleSequence","ModuleID"

# Track max ModuleSequence for each ParentProject item id
$maxSeqByParentItemId = @{}

# First pass: determine max sequence for each parent
foreach ($m in $moduleItems) {
  $parent = $m["ParentProject"]
  if ($null -eq $parent) { continue } # ParentProject is required, but skip old bad data if any

  # ParentProject is a lookup -> FieldLookupValue
  $parentItemId = $parent.LookupId

  $seq = $m["ModuleSequence"]
  $mid = [string]$m["ModuleID"]

  # If ModuleSequence empty but ModuleID exists, attempt derive last 3 digits
  if (($null -eq $seq -or $seq -eq "") -and -not [string]::IsNullOrWhiteSpace($mid)) {
    $last = $mid.Split('-')[-1]
    if ($last -match '^\d+$') { $seq = [int]$last }
  }

  if ($seq -is [string] -and $seq -match '^\d+$') { $seq = [int]$seq }

  if ($seq -is [int]) {
    if (-not $maxSeqByParentItemId.ContainsKey($parentItemId)) {
      $maxSeqByParentItemId[$parentItemId] = $seq
    } else {
      if ($seq -gt $maxSeqByParentItemId[$parentItemId]) { $maxSeqByParentItemId[$parentItemId] = $seq }
    }
  } else {
    if (-not $maxSeqByParentItemId.ContainsKey($parentItemId)) { $maxSeqByParentItemId[$parentItemId] = 0 }
  }
}

# Second pass: assign missing sequences and ModuleIDs
foreach ($m in $moduleItems) {
  $moduleSpId = [int]$m.Id
  $parent = $m["ParentProject"]
  if ($null -eq $parent) { continue }  # cannot assign without parent

  $parentItemId = $parent.LookupId
  $projectId = $projectIdByItemId[$parentItemId]

  $seq = $m["ModuleSequence"]
  $moduleId = [string]$m["ModuleID"]

  $needsSeq = ($null -eq $seq -or $seq -eq "")
  $needsId  = [string]::IsNullOrWhiteSpace($moduleId)

  if ($needsSeq) {
    $maxSeqByParentItemId[$parentItemId] = [int]$maxSeqByParentItemId[$parentItemId] + 1
    $seq = $maxSeqByParentItemId[$parentItemId]
  }

  if ($needsId) {
    $moduleId = "$projectId-" + (Pad3 ([int]$seq))
  }

  # Update item if anything changed
  Set-PnPListItem -List $ModulesList -Identity $moduleSpId -Values @{
    "ModuleSequence" = [double]$seq
    "ModuleID"       = $moduleId
  } | Out-Null  # [8](https://pnp.github.io/powershell/cmdlets/Set-PnPListItem.html)
}

Write-Host "✅ Backfill complete: ProjectID + ModuleSequence + ModuleID populated where missing."