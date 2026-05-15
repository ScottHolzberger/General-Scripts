<#
Authoritative runtime document.
If it’s not described here, it’s not supported.

Update-All-Hubs.ps1 (FINAL)

Purpose:
- Project Hub: maintain Module Index (now hyperlinked to ModuleHub URL)
- Module Hub: maintain Parent Project backlink (hyperlinked to ProjectHub URL when available)

Design:
- Idempotent: removes only automation-owned blocks (by durable marker) before inserting
- Component-safe: only inspects/removes text components (ignores dividers/spacers/images)
- Full-width safe: inserts into Section 2 / Column 1 (ensures Section 2 exists)
- Does NOT touch scaffolds or runbook link blocks (those are Provision-Hubs responsibilities)
#>

[CmdletBinding()]
param(
  [switch]$WhatIf,
  [switch]$Enforce,

  # AUTH (App-only cert)
  [string]$SiteUrl  = "https://zahe.sharepoint.com/sites/Projects",
  [string]$Tenant   = "zahe.onmicrosoft.com",
  [string]$ClientId = "01e1b71f-cbcb-48df-a076-871aa4ba10d9",
  [string]$PfxPath  = ".\ZaheZone-PnP-Projects.pfx",
  [string]$PfxPasswordPlain = "UseA-LongRandomPasswordHere",

  # LISTS / FIELDS
  [string]$ProjectsList    = "Project Register",
  [string]$ModulesList     = "Project Modules",
  [string]$ProjectHubField = "ProjectHub",
  [string]$ModuleHubField  = "ModuleHub"
)

Set-StrictMode -Version Latest
Import-Module PnP.PowerShell -Force

# Durable visible markers (NOT HTML comments)
$MarkerModules = "ZZ-AUTO-MODULE-INDEX"
$MarkerParent  = "ZZ-AUTO-PARENT-PROJECT"

function Fail-Or-Warn([string]$msg) {
  if ($Enforce) { throw $msg }
  Write-Warning $msg
}

function Connect-Once {
  $pwd = ConvertTo-SecureString $PfxPasswordPlain -AsPlainText -Force
  Disconnect-PnPOnline -ErrorAction SilentlyContinue
  Connect-PnPOnline -Url $SiteUrl -Tenant $Tenant -ClientId $ClientId -CertificatePath $PfxPath -CertificatePassword $pwd
}

function Get-HubUrlFromFieldValue($FieldValue) {
  if ($null -eq $FieldValue) { return $null }
  if ($FieldValue -is [string]) { return $FieldValue }
  if ($FieldValue.PSObject.Properties.Name -contains "Url") { return $FieldValue.Url }
  return $null
}

function Resolve-PageIdentityFromUrl([string]$Url) {
  if ([string]::IsNullOrWhiteSpace($Url)) { return $null }

  $split = $Url -split "/SitePages/", 2
  if ($split.Count -lt 2) { return $null }

  $raw = $split[1]
  $decoded = [System.Uri]::UnescapeDataString($raw)
  $file = [System.IO.Path]::GetFileName($decoded)

  foreach ($c in @($decoded, $raw, $file) | Select-Object -Unique) {
    if ([string]::IsNullOrWhiteSpace($c)) { continue }
    if (Get-PnPPage -Identity $c -ErrorAction SilentlyContinue) { return $c }
  }
  return $null
}

function Component-Text([object]$component) {
  if ($component.PSObject.Properties.Name -contains "Text") { return [string]$component.Text }
  return ""
}

function Get-TextComponentsByMarker([string]$PageId, [string]$MarkerText) {
  $hits = @()
  foreach ($c in Get-PnPPageComponent -Page $PageId) {
    $t = Component-Text $c
    if (-not [string]::IsNullOrWhiteSpace($t) -and $t -like "*$MarkerText*") { $hits += $c }
  }
  return $hits
}

function Remove-TextComponentsByMarker([string]$PageId, [string]$MarkerText) {
  foreach ($c in Get-TextComponentsByMarker -PageId $PageId -MarkerText $MarkerText) {
    if ($WhatIf) {
      Write-Host "WHATIF: Would remove component $($c.InstanceId) on $PageId (marker=$MarkerText)" -ForegroundColor Yellow
    } else {
      Remove-PnPPageComponent -Page $PageId -InstanceId $c.InstanceId -Force | Out-Null
    }
  }
}

function Ensure-Section2Exists([string]$PageId) {
  if ($WhatIf) { return }

  for ($i=0; $i -lt 6; $i++) {
    $page = Get-PnPPage -Identity $PageId -ErrorAction Stop
    $count = 0
    if ($page.PSObject.Properties.Name -contains "Sections" -and $page.Sections) { $count = $page.Sections.Count }
    if ($count -ge 2) { return }

    Add-PnPPageSection -Page $page -SectionTemplate OneColumn -Order 2 | Out-Null
    $page = Get-PnPPage -Identity $PageId -ErrorAction Stop
    $page.Save(); $page.Publish()
    Start-Sleep -Seconds 2
  }

  throw "Could not ensure Section 2 exists on $PageId"
}

function Add-TextToSection2([string]$PageId, [string]$Html) {
  if ($WhatIf) {
    Write-Host "WHATIF: Would add text block to Section 2/Col 1 on $PageId" -ForegroundColor Yellow
    return
  }

  Ensure-Section2Exists -PageId $PageId

  $lastErr = $null
  for ($i=0; $i -lt 6; $i++) {
    try {
      Add-PnPPageTextPart -Page $PageId -Text $Html -Section 2 -Column 1 -ErrorAction Stop | Out-Null
      $p = Get-PnPPage -Identity $PageId
      $p.Save(); $p.Publish()
      return
    } catch {
      $lastErr = $_
      Start-Sleep -Seconds 2
    }
  }
  throw $lastErr
}

# ======================
# RUN
# ======================
Connect-Once
Write-Host "Connected to SharePoint"

$projects = Get-PnPListItem -List $ProjectsList -PageSize 2000 -Fields "ID","ProjectID","ProjectName",$ProjectHubField
$modules  = Get-PnPListItem -List $ModulesList  -PageSize 2000 -Fields "ID","ModuleID","ModuleName","ParentProject",$ModuleHubField

# Cache projects by list item id
$projectByListId = @{}
foreach ($p in $projects) { $projectByListId[[int]$p.Id] = $p }

# Group modules by parent project list item id
$modulesByProjectId = @{}
foreach ($m in $modules) {
  $parent = $m["ParentProject"]
  if ($null -eq $parent) { continue }
  $parentId = $parent.LookupId
  if (-not $modulesByProjectId.ContainsKey($parentId)) { $modulesByProjectId[$parentId] = @() }
  $modulesByProjectId[$parentId] += $m
}

# ======================
# 1) Project Hubs: Module Index (HYPERLINKED)
# ======================
foreach ($p in $projects) {
  $projectCode = [string]$p["ProjectID"]
  $projectHubUrl = Get-HubUrlFromFieldValue $p[$ProjectHubField]
  $pageId = Resolve-PageIdentityFromUrl $projectHubUrl

  if (-not $pageId) {
    Fail-Or-Warn "Project $projectCode hub page not found"
    continue
  }

  # Replace only the module index block (preserves everything else)
  Remove-TextComponentsByMarker -PageId $pageId -MarkerText $MarkerModules

  $items = @()
  $projectListItemId = [int]$p.Id
  if ($modulesByProjectId.ContainsKey($projectListItemId)) {
    foreach ($m in ($modulesByProjectId[$projectListItemId] | Sort-Object @{Expression={ [string]$_["ModuleID"] }; Ascending=$true })) {
      $mid  = [string]$m["ModuleID"]
      $name = [string]$m["ModuleName"]
      $mhub = Get-HubUrlFromFieldValue $m[$ModuleHubField]

      if (-not [string]::IsNullOrWhiteSpace($mhub)) {
        $items += "<li><a href=""$mhub"">$mid – $name</a></li>"
      } else {
        $items += "<li>$mid – $name</li>"
      }
    }
  }

  $html = @"
<p><em>$MarkerModules</em></p>
<h2>Modules</h2>
<ul>
$($items -join "`n")
</ul>
"@

  Add-TextToSection2 -PageId $pageId -Html $html
  Write-Host "Updated hyperlinked module index for $projectCode" -ForegroundColor Green
}

# ======================
# 2) Module Hubs: Parent Project backlink (existing behaviour preserved)
# ======================
foreach ($m in $modules) {
  $moduleCode = [string]$m["ModuleID"]
  $moduleHubUrl = Get-HubUrlFromFieldValue $m[$ModuleHubField]
  $pageId = Resolve-PageIdentityFromUrl $moduleHubUrl

  if (-not $pageId) {
    Fail-Or-Warn "Module $moduleCode hub page not found"
    continue
  }

  $parent = $m["ParentProject"]
  if ($null -eq $parent) { continue }

  $parentProjectListItemId = $parent.LookupId
  if (-not $projectByListId.ContainsKey($parentProjectListItemId)) { continue }

  $proj = $projectByListId[$parentProjectListItemId]
  $projCode = [string]$proj["ProjectID"]
  $projName = [string]$proj["ProjectName"]
  $projHubUrl = Get-HubUrlFromFieldValue $proj[$ProjectHubField]

  # Replace only the parent project block (preserves everything else)
  Remove-TextComponentsByMarker -PageId $pageId -MarkerText $MarkerParent

  $linkLine = if (-not [string]::IsNullOrWhiteSpace($projHubUrl)) {
    "<p><a href=""$projHubUrl"">$projCode – $projName</a></p>"
  } else {
    "<p>$projCode – $projName</p>"
  }

  $html = @"
<p><em>$MarkerParent</em></p>
<h2>Parent Project</h2>
$linkLine
"@

  Add-TextToSection2 -PageId $pageId -Html $html
  Write-Host "Updated parent project backlink for $moduleCode" -ForegroundColor Green
}

Write-Host "✅ Update-All-Hubs complete."