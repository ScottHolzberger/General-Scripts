<#
Authoritative runtime document.
If it’s not described here, it’s not supported.

Provision-Hubs.ps1
- Creates/ensures Project + Module hub pages
- Inserts scaffolds once (marker-based)
- Creates/ensures Project + Module runbook pages
- Links runbook from hub pages (clickable hyperlink)
- Writes ProjectHub / ModuleHub / Runbook list columns as plain text URLs
- Full-width safe: falls back to Section 2 for any text insertion when needed

Fixes included:
- No ".Count" calls on pipeline outputs (uses @(...) + .Length)
- Safe removal: only calls Remove-PnPPageComponent when a valid InstanceId GUID exists
#>

[CmdletBinding()]
param(
  [switch]$WhatIf,
  [switch]$Enforce,
  [switch]$SkipUpdateAllHubs,

  [string]$TemplatesPath = "$PSScriptRoot\Templates",

  # AUTH
  [string]$SiteUrl  = "https://zahe.sharepoint.com/sites/Projects",
  [string]$Tenant   = "zahe.onmicrosoft.com",
  [string]$ClientId = "01e1b71f-cbcb-48df-a076-871aa4ba10d9",
  [string]$PfxPath  = "$PSScriptRoot\ZaheZone-PnP-Projects.pfx",
  [string]$PfxPasswordPlain = "UseA-LongRandomPasswordHere",

  # LISTS / FIELDS
  [string]$ProjectsList        = "Project Register",
  [string]$ModulesList         = "Project Modules",
  [string]$ProjectHubField     = "ProjectHub",
  [string]$ModuleHubField      = "ModuleHub",
  [string]$ProjectRunbookField = "Runbook",
  [string]$ModuleRunbookField  = "Runbook",

  # PAGE FOLDERS (under Site Pages)
  [string]$ProjectFolderName = "Project Hubs",
  [string]$ModuleFolderName  = "Module Hubs",

  # RUNBOOK FOLDERS (under Site Pages)
  [string]$RunbooksFolder        = "Runbooks",
  [string]$ModuleRunbooksFolder  = "Runbooks/Modules"
)

Set-StrictMode -Version Latest
Import-Module PnP.PowerShell -Force

# ===== Durable markers (visible text) =====
$ProjectScaffoldMarker = "ZZ-AUTO-SCAFFOLD-PROJECT"
$ModuleScaffoldMarker  = "ZZ-AUTO-SCAFFOLD-MODULE"
$ProjectRunbookMarker  = "ZZ-AUTO-PROJECT-RUNBOOK"
$ModuleRunbookMarker   = "ZZ-AUTO-MODULE-RUNBOOK"
$ProjectRunbookLinkMarker = "ZZ-AUTO-RUNBOOK-LINK-PROJECT"
$ModuleRunbookLinkMarker  = "ZZ-AUTO-RUNBOOK-LINK-MODULE"

function Fail-Or-Warn([string]$msg) {
  if ($Enforce) { throw $msg }
  Write-Warning $msg
}

function Resolve-Template([string]$name) {
  $p1 = Join-Path $TemplatesPath $name
  if (Test-Path $p1) { return $p1 }
  $p2 = Join-Path $PSScriptRoot $name
  if (Test-Path $p2) { return $p2 }
  throw "Missing required template: $name (looked in '$TemplatesPath' and '$PSScriptRoot')"
}

function UrlEncodePath([string]$path) {
  return ($path -replace ' ', '%20')
}

function Connect-Once {
  $pwd = ConvertTo-SecureString $PfxPasswordPlain -AsPlainText -Force
  #Disconnect-PnPOnline -ErrorAction SilentlyContinue
  Connect-PnPOnline -Url $SiteUrl -Tenant $Tenant -ClientId $ClientId -CertificatePath $PfxPath -CertificatePassword $pwd
}

function Ensure-SitePagesFolderPath([string]$folderPath) {
  $root = "SitePages"
  $parts = $folderPath -split "/"
  $current = $root

  foreach ($p in $parts) {
    $next = "$current/$p"
    $exists = Get-PnPFolder -Url $next -ErrorAction SilentlyContinue
    if (-not $exists) {
      if ($WhatIf) {
        Write-Host "WHATIF: Would create folder '$p' under '$current'" -ForegroundColor Yellow
      } else {
        Add-PnPFolder -Name $p -Folder $current | Out-Null
      }
    }
    $current = $next
  }
}

function Ensure-PageExists([string]$pageRelNoExt, [string]$title) {
  $pageRel = "$pageRelNoExt.aspx"
  $existing = Get-PnPPage -Identity $pageRel -ErrorAction SilentlyContinue
  if ($existing) { return $pageRel }

  if ($WhatIf) {
    Write-Host "WHATIF: Would create page $pageRel" -ForegroundColor Yellow
    return $pageRel
  }

  Add-PnPPage -Name $pageRelNoExt -Title $title -LayoutType Article -CommentsEnabled:$false -Publish | Out-Null
  Write-Host "Created page: $pageRel" -ForegroundColor Green

  for ($i=0; $i -lt 8; $i++) {
    if (Get-PnPPage -Identity $pageRel -ErrorAction SilentlyContinue) { break }
    Start-Sleep -Seconds 2
  }

  return $pageRel
}

function Component-Text([object]$component) {
  if ($component.PSObject.Properties.Name -contains "Text") { return [string]$component.Text }
  return ""
}

function Get-ComponentInstanceId([object]$component) {
  # Prefer InstanceId if present and GUID-parseable
  foreach ($propName in @("InstanceId","Id","ComponentId")) {
    if ($component.PSObject.Properties.Name -contains $propName) {
      $v = $component.$propName
      if ($v -is [Guid]) { return $v }
      if ($v) {
        $s = [string]$v
        $out = [Guid]::Empty
        if ([Guid]::TryParse($s, [ref]$out)) { return $out }
      }
    }
  }
  return $null
}

function Get-TextComponentsByMarker([string]$pageIdentity, [string]$markerText) {
  $hits = @()
  $p = Get-PnPPage -Identity $pageIdentity -ErrorAction SilentlyContinue
  if (-not $p) { return @() }

  foreach ($c in Get-PnPPageComponent -Page $pageIdentity) {
    $t = Component-Text $c
    if (-not [string]::IsNullOrWhiteSpace($t) -and $t -like "*$markerText*") {
      $hits += $c
    }
  }
  return @($hits)   # force array
}

function Page-HasMarker([string]$pageIdentity, [string]$markerText) {
  $hits = @(Get-TextComponentsByMarker -pageIdentity $pageIdentity -markerText $markerText)
  return ($hits.Length -gt 0)
}

function Remove-TextPartsByMarker([string]$pageIdentity, [string]$markerText) {
  $hits = @(Get-TextComponentsByMarker -pageIdentity $pageIdentity -markerText $markerText)
  foreach ($c in $hits) {
    $iid = Get-ComponentInstanceId $c
    if (-not $iid) {
      # Component lacks removable instance id; skip safely
      continue
    }

    if ($WhatIf) {
      Write-Host "WHATIF: Would remove component $iid from $pageIdentity (marker=$markerText)" -ForegroundColor Yellow
    } else {
      Remove-PnPPageComponent -Page $pageIdentity -InstanceId $iid -Force | Out-Null
    }
  }
}

function Ensure-Section2Exists([string]$pageIdentity) {
  if ($WhatIf) { return }

  for ($i=0; $i -lt 6; $i++) {
    $page = Get-PnPPage -Identity $pageIdentity -ErrorAction Stop
    $count = 0
    if ($page.PSObject.Properties.Name -contains "Sections" -and $page.Sections) { $count = $page.Sections.Count }
    if ($count -ge 2) { return }

    Add-PnPPageSection -Page $page -SectionTemplate OneColumn -Order 2 | Out-Null
    $page = Get-PnPPage -Identity $pageIdentity -ErrorAction Stop
    $page.Save(); $page.Publish()
    Start-Sleep -Seconds 2
  }

  throw "Could not ensure Section 2 exists on $pageIdentity"
}

function Add-TextPartSafe {
  param(
    [string]$PageIdentity,
    [string]$Html,
    [switch]$PreferSection2
  )

  if ($WhatIf) {
    Write-Host "WHATIF: Would add text part to $PageIdentity" -ForegroundColor Yellow
    return
  }

  if ($PreferSection2) {
    Ensure-Section2Exists -pageIdentity $PageIdentity
    Add-PnPPageTextPart -Page $PageIdentity -Text $Html -Section 2 -Column 1 -ErrorAction Stop | Out-Null
    $p = Get-PnPPage -Identity $PageIdentity
    $p.Save(); $p.Publish()
    return
  }

  try {
    Add-PnPPageTextPart -Page $PageIdentity -Text $Html -ErrorAction Stop | Out-Null
    $p = Get-PnPPage -Identity $PageIdentity
    $p.Save(); $p.Publish()
    return
  } catch {
    $msg = $_.Exception.Message
    if ($msg -match "one column full width") {
      Ensure-Section2Exists -pageIdentity $PageIdentity
      Add-PnPPageTextPart -Page $PageIdentity -Text $Html -Section 2 -Column 1 -ErrorAction Stop | Out-Null
      $p = Get-PnPPage -Identity $PageIdentity
      $p.Save(); $p.Publish()
      return
    }
    throw
  }
}

function Add-ScaffoldIfMissing([string]$pageIdentity, [string]$templatePath, [string]$markerText) {
  if (Page-HasMarker $pageIdentity $markerText) { return }
  $html = Get-Content $templatePath -Raw
  Add-TextPartSafe -PageIdentity $pageIdentity -Html $html
}

function Set-ListTextValue([string]$listName, [int]$itemId, [string]$fieldInternalName, [string]$value) {
  if ([string]::IsNullOrWhiteSpace($value)) {
    Fail-Or-Warn "List '$listName' item $itemId value empty for '$fieldInternalName'"
    return
  }

  if ($WhatIf) {
    Write-Host "WHATIF: Would set $listName item $itemId field '$fieldInternalName' -> $value" -ForegroundColor Yellow
    return
  }

  Set-PnPListItem -List $listName -Identity $itemId -Values @{ $fieldInternalName = $value } -ErrorAction Stop | Out-Null
}

function Ensure-RunbookPage {
  param(
    [string]$FolderRel,
    [string]$PageBaseName,
    [string]$Title,
    [string]$MarkerText,
    [string]$BodyHtml
  )

  Ensure-SitePagesFolderPath $FolderRel

  $pageRelNoExt = "$FolderRel/$PageBaseName"
  $pageIdentity = "$pageRelNoExt.aspx"

  $existing = Get-PnPPage -Identity $pageIdentity -ErrorAction SilentlyContinue
  if (-not $existing) {
    if ($WhatIf) {
      Write-Host "WHATIF: Would create runbook page $pageIdentity" -ForegroundColor Yellow
    } else {
      Add-PnPPage -Name $pageRelNoExt -Title $Title -LayoutType Article -CommentsEnabled:$false -Publish | Out-Null
      Write-Host "Created runbook page: $pageIdentity" -ForegroundColor Green
    }
  }

  if (-not (Page-HasMarker $pageIdentity $MarkerText)) {
    $html = @"
<p><em>$MarkerText</em></p>
$BodyHtml
"@
    Add-TextPartSafe -PageIdentity $pageIdentity -Html $html
  }

  return $pageIdentity
}

function Ensure-RunbookLinkOnHubPage {
  param(
    [string]$HubPageIdentity,
    [string]$MarkerText,
    [string]$LinkUrl,
    [string]$LinkLabel
  )

  Remove-TextPartsByMarker -pageIdentity $HubPageIdentity -markerText $MarkerText

  $html = @"
<p><em>$MarkerText</em></p>
<h3>Runbook</h3>
<p><a href=""$LinkUrl"">$LinkLabel</a></p>
"@

  Add-TextPartSafe -PageIdentity $HubPageIdentity -Html $html -PreferSection2
}

# ======================
# MAIN
# ======================
$ProjectTemplate = Resolve-Template "ProjectHub.html"
$ModuleTemplate  = Resolve-Template "ModuleHub.html"

Connect-Once
$web = Get-PnPWeb
$baseUrl = $web.Url

Ensure-SitePagesFolderPath $ProjectFolderName
Ensure-SitePagesFolderPath $ModuleFolderName
Ensure-SitePagesFolderPath $RunbooksFolder
Ensure-SitePagesFolderPath $ModuleRunbooksFolder

$projectFolderEnc = UrlEncodePath $ProjectFolderName
$moduleFolderEnc  = UrlEncodePath $ModuleFolderName

$projects = Get-PnPListItem -List $ProjectsList -Fields "ID","ProjectID","ProjectName",$ProjectHubField,$ProjectRunbookField
$modules  = Get-PnPListItem -List $ModulesList  -Fields "ID","ModuleID","ModuleName","ParentProject",$ModuleHubField,$ModuleRunbookField

foreach ($p in $projects) {
  $projectCode = [string]$p["ProjectID"]
  if ([string]::IsNullOrWhiteSpace($projectCode)) { Fail-Or-Warn "Project item $($p.Id) missing ProjectID; skipping"; continue }

  $projectName = [string]$p["ProjectName"]
  if ([string]::IsNullOrWhiteSpace($projectName)) { $projectName = "(no name)" }

  $hubRelNoExt = "$ProjectFolderName/$projectCode"
  $hubPage = Ensure-PageExists -pageRelNoExt $hubRelNoExt -title "$projectCode – $projectName"

  Add-ScaffoldIfMissing -pageIdentity $hubPage -templatePath $ProjectTemplate -markerText $ProjectScaffoldMarker

  $projectHubUrl = "$baseUrl/SitePages/$projectFolderEnc/$projectCode.aspx"
  Set-ListTextValue -listName $ProjectsList -itemId $p.Id -fieldInternalName $ProjectHubField -value $projectHubUrl

  $runbookBaseName = "$projectCode-Runbook"
  $null = Ensure-RunbookPage -FolderRel $RunbooksFolder -PageBaseName $runbookBaseName -Title "$projectCode – Runbook" -MarkerText $ProjectRunbookMarker -BodyHtml @"
<h1>$projectCode – Project Runbook</h1>
<h2>Purpose</h2><p>Authoritative runbook for operating this project.</p>
<h2>Supported Operations</h2><ul><li>Provision-Hubs.ps1</li><li>Update-All-Hubs.ps1</li></ul>
<h2>Change Control</h2><p>All changes must align with PRJ-002 governance rules.</p>
"@

  $projectRunbookUrl = "$baseUrl/SitePages/$(UrlEncodePath $RunbooksFolder)/$runbookBaseName.aspx"
  Set-ListTextValue -listName $ProjectsList -itemId $p.Id -fieldInternalName $ProjectRunbookField -value $projectRunbookUrl

  Ensure-RunbookLinkOnHubPage -HubPageIdentity $hubPage -MarkerText $ProjectRunbookLinkMarker -LinkUrl $projectRunbookUrl -LinkLabel "$projectCode – Runbook"

  Write-Host "Linked ProjectHub + Runbook for $projectCode" -ForegroundColor Green
}

foreach ($m in $modules) {
  $moduleCode = [string]$m["ModuleID"]
  if ([string]::IsNullOrWhiteSpace($moduleCode)) { Fail-Or-Warn "Module item $($m.Id) missing ModuleID; skipping"; continue }
  if (-not $m["ParentProject"]) { Fail-Or-Warn "Module $moduleCode missing ParentProject; skipping"; continue }

  $moduleName = [string]$m["ModuleName"]
  if ([string]::IsNullOrWhiteSpace($moduleName)) { $moduleName = "(no name)" }

  $hubRelNoExt = "$ModuleFolderName/$moduleCode"
  $hubPage = Ensure-PageExists -pageRelNoExt $hubRelNoExt -title "$moduleCode – $moduleName"

  Add-ScaffoldIfMissing -pageIdentity $hubPage -templatePath $ModuleTemplate -markerText $ModuleScaffoldMarker

  $moduleHubUrl = "$baseUrl/SitePages/$moduleFolderEnc/$moduleCode.aspx"
  Set-ListTextValue -listName $ModulesList -itemId $m.Id -fieldInternalName $ModuleHubField -value $moduleHubUrl

  $runbookBaseName = "$moduleCode-Runbook"
  $null = Ensure-RunbookPage -FolderRel $ModuleRunbooksFolder -PageBaseName $runbookBaseName -Title "$moduleCode – Runbook" -MarkerText $ModuleRunbookMarker -BodyHtml @"
<h1>$moduleCode – Module Runbook</h1>
<h2>Purpose</h2><p>Authoritative runbook for operating this module.</p>
<h2>Supported Operations</h2><ul><li>Provision-Hubs.ps1</li><li>Update-All-Hubs.ps1</li></ul>
<h2>Change Control</h2><p>All changes must align with PRJ-002 governance rules.</p>
"@

  $moduleRunbookUrl = "$baseUrl/SitePages/$(UrlEncodePath $ModuleRunbooksFolder)/$runbookBaseName.aspx"
  Set-ListTextValue -listName $ModulesList -itemId $m.Id -fieldInternalName $ModuleRunbookField -value $moduleRunbookUrl

  Ensure-RunbookLinkOnHubPage -HubPageIdentity $hubPage -MarkerText $ModuleRunbookLinkMarker -LinkUrl $moduleRunbookUrl -LinkLabel "$moduleCode – Runbook"

  Write-Host "Linked ModuleHub + Runbook for $moduleCode" -ForegroundColor Green
}

if (-not $SkipUpdateAllHubs) {
  $update = Join-Path $PSScriptRoot "Update-All-Hubs.ps1"
  if (Test-Path $update) {
    if ($WhatIf) {
      Write-Host "WHATIF: Would run Update-All-Hubs.ps1" -ForegroundColor Yellow
    } else {
      Write-Host "Running Update-All-Hubs.ps1..." -ForegroundColor Cyan
      & $update
    }
  } else {
    Fail-Or-Warn "Update-All-Hubs.ps1 not found at $update"
  }
}

Write-Host "✅ Provision-Hubs complete." -ForegroundColor Green