# ============================================
# Ensure Project & Module Hub Pages (PnP.PowerShell 3.x)
# - Creates hub pages under Site Pages/<Folder>
# - Writes Hub URLs back to ProjectHub / ModuleHub fields
# - Safe to re-run (idempotent)
# ============================================

Import-Module PnP.PowerShell -Force
Disconnect-PnPOnline -ErrorAction SilentlyContinue

# ======================
# AUTH (APP-ONLY CERT)
# ======================
$SiteUrl  = "https://zahe.sharepoint.com/sites/Projects"
$Tenant   = "zahe.onmicrosoft.com"
$ClientId = "01e1b71f-cbcb-48df-a076-871aa4ba10d9"   # <-- your working App Registration client id

$PfxPath = ".\ZaheZone-PnP-Projects.pfx"
$PfxPwd  = ConvertTo-SecureString "UseA-LongRandomPasswordHere" -AsPlainText -Force

Connect-PnPOnline `
  -Url $SiteUrl `
  -Tenant $Tenant `
  -ClientId $ClientId `
  -CertificatePath $PfxPath `
  -CertificatePassword $PfxPwd

# ======================
# SETTINGS
# ======================
$ProjectsList  = "Project Register"
$ModulesList   = "Project Modules"

$ProjectFolder = "Project Hubs"
$ModuleFolder  = "Module Hubs"

# ======================
# CMDLET GUARDS
# ======================
$requiredCmdlets = @(
  "Add-PnPPage",
  "Get-PnPPage",
  "Set-PnPPage",
  "Add-PnPPageSection",
  "Add-PnPPageTextPart",
  "Set-PnPListItem",
  "Get-PnPListItem",
  "Add-PnPFolder",
  "Get-PnPFolder"
)

foreach ($c in $requiredCmdlets) {
  if (-not (Get-Command $c -ErrorAction SilentlyContinue)) {
    throw "Required cmdlet not found: $c. Ensure PnP.PowerShell 3.x is installed and imported."
  }
}

# ======================
# HELPERS
# ======================
function UrlEncode-Folder([string]$folderName) {
  if ([string]::IsNullOrWhiteSpace($folderName)) { return $folderName }
  # minimal encoding for spaces (SharePoint folder urls)
  return ($folderName -replace ' ', '%20')
}

function Safe-BaseName([string]$value) {
  if ($null -eq $value) { return "Hub" }
  $safe = ($value -replace '[^A-Za-z0-9\-]', '').Trim()
  if ([string]::IsNullOrWhiteSpace($safe)) { $safe = "Hub" }
  return $safe
}

function Ensure-SitePagesFolders {
  $sitePages = Get-PnPList -Identity "Site Pages" -ErrorAction Stop
  Get-PnPProperty -ClientObject $sitePages -Property RootFolder | Out-Null
  Get-PnPProperty -ClientObject $sitePages.RootFolder -Property ServerRelativeUrl | Out-Null

  $root = $sitePages.RootFolder.ServerRelativeUrl

  foreach ($folder in @($ProjectFolder, $ModuleFolder)) {
    $folderUrl = "$root/$folder"
    $exists = Get-PnPFolder -Url $folderUrl -ErrorAction SilentlyContinue
    if (-not $exists) {
      Add-PnPFolder -Name $folder -Folder $root -ErrorAction Stop | Out-Null
    }
  }
}

function Has-HubUrl($fieldValue) {
  if ($null -eq $fieldValue) { return $false }

  # If it is a FieldUrlValue-like object
  if ($fieldValue.PSObject.Properties.Name -contains "Url") {
    return -not [string]::IsNullOrWhiteSpace([string]$fieldValue.Url)
  }

  # If it is a hashtable
  if ($fieldValue -is [hashtable]) {
    return ($fieldValue.ContainsKey("Url") -and -not [string]::IsNullOrWhiteSpace([string]$fieldValue["Url"]))
  }

  # If it is a plain string
  if ($fieldValue -is [string]) {
    return -not [string]::IsNullOrWhiteSpace($fieldValue)
  }

  return $false
}

function Set-LinkFieldRobust {
  param(
    [string]$ListName,
    [int]$ItemId,
    [string]$FieldInternalName,
    [string]$Url,
    [string]$Description
  )

  if ([string]::IsNullOrWhiteSpace($FieldInternalName)) {
    throw "FieldInternalName is null/empty for list [$ListName]."
  }
  if ([string]::IsNullOrWhiteSpace($Url)) {
    throw "URL is null/empty (cannot update [$ListName] item [$ItemId])."
  }

  # Method 1: Hashtable Url/Description (most common in PnP) 
  try {
    Set-PnPListItem -List $ListName -Identity $ItemId -Values @{
      $FieldInternalName = @{ Url = $Url; Description = $Description }
    } -ErrorAction Stop | Out-Null
    return
  } catch {}

  # Method 2: Plain string (some URL fields accept it)
  try {
    Set-PnPListItem -List $ListName -Identity $ItemId -Values @{
      $FieldInternalName = $Url
    } -ErrorAction Stop | Out-Null
    return
  } catch {}

  # Method 3: CSOM FieldUrlValue (last resort)
  $v = New-Object Microsoft.SharePoint.Client.FieldUrlValue
  $v.Url = $Url
  $v.Description = $Description
  Set-PnPListItem -List $ListName -Identity $ItemId -Values @{
    $FieldInternalName = $v
  } -ErrorAction Stop | Out-Null
}

function Ensure-HubPage {
  param(
    [string]$FolderName,
    [string]$BaseName,
    [string]$Title,
    [string]$BodyText
  )

  $folderNameSafe = [string]$FolderName
  $baseNameSafe   = Safe-BaseName $BaseName

  if ([string]::IsNullOrWhiteSpace($folderNameSafe)) { throw "FolderName is null/empty in Ensure-HubPage." }
  if ([string]::IsNullOrWhiteSpace($baseNameSafe))   { throw "BaseName is null/empty in Ensure-HubPage." }

  # Add-PnPPage supports "Folder/NewPage" style paths [6](https://pnp.github.io/powershell/cmdlets/Add-PnPPage.html)
  $pagePathPlain = "$folderNameSafe/$baseNameSafe"
  $pagePathEnc   = "$(UrlEncode-Folder $folderNameSafe)/$baseNameSafe"

  # Try to detect existing page (plain or URL-encoded folder)
  $existing = Get-PnPPage -Identity "$pagePathPlain.aspx" -ErrorAction SilentlyContinue
  if (-not $existing) {
    $existing = Get-PnPPage -Identity "$pagePathEnc.aspx" -ErrorAction SilentlyContinue
  }

  if (-not $existing) {
    # Create
    $page = Add-PnPPage -Name $pagePathPlain -LayoutType Article -ErrorAction Stop   # [6](https://pnp.github.io/powershell/cmdlets/Add-PnPPage.html)
    Set-PnPPage -Identity $page -Title $Title -CommentsEnabled:$false -ErrorAction Stop  # [8](https://www.sharepointdiary.com/2019/08/create-modern-page-in-sharepoint-online-using-powershell.html)

    Add-PnPPageSection -Page $page -SectionTemplate OneColumn -Order 1 -ErrorAction Stop   # [7](https://learn.microsoft.com/en-us/microsoft-365/community/working-with-modern-clientside-pages-using-pnp-powershell)[8](https://www.sharepointdiary.com/2019/08/create-modern-page-in-sharepoint-online-using-powershell.html)
    Add-PnPPageTextPart -Page $page -Text $BodyText -Section 1 -Column 1 -ErrorAction Stop # [7](https://learn.microsoft.com/en-us/microsoft-365/community/working-with-modern-clientside-pages-using-pnp-powershell)[8](https://www.sharepointdiary.com/2019/08/create-modern-page-in-sharepoint-online-using-powershell.html)

    $page.Publish()  # [8](https://www.sharepointdiary.com/2019/08/create-modern-page-in-sharepoint-online-using-powershell.html)
  }

  # Return absolute URL (folder will be URL-encoded in the URL)
  $web = Get-PnPWeb
  return "$($web.Url)/SitePages/$pagePathEnc.aspx"
}

# ======================
# TEMPLATES (plain text)
# ======================
$ProjectTemplate = @"
Project Purpose
- Why this project exists

Scope
In Scope
- …

Out of Scope
- …

Current Focus
- …

Key Decisions
- YYYY-MM-DD – Decision – Outcome

Modules
- Related module hub pages

Risks / Constraints
- Facts only

Authoritative Artefacts
- Docs / Runbooks
"@

$ModuleTemplate = @"
Module Purpose
- What this delivers

Scope
In Scope
- …

Out of Scope
- …

Current Focus
- …

Key Decisions
- YYYY-MM-DD – Decision – Outcome

Dependencies
- Other modules / vendors

Risks / Constraints
- Facts only

Authoritative Artefacts
- Docs / Runbooks
"@

# ======================
# RUN
# ======================
Ensure-SitePagesFolders

Write-Host "Using Project hub field: ProjectHub"
Write-Host "Using Module hub field : ModuleHub"

# ---- Projects
$projectItems = Get-PnPListItem -List $ProjectsList -PageSize 2000 -Fields "ID","ProjectID","ProjectName","ProjectHub"

foreach ($item in $projectItems) {

  $projectId   = [string]$item["ProjectID"]
  if ([string]::IsNullOrWhiteSpace($projectId)) { continue }

  $hubVal = $item["ProjectHub"]
  if (Has-HubUrl $hubVal) { continue }

  $projectName = [string]$item["ProjectName"]
  if ([string]::IsNullOrWhiteSpace($projectName)) { $projectName = "(no name)" }

  $base  = Safe-BaseName $projectId
  $title = "$projectId – $projectName"

  try {
    $url = Ensure-HubPage -FolderName $ProjectFolder -BaseName $base -Title $title -BodyText $ProjectTemplate
    Set-LinkFieldRobust -ListName $ProjectsList -ItemId $item.Id -FieldInternalName "ProjectHub" -Url $url -Description "Project Hub"
  }
  catch {
    Write-Warning "Project $projectId failed: $($_.Exception.Message)"
  }
}

# ---- Modules
$moduleItems = Get-PnPListItem -List $ModulesList -PageSize 2000 -Fields "ID","ModuleID","ModuleName","ModuleHub"

foreach ($item in $moduleItems) {

  $moduleId = [string]$item["ModuleID"]
  if ([string]::IsNullOrWhiteSpace($moduleId)) { continue }

  $hubVal = $item["ModuleHub"]
  if (Has-HubUrl $hubVal) { continue }

  $moduleName = [string]$item["ModuleName"]
  if ([string]::IsNullOrWhiteSpace($moduleName)) { $moduleName = "(no name)" }

  $base  = Safe-BaseName $moduleId
  $title = "$moduleId – $moduleName"

  try {
    $url = Ensure-HubPage -FolderName $ModuleFolder -BaseName $base -Title $title -BodyText $ModuleTemplate
    Set-LinkFieldRobust -ListName $ModulesList -ItemId $item.Id -FieldInternalName "ModuleHub" -Url $url -Description "Module Hub"
  }
  catch {
    Write-Warning "Module $moduleId failed: $($_.Exception.Message)"
  }
}

Write-Host "✅ Project & Module Hub Page automation complete."