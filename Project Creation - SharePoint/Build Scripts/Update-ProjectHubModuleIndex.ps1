# ============================================
# Update Project Hub Pages - Module Index (FullWidth-safe)
# PnP.PowerShell 3.x | App-only Cert Auth
#
# Strategy:
# - If page has a FullWidth top section, never write into Section 1
# - Ensure Section 2 exists as OneColumn (normal section)
# - Remove any existing Module Index text part (marker-based)
# - Add a new Module Index text part to Section 2, Column 1
# ============================================

Import-Module PnP.PowerShell -Force
#Disconnect-PnPOnline -ErrorAction SilentlyContinue

# ========== AUTH ==========
$SiteUrl  = "https://zahe.sharepoint.com/sites/Projects"
$Tenant   = "zahe.onmicrosoft.com"
$ClientId = "01e1b71f-cbcb-48df-a076-871aa4ba10d9"

$PfxPath = ".\ZaheZone-PnP-Projects.pfx"
$PfxPwd  = ConvertTo-SecureString "UseA-LongRandomPasswordHere" -AsPlainText -Force

Connect-PnPOnline `
  -Url $SiteUrl `
  -Tenant $Tenant `
  -ClientId $ClientId `
  -CertificatePath $PfxPath `
  -CertificatePassword $PfxPwd

# ========== LISTS ==========
$ProjectsList = "Project Register"
$ModulesList  = "Project Modules"
$ProjectHubField = "ProjectHub"
$ModuleHubField  = "ModuleHub"

# Marker used to find/remove the module index text part
$Marker = "<!--MODULE_INDEX-->"

# ========== HELPERS ==========
function Get-PageIdentityCandidatesFromHubUrl {
    param([string]$HubUrl)

    if ([string]::IsNullOrWhiteSpace($HubUrl)) { return @() }

    $split = $HubUrl -split "/SitePages/", 2
    if ($split.Count -lt 2) { return @() }

    $raw = $split[1]  # e.g. Project%20Hubs/PRJ-001.aspx
    $decoded = [System.Uri]::UnescapeDataString($raw)  # e.g. Project Hubs/PRJ-001.aspx
    $fileName = [System.IO.Path]::GetFileName($decoded) # PRJ-001.aspx

    return @($decoded, $raw, $fileName) | Select-Object -Unique
}

function Get-ExistingPageIdentity {
    param([string[]]$Candidates)

    foreach ($c in $Candidates) {
        if ([string]::IsNullOrWhiteSpace($c)) { continue }
        $p = Get-PnPPage -Identity $c -ErrorAction SilentlyContinue
        if ($p) { return $c }
    }
    return $null
}

function Build-ModuleIndexHtml {
    param([array]$ProjectModules)

    if ($null -eq $ProjectModules -or $ProjectModules.Count -eq 0) {
        return @"
$Marker
<h2>Modules</h2>
<p><i>No modules registered.</i></p>
"@
    }

    $html = @"
$Marker
<h2>Modules</h2>
<ul>
"@

    foreach ($m in $ProjectModules) {
        $mid  = [string]$m["ModuleID"]
        $name = [string]$m["ModuleName"]
        $hub  = $m[$ModuleHubField]

        $url = $null
        if ($hub -and $hub.PSObject.Properties.Name -contains "Url") { $url = $hub.Url }
        elseif ($hub -is [string]) { $url = $hub }

        if (-not [string]::IsNullOrWhiteSpace($url)) {
            $html += "<li><a href=""$url"">$mid – $name</a></li>`n"
        } else {
            $html += "<li>$mid – $name</li>`n"
        }
    }

    $html += "</ul>`n"
    return $html
}

function Ensure-NormalSection2 {
    param([string]$PageIdentity)

    # Load the page object
    $page = Get-PnPPage -Identity $PageIdentity -ErrorAction Stop

    # If Sections is available, only add if missing
    $sectionCount = 0
    if ($page.PSObject.Properties.Name -contains "Sections" -and $page.Sections) {
        $sectionCount = $page.Sections.Count
    }

    if ($sectionCount -lt 2) {
        # Add a normal OneColumn section as section 2 (safe for text) odern-clientside-pages-using-pnp-powershell)
        Add-PnPPageSection -Page $page -SectionTemplate OneColumn -Order 2 | Out-Null
        $page.Save()
        $page.Publish()
    }
}

function Remove-ExistingModuleIndexParts {
    param([string]$PageIdentity)

    $components = Get-PnPPageComponent -Page $PageIdentity  # Retrieve components [3](https://pnp.github.io/powershell/cmdlets/Get-PnPPageComponent.html)
    foreach ($c in $components) {
        $hasTextProp = ($c.PSObject.Properties.Name -contains "Text")
        $hasJsonProp = ($c.PSObject.Properties.Name -contains "PropertiesJson")

        $match = $false
        if ($hasTextProp -and ($c.Text -like "*MODULE_INDEX*")) { $match = $true }
        if ($hasJsonProp -and ($c.PropertiesJson -like "*MODULE_INDEX*")) { $match = $true }

        if ($match) {
            Remove-PnPPageComponent -Page $PageIdentity -InstanceId $c.InstanceId -Force  # Remove component [4](https://pnp.github.io/powershell/cmdlets/Remove-PnPPageComponent.html)
        }
    }
}

function Add-ModuleIndexToSection2 {
    param(
        [string]$PageIdentity,
        [string]$Html
    )

    # Add text part specifically to Section 2, Column 1 (avoids FullWidth section 1)
    # Positioned parameter set requires Section & Column [3](https://pnp.github.io/powershell/cmdlets/Add-PnPPageTextPart.html)
    Add-PnPPageTextPart -Page $PageIdentity -Text $Html -Section 2 -Column 1 | Out-Null  # [3](https://pnp.github.io/powershell/cmdlets/Add-PnPPageTextPart.html)

    $page = Get-PnPPage -Identity $PageIdentity
    $page.Save()
    $page.Publish()
}

# ========== LOAD DATA ==========
$projects = Get-PnPListItem -List $ProjectsList -PageSize 2000 -Fields "ID","ProjectID",$ProjectHubField
$modules  = Get-PnPListItem -List $ModulesList  -PageSize 2000 -Fields "ID","ModuleID","ModuleName","ModuleSequence","ParentProject",$ModuleHubField

# Group modules by parent project list item id
$modulesByParentId = @{}
foreach ($m in $modules) {
    $parent = $m["ParentProject"]
    if ($null -eq $parent) { continue }

    $parentItemId = $parent.LookupId
    if (-not $modulesByParentId.ContainsKey($parentItemId)) {
        $modulesByParentId[$parentItemId] = @()
    }
    $modulesByParentId[$parentItemId] += $m
}

# ========== UPDATE PROJECT HUB PAGES ==========
foreach ($p in $projects) {

    $projectId = [string]$p["ProjectID"]
    if ([string]::IsNullOrWhiteSpace($projectId)) { continue }

    $hubObj = $p[$ProjectHubField]
    $hubUrl = $null
    if ($hubObj -and $hubObj.PSObject.Properties.Name -contains "Url") { $hubUrl = $hubObj.Url }
    elseif ($hubObj -is [string]) { $hubUrl = $hubObj }

    if ([string]::IsNullOrWhiteSpace($hubUrl)) {
        Write-Warning "Project $projectId has no ProjectHub URL - skipping"
        continue
    }

    $candidates = Get-PageIdentityCandidatesFromHubUrl -HubUrl $hubUrl
    $pageIdentity = Get-ExistingPageIdentity -Candidates $candidates
    if ([string]::IsNullOrWhiteSpace($pageIdentity)) {
        Write-Warning "❌ Failed updating $projectId : page not found. Candidates tried: $($candidates -join ', ')"
        continue
    }

    $projectListItemId = [int]$p.Id
    $projectModules = @()
    if ($modulesByParentId.ContainsKey($projectListItemId)) {
        $projectModules = $modulesByParentId[$projectListItemId] | Sort-Object `
            @{Expression = { $_["ModuleSequence"] -as [double] }; Ascending = $true}, `
            @{Expression = { [string]$_["ModuleID"] }; Ascending = $true}
    }

    $html = Build-ModuleIndexHtml -ProjectModules $projectModules

    try {
        Ensure-NormalSection2 -PageIdentity $pageIdentity
        Remove-ExistingModuleIndexParts -PageIdentity $pageIdentity
        Add-ModuleIndexToSection2 -PageIdentity $pageIdentity -Html $html

        Write-Host "✅ Updated module index for $projectId (page: $pageIdentity)"
    }
    catch {
        Write-Warning "❌ Failed updating $projectId : $($_.Exception.Message)"
    }
}

Write-Host "✅ Project Hub module indexes updated."