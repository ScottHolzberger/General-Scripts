# ============================================
# Update Module Hub Pages - Parent Project Backlink (FullWidth-safe)
# PnP.PowerShell 3.x | App-only Cert Auth
#
# - Reads Project Modules list
# - Uses ParentProject lookup to find parent project + ProjectHub URL
# - Updates/creates a dedicated Text part on each Module Hub page
# - Writes into Section 2, Column 1 to avoid FullWidth section restrictions
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

# Marker used to find/remove the correct Text part
$Marker = "<!--PARENT_PROJECT-->"

# ========== HELPERS ==========
function Get-PageIdentityCandidatesFromHubUrl {
    param([string]$HubUrl)

    if ([string]::IsNullOrWhiteSpace($HubUrl)) { return @() }

    $split = $HubUrl -split "/SitePages/", 2
    if ($split.Count -lt 2) { return @() }

    $raw = $split[1]  # e.g. Module%20Hubs/PRJ-001-001.aspx
    $decoded = [System.Uri]::UnescapeDataString($raw)  # Module Hubs/PRJ-001-001.aspx
    $fileName = [System.IO.Path]::GetFileName($decoded) # PRJ-001-001.aspx

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

function Ensure-NormalSection2 {
    param([string]$PageIdentity)

    $page = Get-PnPPage -Identity $PageIdentity -ErrorAction Stop

    $sectionCount = 0
    if ($page.PSObject.Properties.Name -contains "Sections" -and $page.Sections) {
        $sectionCount = $page.Sections.Count
    }

    if ($sectionCount -lt 2) {
        # Add a normal OneColumn section as section 2 (safe for text)
        Add-PnPPageSection -Page $page -SectionTemplate OneColumn -Order 2 | Out-Null
        $page.Save()
        $page.Publish()
    }
}

function Remove-ExistingParentProjectParts {
    param([string]$PageIdentity)

    $components = Get-PnPPageComponent -Page $PageIdentity  # [1](https://pnp.github.io/powershell/cmdlets/Get-PnPPageComponent.html)

    foreach ($c in $components) {
        $match = $false

        if (($c.PSObject.Properties.Name -contains "Text") -and ($c.Text -like "*PARENT_PROJECT*")) { $match = $true }
        if (($c.PSObject.Properties.Name -contains "PropertiesJson") -and ($c.PropertiesJson -like "*PARENT_PROJECT*")) { $match = $true }

        if ($match) {
            Remove-PnPPageComponent -Page $PageIdentity -InstanceId $c.InstanceId -Force  # [2](https://pnp.github.io/powershell/cmdlets/Remove-PnPPageComponent.html)
        }
    }
}

function Add-ParentProjectBacklinkToSection2 {
    param(
        [string]$PageIdentity,
        [string]$Html
    )

    # Place the backlink in Section 2 / Column 1 to avoid FullWidth restrictions
    Add-PnPPageTextPart -Page $PageIdentity -Text $Html -Section 2 -Column 1 | Out-Null  # [1](https://pnp.github.io/powershell/cmdlets/Add-PnPPageTextPart.html)

    $page = Get-PnPPage -Identity $PageIdentity
    $page.Save()
    $page.Publish()
}

function Build-BacklinkHtml {
    param(
        [string]$ProjectId,
        [string]$ProjectName,
        [string]$ProjectHubUrl
    )

    $safeName = $ProjectName
    if ([string]::IsNullOrWhiteSpace($safeName)) { $safeName = "(no name)" }

    if ([string]::IsNullOrWhiteSpace($ProjectHubUrl)) {
        return @"
$Marker
<h2>Parent Project</h2>
<p><b>$ProjectId</b> – $safeName</p>
"@
    }

    return @"
$Marker
<h2>Parent Project</h2>
<p><a href="$ProjectHubUrl">$ProjectId – $safeName</a></p>
"@
}

# ========== LOAD DATA ==========
$projects = Get-PnPListItem -List $ProjectsList -PageSize 2000 -Fields "ID","ProjectID","ProjectName",$ProjectHubField
$projectById = @{}
foreach ($p in $projects) { $projectById[[int]$p.Id] = $p }

$modules = Get-PnPListItem -List $ModulesList -PageSize 2000 -Fields "ID","ModuleID","ModuleName","ParentProject",$ModuleHubField

# ========== UPDATE MODULE PAGES ==========
foreach ($m in $modules) {

    $moduleId = [string]$m["ModuleID"]
    if ([string]::IsNullOrWhiteSpace($moduleId)) { continue }

    # Need a module hub url to know which page to edit
    $moduleHubObj = $m[$ModuleHubField]
    $moduleHubUrl = $null
    if ($moduleHubObj -and $moduleHubObj.PSObject.Properties.Name -contains "Url") { $moduleHubUrl = $moduleHubObj.Url }
    elseif ($moduleHubObj -is [string]) { $moduleHubUrl = $moduleHubObj }

    if ([string]::IsNullOrWhiteSpace($moduleHubUrl)) {
        Write-Warning "Module $moduleId has no ModuleHub URL - skipping"
        continue
    }

    # Resolve module page identity
    $candidates = Get-PageIdentityCandidatesFromHubUrl -HubUrl $moduleHubUrl
    $pageIdentity = Get-ExistingPageIdentity -Candidates $candidates
    if ([string]::IsNullOrWhiteSpace($pageIdentity)) {
        Write-Warning "❌ Module $moduleId page not found. Candidates tried: $($candidates -join ', ')"
        continue
    }

    # Get parent project
    $parent = $m["ParentProject"]
    if ($null -eq $parent) {
        Write-Warning "Module $moduleId has no ParentProject lookup - skipping"
        continue
    }

    $parentItemId = $parent.LookupId
    if (-not $projectById.ContainsKey($parentItemId)) {
        Write-Warning "Module $moduleId parent project item id $parentItemId not found in cache - skipping"
        continue
    }

    $proj = $projectById[$parentItemId]
    $projectId = [string]$proj["ProjectID"]
    $projectName = [string]$proj["ProjectName"]

    $projectHubObj = $proj[$ProjectHubField]
    $projectHubUrl = $null
    if ($projectHubObj -and $projectHubObj.PSObject.Properties.Name -contains "Url") { $projectHubUrl = $projectHubObj.Url }
    elseif ($projectHubObj -is [string]) { $projectHubUrl = $projectHubObj }

    $html = Build-BacklinkHtml -ProjectId $projectId -ProjectName $projectName -ProjectHubUrl $projectHubUrl

    try {
        Ensure-NormalSection2 -PageIdentity $pageIdentity
        Remove-ExistingParentProjectParts -PageIdentity $pageIdentity
        Add-ParentProjectBacklinkToSection2 -PageIdentity $pageIdentity -Html $html
        Write-Host "✅ Updated Parent Project backlink on $moduleId (page: $pageIdentity)"
    }
    catch {
        Write-Warning "❌ Failed updating $moduleId : $($_.Exception.Message)"
    }
}

Write-Host "✅ Module Parent Project backlinks updated."