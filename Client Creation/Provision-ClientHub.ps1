param(
    [Parameter(Mandatory)]
    [string]$ClientLookupValue
)

Write-Host "==== Starting Client Provisioning ====" -ForegroundColor Cyan

# ==========================================================
# AUTH
# ==========================================================
$Tenant = "zahe.onmicrosoft.com"
$ClientId = "01e1b71f-cbcb-48df-a076-871aa4ba10d9"
$PfxPath = ".\ZaheZone-PnP-Projects.pfx"
$PfxPasswordPlain = "UseA-LongRandomPasswordHere"
$PfxPassword = ConvertTo-SecureString $PfxPasswordPlain -AsPlainText -Force

Set-Variable -Name ClientId -Value $ClientId -Option ReadOnly -Force

# ==========================================================
# CONFIG
# ==========================================================
$TenantRootUrl = "https://zahe.sharepoint.com"
$TenantAdminUrl = "https://zahe-admin.sharepoint.com"
$ControlSiteUrl = "https://zahe.sharepoint.com/sites/ZZ-Control"
$HubSiteUrl = "https://zahe.sharepoint.com/sites/ZZ-Clients-Hub"

$ClientRegistryList = "Client Registry"
$SiteOwner = "scott@zahezone.com.au"

# Template pages (stored in ZZ-Control Site Pages root)
$NetworkTemplateServerRelativeUrl = "/sites/ZZ-Control/SitePages/Network-Template.aspx"
$SoftwareTemplateServerRelativeUrl = "/sites/ZZ-Control/SitePages/Software-Template.aspx"

$RetryCount = 30
$RetryDelaySeconds = 5

# ==========================================================
# FUNCTIONS
# ==========================================================

function Configure-CleanNavigation {
    param(
        [string]$SiteUrl
    )

    Write-Host "Rebuilding navigation..." -ForegroundColor Cyan

    # ---------------------------
    # REMOVE ALL EXISTING NAV
    # ---------------------------
    $navNodes = @(Get-PnPNavigationNode -Location QuickLaunch)

    foreach ($node in $navNodes) {
        try {
            Remove-PnPNavigationNode -Identity $node.Id -Force -ErrorAction Stop
        }
        catch {
            Write-Host "Could not remove: $($node.Title)"
        }
    }

    Write-Host "Cleared existing navigation"

    # ---------------------------
    # REBUILD NAV (CLEAN)
    # ---------------------------

    # Main pages
    Add-PnPNavigationNode -Title "Overview"   -Url "$SiteUrl/SitePages/Overview.aspx"   -Location QuickLaunch
    Add-PnPNavigationNode -Title "Network"    -Url "$SiteUrl/SitePages/Network.aspx"    -Location QuickLaunch
    Add-PnPNavigationNode -Title "Software"   -Url "$SiteUrl/SitePages/Software.aspx"   -Location QuickLaunch
    Add-PnPNavigationNode -Title "Projects"   -Url "$SiteUrl/SitePages/Projects.aspx"   -Location QuickLaunch
    Add-PnPNavigationNode -Title "Operations" -Url "$SiteUrl/SitePages/Operations.aspx" -Location QuickLaunch
    Add-PnPNavigationNode -Title "Security"   -Url "$SiteUrl/SitePages/Security.aspx"   -Location QuickLaunch

    # ---------------------------
    # LISTS HEADER
    # ---------------------------
    $listsNode = Add-PnPNavigationNode `
    -Title "Lists" `
    -Url "$SiteUrl/_layouts/15/viewlsts.aspx" `
    -Location QuickLaunch

    $listsNode = Get-PnPNavigationNode -Location QuickLaunch | Where-Object { $_.Title -eq "Lists" }

    # ---------------------------
    # CHILD ITEMS
    # ---------------------------
    Add-PnPNavigationNode `
        -Title "Client Network" `
        -Url "$SiteUrl/Lists/Client Network" `
        -Location QuickLaunch `
        -Parent $listsNode.Id

    Add-PnPNavigationNode `
        -Title "Client Software" `
        -Url "$SiteUrl/Lists/Client Software" `
        -Location QuickLaunch `
        -Parent $listsNode.Id

    Write-Host "✅ Navigation rebuilt cleanly" -ForegroundColor Green
}


function Ensure-Field {
    param(
        [string]$ListName,
        [string]$InternalName,
        [string]$DisplayName,
        [string]$Type
    )

    $field = Get-PnPField -List $ListName | Where-Object {
        $_.InternalName -eq $InternalName
    }

    if (-not $field) {
        Write-Host "Creating field: $DisplayName"

        Add-PnPField `
            -List $ListName `
            -InternalName $InternalName `
            -DisplayName $DisplayName `
            -Type $Type | Out-Null
    }
    else {
        Write-Host "✅ Field exists: $DisplayName"
    }
}

function Ensure-DefaultView {
    param(
        [string]$ListName,
        [string]$ViewName,
        [string[]]$Fields,
        [string]$SortField = "Title",
        [bool]$SortAscending = $true
    )

    Write-Host "Configuring view: $ListName -> $ViewName" -ForegroundColor Cyan

    # Get existing view
    $view = Get-PnPView -List $ListName | Where-Object { $_.Title -eq $ViewName }

    if (-not $view) {
        Write-Host "Creating view: $ViewName"

        $view = Add-PnPView `
            -List $ListName `
            -Title $ViewName `
            -Fields $Fields `
            -SetAsDefault `
            -Query "<OrderBy><FieldRef Name='$SortField' Ascending='$SortAscending' /></OrderBy>" `
            -ErrorAction Stop
    }
    else {
        Write-Host "Updating view: $ViewName"

        Set-PnPView `
            -List $ListName `
            -Identity $view.Id `
            -Fields $Fields `
            -SetAsDefault `
            -Query "<OrderBy><FieldRef Name='$SortField' Ascending='$SortAscending' /></OrderBy>" `
            -ErrorAction Stop
    }

    # -------------------------------
    # Hide ALL system fields explicitly
    # -------------------------------
    $systemFields = @(
        "Modified",
        "Created",
        "Author",
        "Editor",
        "_UIVersionString",
        "Attachments",
        "GUID"
    )

    # -------------------------------
    # Remove 'All Items' view (optional)
    # -------------------------------
    $defaultView = Get-PnPView -List $ListName | Where-Object { $_.Title -eq "All Items" }

    if ($defaultView) {
        try {
            Remove-PnPView -List $ListName -Identity $defaultView.Id -Force -ErrorAction Stop
            Write-Host "Removed 'All Items' view"
        }
        catch {
            Write-Host "Could not remove 'All Items' (safe to ignore)"
        }
    }

    Write-Host "✅ View configured: $ViewName" -ForegroundColor Green
}


function Ensure-ClientList {
    param(
        [string]$ListName
    )

    try {
        Get-PnPList -Identity $ListName -ErrorAction Stop | Out-Null
        Write-Host "✅ List exists: $ListName"
    }
    catch {
        Write-Host "Creating list: $ListName"

        New-PnPList -Title $ListName -Template GenericList | Out-Null
    }
}

function Connect-ZZ {
    param([string]$Url)

    Connect-PnPOnline -Url $Url `
        -ClientId $ClientId `
        -Tenant $Tenant `
        -CertificatePath $PfxPath `
        -CertificatePassword $PfxPassword
}

function Invoke-WithRetry {
    param(
        [scriptblock]$ScriptBlock,
        [string]$ActionName = "Operation",
        [int]$MaxRetries = $RetryCount,
        [int]$DelaySeconds = $RetryDelaySeconds
    )

    for ($i = 1; $i -le $MaxRetries; $i++) {
        try {
            return & $ScriptBlock
        }
        catch {
            if ($i -ge $MaxRetries) {
                throw "FAILED: $ActionName after $MaxRetries attempts. Last error: $($_.Exception.Message)"
            }

            Write-Host "Retry $i/$MaxRetries for: $ActionName" -ForegroundColor Yellow
            Start-Sleep -Seconds $DelaySeconds
        }
    }
}

function Test-HasValue {
    param($Value)

    if ($null -eq $Value) { return $false }
    if ($Value -is [string] -and $Value.Trim() -eq "") { return $false }
    return $true
}

function Convert-ToSlug {
    param([string]$Text)

    $slug = $Text.Trim()
    $slug = $slug -replace '[^a-zA-Z0-9 ]',''
    $slug = $slug -replace '\s+','-'
    $slug = $slug -replace '-{2,}','-'
    $slug = $slug.Trim('-')

    if (-not (Test-HasValue $slug)) {
        throw "Unable to create slug from '$Text'"
    }

    return $slug
}

function Escape-Html {
    param($Text)

    if ($null -eq $Text) { return "" }

    $escaped = [string]$Text
    $escaped = $escaped -replace '&','&amp;'
    $escaped = $escaped -replace '<','&lt;'
    $escaped = $escaped -replace '>','&gt;'
    $escaped = $escaped -replace '"','&quot;'
    return $escaped
}

function Get-Client {
    param([string]$Value)

    Connect-ZZ $ControlSiteUrl
    $items = Get-PnPListItem -List $ClientRegistryList -PageSize 2000

    return $items | Where-Object {
        $_.FieldValues["ClientID"] -eq $Value -or
        $_.FieldValues["ClientName"] -eq $Value
    } | Select-Object -First 1
}

function Ensure-ClientRegistryFields {
    Connect-ZZ $ControlSiteUrl

    $siteUrlField = Get-PnPField -List $ClientRegistryList | Where-Object {
        $_.InternalName -eq "SiteURL" -or $_.Title -eq "SiteURL"
    }

    if (-not $siteUrlField) {
        Add-PnPField -List $ClientRegistryList -InternalName "SiteURL" -DisplayName "SiteURL" -Type URL | Out-Null
        Write-Host "Created SiteURL field" -ForegroundColor Green
    }

    $lastSyncedField = Get-PnPField -List $ClientRegistryList | Where-Object {
        $_.InternalName -eq "LastSynced" -or $_.Title -eq "LastSynced"
    }

    if (-not $lastSyncedField) {
        Add-PnPField -List $ClientRegistryList -InternalName "LastSynced" -DisplayName "LastSynced" -Type DateTime | Out-Null
        Write-Host "Created LastSynced field" -ForegroundColor Green
    }
}

function Update-ClientRegistryMetadata {
    param(
        [int]$ClientListItemId,
        [string]$SiteUrl
    )

    Connect-ZZ $ControlSiteUrl

    Set-PnPListItem -List $ClientRegistryList -Identity $ClientListItemId -Values @{
        "SiteURL"    = $SiteUrl
        "LastSynced" = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
    } | Out-Null

    Write-Host "Client Registry metadata updated" -ForegroundColor Green
}

function Wait-ForSiteReady {
    param([string]$SiteUrl)

    Invoke-WithRetry -ActionName "Wait for tenant site existence" -ScriptBlock {
        Connect-ZZ $TenantAdminUrl
        $site = Get-PnPTenantSite -Url $SiteUrl -ErrorAction Stop
        if (-not $site) { throw "Tenant site not yet available" }
        return $site
    } | Out-Null

    Invoke-WithRetry -ActionName "Wait for web connection" -ScriptBlock {
        Connect-ZZ $SiteUrl
        $web = Get-PnPWeb -ErrorAction Stop
        if (-not $web) { throw "Web not ready" }
        return $web
    } | Out-Null

    Invoke-WithRetry -ActionName "Wait for Site Pages library" -ScriptBlock {
        Connect-ZZ $SiteUrl
        $list = Get-PnPList -Identity "Site Pages" -ErrorAction Stop
        if (-not $list) { throw "Site Pages library not ready" }
        return $list
    } | Out-Null
}

function Ensure-Site {
    param(
        [string]$SiteTitle,
        [string]$SiteUrl
    )

    Connect-ZZ $TenantAdminUrl

    $siteExists = $false
    try {
        Get-PnPTenantSite -Url $SiteUrl -ErrorAction Stop | Out-Null
        $siteExists = $true
        Write-Host "Site already exists: $SiteUrl" -ForegroundColor Green
    }
    catch {
        $siteExists = $false
    }

    if (-not $siteExists) {
        Write-Host "Creating site: $SiteUrl" -ForegroundColor Cyan
        try {
            New-PnPSite -Type TeamSiteWithoutMicrosoft365Group `
                -Title $SiteTitle `
                -Url $SiteUrl `
                -Owner $SiteOwner `
                -ErrorAction Stop | Out-Null
        }
        catch {
            Write-Host "New-PnPSite returned: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }

    Wait-ForSiteReady -SiteUrl $SiteUrl
}

function Ensure-HubAssociation {
    param(
        [string]$SiteUrl,
        [string]$HubSiteUrl
    )

    Connect-ZZ $TenantAdminUrl

    try {
        Add-PnPHubSiteAssociation -Site $SiteUrl -HubSite $HubSiteUrl -ErrorAction Stop | Out-Null
        Write-Host "Hub associated" -ForegroundColor Green
    }
    catch {
        if ($_.Exception.Message -match "already associated") {
            Write-Host "Hub association already present" -ForegroundColor Green
        }
        else {
            Write-Host "Hub association notice: $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

function Ensure-SectionOneExists {
    param([string]$PageName)

    $components = @()
    try {
        $components = @(Get-PnPPageComponent -Page $PageName)
    }
    catch {
        $components = @()
    }

    if ($components.Count -eq 0) {
        Add-PnPPageSection -Page $PageName -SectionTemplate OneColumn -Order 1 -ErrorAction Stop | Out-Null
    }
}

function Ensure-OverviewPage {
    param(
        [string]$PageName,
        [string]$PageTitle,
        [string]$ClientName,
        [string]$ClientRecordId,
        [string]$PrimaryDomain,
        [string]$Status,
        [string]$SupportTier
    )

    try {
        Get-PnPPage -Identity $PageName -ErrorAction Stop | Out-Null
        Write-Host "Page exists: $PageName" -ForegroundColor Green
        return
    }
    catch {}

    Invoke-WithRetry -ActionName "Create Overview page" -ScriptBlock {
        Add-PnPPage -Name $PageName -Title $PageTitle -LayoutType Article -ErrorAction Stop | Out-Null
        Ensure-SectionOneExists -PageName $PageName

        $summaryHtml = @"
<p><strong>ZZ-AUTO-CLIENT-SCAFFOLD</strong></p>
<h2>Client Summary</h2>
<ul>
<li><strong>Name:</strong> $(Escape-Html $ClientName)</li>
<li><strong>Client ID:</strong> $(Escape-Html $ClientRecordId)</li>
<li><strong>Domain:</strong> $(Escape-Html $PrimaryDomain)</li>
<li><strong>Status:</strong> $(Escape-Html $Status)</li>
<li><strong>Support Tier:</strong> $(Escape-Html $SupportTier)</li>
</ul>
"@

        $servicesHtml = @"
<p><strong>ZZ-AUTO-SERVICE-LINKS</strong></p>
<h3>Services</h3>
<p>Content injected via Update-ClientHub.ps1</p>
"@

        $contactsHtml = @"
<p><strong>ZZ-AUTO-CONTACTS</strong></p>
<h3>Contacts</h3>
<p>Content injected via Update-ClientHub.ps1</p>
"@

        $notesHtml = @"
<p><strong>ZZ-AUTO-PAGE-SCAFFOLD</strong></p>
<h3>Notes</h3>
<p>Manual content allowed here.</p>
"@

        Add-PnPPageTextPart -Page $PageName -Section 1 -Column 1 -Text $summaryHtml -ErrorAction Stop | Out-Null
        Add-PnPPageTextPart -Page $PageName -Section 1 -Column 1 -Text $servicesHtml -ErrorAction Stop | Out-Null
        Add-PnPPageTextPart -Page $PageName -Section 1 -Column 1 -Text $contactsHtml -ErrorAction Stop | Out-Null
        Add-PnPPageTextPart -Page $PageName -Section 1 -Column 1 -Text $notesHtml -ErrorAction Stop | Out-Null

        Set-PnPPage -Identity $PageName -Publish -ErrorAction Stop | Out-Null
    }

    Write-Host "Overview page created" -ForegroundColor Green
}

function Ensure-ScaffoldPage {
    param(
        [string]$PageName,
        [string]$PageTitle,
        [string]$IntroHtml
    )

    try {
        Get-PnPPage -Identity $PageName -ErrorAction Stop | Out-Null
        Write-Host "Page exists: $PageName" -ForegroundColor Green
        return
    }
    catch {}

    Invoke-WithRetry -ActionName "Create page $PageName" -ScriptBlock {
        Add-PnPPage -Name $PageName -Title $PageTitle -LayoutType Article -ErrorAction Stop | Out-Null
        Ensure-SectionOneExists -PageName $PageName

        $manualNotesHtml = @"
<p><strong>ZZ-AUTO-PAGE-SCAFFOLD</strong></p>
<h3>Manual Notes</h3>
<p>Use this area for client-specific notes, exceptions, context, and additional commentary.</p>
"@

        Add-PnPPageTextPart -Page $PageName -Section 1 -Column 1 -Text $IntroHtml -ErrorAction Stop | Out-Null
        Add-PnPPageTextPart -Page $PageName -Section 1 -Column 1 -Text $manualNotesHtml -ErrorAction Stop | Out-Null

        Set-PnPPage -Identity $PageName -Publish -ErrorAction Stop | Out-Null
    }

    Write-Host "Page created: $PageName" -ForegroundColor Green
}

function Ensure-PageFromTemplate {
    param(
        [string]$TemplateServerRelativeUrl,
        [string]$TargetPageName
    )

    # Target page already exists
    try {
        Get-PnPPage -Identity $TargetPageName -ErrorAction Stop | Out-Null
        Write-Host "Page exists: $TargetPageName" -ForegroundColor Green
        return
    }
    catch {}

    $templateFileName = Split-Path $TemplateServerRelativeUrl -Leaf
    $targetSitePagesFolderServerRelativeUrl = "$((Get-PnPWeb).ServerRelativeUrl)/SitePages"
    $copiedTemplateServerRelativeUrl = "$targetSitePagesFolderServerRelativeUrl/$templateFileName"
    $targetPageServerRelativeUrl = "$targetSitePagesFolderServerRelativeUrl/$TargetPageName"

    # If copied template already exists on target from a previous partial run, just rename it
    $templateAlreadyCopied = $false
    try {
        Get-PnPFile -Url $copiedTemplateServerRelativeUrl -ErrorAction Stop | Out-Null
        $templateAlreadyCopied = $true
    }
    catch {
        $templateAlreadyCopied = $false
    }

    if (-not $templateAlreadyCopied) {
        Write-Host "Copying template $templateFileName to target site..." -ForegroundColor Cyan

        # Copy from ZZ-Control to the target site's Site Pages folder.
        # Copy-PnPFile supports cross-site copy, but target must be a folder when copying across sites. [1](https://pnp.github.io/powershell/cmdlets/Copy-PnPFile.html)
        Connect-ZZ $ControlSiteUrl
        Invoke-WithRetry -ActionName "Copy $templateFileName to target site" -ScriptBlock {
            Copy-PnPFile `
                -SourceUrl $TemplateServerRelativeUrl `
                -TargetUrl $targetSitePagesFolderServerRelativeUrl `
                -Overwrite `
                -Force `
                -ErrorAction Stop | Out-Null
        } | Out-Null

        # Return to target site connection
        Connect-ZZ $siteUrl
    }
    else {
        Write-Host "Template file already present on target site: $templateFileName" -ForegroundColor Yellow
    }

    # Rename in the current site so the target file name is correct.
    # Rename-PnPFile supports renaming a file in its current location. [2](https://pnp.github.io/powershell/cmdlets/Rename-PnPFile.html)
    Write-Host "Renaming $templateFileName to $TargetPageName ..." -ForegroundColor Cyan
    Invoke-WithRetry -ActionName "Rename $templateFileName to $TargetPageName" -ScriptBlock {
        Rename-PnPFile -ServerRelativeUrl $copiedTemplateServerRelativeUrl `
            -TargetFileName $TargetPageName `
            -OverwriteIfAlreadyExists `
            -Force `
            -ErrorAction Stop | Out-Null
    } | Out-Null

    # Publish cloned page
    Invoke-WithRetry -ActionName "Publish $TargetPageName" -ScriptBlock {
        Set-PnPPage -Identity $TargetPageName -Publish -ErrorAction Stop | Out-Null
    } | Out-Null

    Write-Host "Template page created: $TargetPageName" -ForegroundColor Green
}

function Set-HomePageSafe {
    param([string]$PageServerRelativeUrl)

    Invoke-WithRetry -ActionName "Set homepage" -ScriptBlock {
        Set-PnPHomePage -RootFolderRelativeUrl $PageServerRelativeUrl -ErrorAction Stop | Out-Null
    }
}

function Ensure-Nav {
    param(
        [string]$Title,
        [string]$Url
    )

    Invoke-WithRetry -ActionName "Ensure nav $Title" -ScriptBlock {
        $existing = @(Get-PnPNavigationNode -Location QuickLaunch -ErrorAction Stop) | Where-Object {
            $_.Title -eq $Title -or $_.Url -eq $Url
        }

        if (-not $existing) {
            Add-PnPNavigationNode -Title $Title -Url $Url -Location QuickLaunch -ErrorAction Stop | Out-Null
            Write-Host "Added nav: $Title"
        }
        else {
            Write-Host "Nav exists: $Title"
        }
    }
}

# MAIN
$client = Get-Client $ClientLookupValue
if (-not $client) {
    throw "Client not found: $ClientLookupValue"
}

$clientName = [string]$client.FieldValues["ClientName"]
$clientRegistryId = [string]$client.FieldValues["ClientID"]
$clientListItemId = [int]$client.Id
$primaryDomain = [string]$client.FieldValues["PrimaryDomain"]
$status = [string]$client.FieldValues["Status"]
$supportTier = [string]$client.FieldValues["SupportTier"]

$siteSlug = Convert-ToSlug $clientName
$siteName = "CLIENT-$siteSlug"
$siteUrl = "$TenantRootUrl/sites/$siteName"

Write-Host "Client: $clientName" -ForegroundColor Cyan
Write-Host "Target Site: $siteUrl" -ForegroundColor Cyan

Ensure-ClientRegistryFields
Ensure-Site -SiteTitle $siteName -SiteUrl $siteUrl
Update-ClientRegistryMetadata -ClientListItemId $clientListItemId -SiteUrl $siteUrl
Ensure-HubAssociation -SiteUrl $siteUrl -HubSiteUrl $HubSiteUrl

Connect-ZZ $siteUrl

# -------------------------
# Lists
# -------------------------

Ensure-ClientList -ListName "Client Network"
Ensure-ClientList -ListName "Client Software"

# -------------------------
# Fields - Client Network
# -------------------------

Ensure-Field "Client Network" "NetworkType" "Network Type" "Choice"
Ensure-Field "Client Network" "Location" "Location" "Text"
Ensure-Field "Client Network" "RouterIP" "Router IP" "Text"
Ensure-Field "Client Network" "NetworkRange" "Network Range" "Text"
Ensure-Field "Client Network" "SubnetMask" "Subnet Mask" "Text"
Ensure-Field "Client Network" "DHCPServer" "DHCP Server" "Text"

# -------------------------
# Fields - Client Software
# -------------------------

Ensure-Field "Client Software" "Vendor" "Vendor" "Text"
Ensure-Field "Client Software" "SupportedByZZ" "Supported By ZZ" "Boolean"
Ensure-Field "Client Software" "VendorPhone" "Vendor Phone" "Text"
Ensure-Field "Client Software" "VendorEmail" "Vendor Email" "Text"

# -------------------------------
# Configure views
# -------------------------------

Start-Sleep -Seconds 5 # Short delay to ensure fields are fully provisioned before configuring views    


Ensure-DefaultView `
    -ListName "Client Network" `
    -ViewName "Client Networks" `
    -Fields @(
        "Title",
        "NetworkType",
        "Location",
        "RouterIP",
        "NetworkRange",
        "SubnetMask",
        "DHCPServer"
    ) `
    -SortField "Location" `
    -SortAscending $true


Ensure-DefaultView `
    -ListName "Client Software" `
    -ViewName "Client Software" `
    -Fields @(
        "Title",
        "Vendor",
        "SupportedByZZ",
        "VendorPhone",
        "VendorEmail"
    ) `
    -SortField "Title" `
    -SortAscending $true


# Overview stays generated by script
Ensure-OverviewPage `
    -PageName "Overview.aspx" `
    -PageTitle "Overview" `
    -ClientName $clientName `
    -ClientRecordId $clientRegistryId `
    -PrimaryDomain $primaryDomain `
    -Status $status `
    -SupportTier $supportTier

# Clone template pages for Network and Software
Ensure-PageFromTemplate `
    -TemplateServerRelativeUrl $NetworkTemplateServerRelativeUrl `
    -TargetPageName "Network.aspx"

Ensure-PageFromTemplate `
    -TemplateServerRelativeUrl $SoftwareTemplateServerRelativeUrl `
    -TargetPageName "Software.aspx"


# Keep other pages as scaffold/manual pages
$projectsIntro = @"
<p><strong>ZZ-AUTO-PAGE-SCAFFOLD</strong></p>
<h2>Projects</h2>
<p>Manual project documentation goes here.</p>
"@

$operationsIntro = @"
<p><strong>ZZ-AUTO-PAGE-SCAFFOLD</strong></p>
<h2>Operations</h2>
<p>Manual operations documentation goes here.</p>
"@

$securityIntro = @"
<p><strong>ZZ-AUTO-PAGE-SCAFFOLD</strong></p>
<h2>Security</h2>
<p>Manual security documentation goes here.</p>
"@

Ensure-ScaffoldPage -PageName "Projects.aspx" -PageTitle "Projects" -IntroHtml $projectsIntro
Ensure-ScaffoldPage -PageName "Operations.aspx" -PageTitle "Operations" -IntroHtml $operationsIntro
Ensure-ScaffoldPage -PageName "Security.aspx" -PageTitle "Security" -IntroHtml $securityIntro

Set-HomePageSafe -PageServerRelativeUrl "SitePages/Overview.aspx"

<#Ensure-Nav -Title "Overview" -Url "$siteUrl/SitePages/Overview.aspx"
Ensure-Nav -Title "Network" -Url "$siteUrl/SitePages/Network.aspx"
Ensure-Nav -Title "Software" -Url "$siteUrl/SitePages/Software.aspx"
Ensure-Nav -Title "Projects" -Url "$siteUrl/SitePages/Projects.aspx"
Ensure-Nav -Title "Operations" -Url "$siteUrl/SitePages/Operations.aspx"
Ensure-Nav -Title "Security" -Url "$siteUrl/SitePages/Security.aspx"
#>
Configure-CleanNavigation -SiteUrl $siteUrl




Write-Host "Provision complete: $siteUrl" -ForegroundColor Green