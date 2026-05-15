param(
    [Parameter(Mandatory)]
    [string]$ClientLookupValue
)

Write-Host "==== Starting Client Hub Update ====" -ForegroundColor Cyan

# AUTH
$Tenant = "zahe.onmicrosoft.com"
$ClientId = "01e1b71f-cbcb-48df-a076-871aa4ba10d9"
$PfxPath = ".\ZaheZone-PnP-Projects.pfx"
$PfxPasswordPlain = "UseA-LongRandomPasswordHere"
$PfxPassword = ConvertTo-SecureString $PfxPasswordPlain -AsPlainText -Force

Set-Variable -Name ClientId -Value $ClientId -Option ReadOnly -Force

# CONFIG
$TenantRootUrl = "https://zahe.sharepoint.com"
$ControlSiteUrl = "https://zahe.sharepoint.com/sites/ZZ-Control"

$ClientRegistryList = "Client Registry"
$ServiceMappingList = "Client Service Mapping"
$ContactsList = "Client Contacts"

$RetryCount = 24
$RetryDelaySeconds = 5

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

function Get-ListFieldDefinition {
    param(
        [string]$ListName,
        [string]$FieldName
    )

    $fields = @(Get-PnPField -List $ListName)
    return $fields | Where-Object {
        $_.InternalName -eq $FieldName -or $_.Title -eq $FieldName
    } | Select-Object -First 1
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

function Get-ClientSiteUrl {
    param($Client)

    $siteField = $Client.FieldValues["SiteURL"]

    if ($null -ne $siteField) {
        if ($siteField -is [string]) {
            if (Test-HasValue $siteField) {
                return $siteField
            }
        }

        if ($siteField.PSObject -and $siteField.PSObject.Properties["Url"]) {
            if (Test-HasValue $siteField.Url) {
                return [string]$siteField.Url
            }
        }
    }

    $clientName = [string]$Client.FieldValues["ClientName"]
    $siteSlug = Convert-ToSlug $clientName
    return "$TenantRootUrl/sites/CLIENT-$siteSlug"
}

function Get-ServiceMappingsForClient {
    param([int]$ClientListItemId)

    Connect-ZZ $ControlSiteUrl
    $items = @(Get-PnPListItem -List $ServiceMappingList -PageSize 2000)

    return $items | Where-Object {
        $lookup = $_.FieldValues["Client"]
        $lookup -and $lookup.LookupId -eq $ClientListItemId
    }
}

function Get-ContactsForClient {
    param([int]$ClientListItemId)

    Connect-ZZ $ControlSiteUrl
    $items = @(Get-PnPListItem -List $ContactsList -PageSize 2000)

    return $items | Where-Object {
        $lookup = $_.FieldValues["Client"]
        $lookup -and $lookup.LookupId -eq $ClientListItemId
    }
}

function Wait-ForOverviewPage {
    param([string]$SiteUrl)

    Invoke-WithRetry -ActionName "Wait for Overview.aspx" -ScriptBlock {
        Connect-ZZ $SiteUrl
        $page = Get-PnPPage -Identity "Overview.aspx" -ErrorAction Stop
        if (-not $page) { throw "Overview.aspx not found" }
        return $page
    } | Out-Null
}

function Get-OverviewTextComponents {
    $components = @()

    try {
        $all = @(Get-PnPPageComponent -Page "Overview.aspx")
    }
    catch {
        return @()
    }

    foreach ($c in $all) {
        $hasText = $false
        $hasInstanceId = $false

        if ($null -ne $c.PSObject.Properties["Text"]) { $hasText = $true }
        if ($null -ne $c.PSObject.Properties["InstanceId"]) { $hasInstanceId = $true }

        if ($hasText -and $hasInstanceId) {
            $components += $c
        }
    }

    return $components
}

function Get-ExistingManualNotesHtml {
    param([array]$TextComponents)

    foreach ($component in $TextComponents) {
        $text = [string]$component.Text

        if ($text -match 'ZZ-AUTO-PAGE-SCAFFOLD') {
            $idx = $text.IndexOf('ZZ-AUTO-PAGE-SCAFFOLD')
            if ($idx -ge 0) {
                $prefixStart = $text.LastIndexOf('<p><strong>', $idx)
                if ($prefixStart -ge 0) {
                    return $text.Substring($prefixStart)
                }
                else {
                    return $text.Substring($idx)
                }
            }
        }
    }

    return @"
<p><strong>ZZ-AUTO-PAGE-SCAFFOLD</strong></p>
<h3>Notes</h3>
<p>Manual content allowed here.</p>
"@
}

function Build-ServicesHtml {
    param([array]$Mappings)

    if (-not $Mappings -or $Mappings.Count -eq 0) {
        return @"
<p><strong>ZZ-AUTO-SERVICE-LINKS</strong></p>
<h3>Services</h3>
<ul>
  <li>None recorded</li>
</ul>
"@
    }

    $rows = foreach ($m in $Mappings) {
        $serviceLookup = $m.FieldValues["Service"]
        $serviceName = ""
        if ($serviceLookup) {
            $serviceName = [string]$serviceLookup.LookupValue
        }

        $externalId = [string]$m.FieldValues["ExternalID"]

        if (Test-HasValue $serviceName) {
            "<li><strong>$(Escape-Html $serviceName)</strong> - External ID: $(Escape-Html $externalId)</li>"
        }
    }

    $rows = $rows | Where-Object { $_ } | Sort-Object
    if (-not $rows -or $rows.Count -eq 0) {
        $rows = @("  <li>None recorded</li>")
    }

    return @"
<p><strong>ZZ-AUTO-SERVICE-LINKS</strong></p>
<h3>Services</h3>
<ul>
$($rows -join "`n")
</ul>
"@
}

function Build-ContactsHtml {
    param([array]$Contacts)

    if (-not $Contacts -or $Contacts.Count -eq 0) {
        return @"
<p><strong>ZZ-AUTO-CONTACTS</strong></p>
<h3>Contacts</h3>
<ul>
  <li>None recorded</li>
</ul>
"@
    }

    $rows = foreach ($c in $Contacts) {
        $name  = [string]$c.FieldValues["Name"]
        $role  = [string]$c.FieldValues["Role"]
        $email = [string]$c.FieldValues["Email"]

        $parts = @()

        if (Test-HasValue $name)  { $parts += "<strong>$(Escape-Html $name)</strong>" }
        if (Test-HasValue $role)  { $parts += "$(Escape-Html $role)" }
        if (Test-HasValue $email) { $parts += "$(Escape-Html $email)" }

        if ($parts.Count -gt 0) {
            "<li>$($parts -join ' - ')</li>"
        }
    }

    $rows = $rows | Where-Object { $_ } | Sort-Object
    if (-not $rows -or $rows.Count -eq 0) {
        $rows = @("  <li>None recorded</li>")
    }

    return @"
<p><strong>ZZ-AUTO-CONTACTS</strong></p>
<h3>Contacts</h3>
<ul>
$($rows -join "`n")
</ul>
"@
}

function Build-ClientSummaryHtml {
    param($Client)

    $clientName = [string]$Client.FieldValues["ClientName"]
    $clientRecordId = [string]$Client.FieldValues["ClientID"]
    $primaryDomain = [string]$Client.FieldValues["PrimaryDomain"]
    $status = [string]$Client.FieldValues["Status"]
    $supportTier = [string]$Client.FieldValues["SupportTier"]

    return @"
<p><strong>ZZ-AUTO-CLIENT-SCAFFOLD</strong></p>

<h2>Client Summary</h2>
<ul>
<li><strong>Name:</strong> $(Escape-Html $clientName)</li>
<li><strong>Client ID:</strong> $(Escape-Html $clientRecordId)</li>
<li><strong>Domain:</strong> $(Escape-Html $primaryDomain)</li>
<li><strong>Status:</strong> $(Escape-Html $status)</li>
<li><strong>Support Tier:</strong> $(Escape-Html $supportTier)</li>
</ul>
"@
}

function Remove-OverviewAutoComponents {
    param([array]$TextComponents)

    foreach ($component in $TextComponents) {
        $text = [string]$component.Text

        if ($text -match 'ZZ-AUTO-CLIENT-SCAFFOLD' -or
            $text -match 'ZZ-AUTO-SERVICE-LINKS' -or
            $text -match 'ZZ-AUTO-CONTACTS') {

            if ($component.InstanceId) {
                Remove-PnPPageComponent -Page "Overview.aspx" -InstanceId $component.InstanceId -Force -ErrorAction Stop | Out-Null
                Write-Host "Removed existing automation component: $($component.InstanceId)" -ForegroundColor Yellow
            }
        }
    }
}

function Ensure-OverviewTextBlock {
    param([string]$TextHtml)

    Invoke-WithRetry -ActionName "Add Overview automation block" -ScriptBlock {
        Add-PnPPageTextPart -Page "Overview.aspx" -Section 1 -Column 1 -Text $TextHtml -ErrorAction Stop | Out-Null
        Set-PnPPage -Identity "Overview.aspx" -Publish -ErrorAction Stop | Out-Null
    }
}

function Update-ClientRegistryMetadata {
    param(
        [int]$ClientListItemId,
        [string]$SiteUrl
    )

    Connect-ZZ $ControlSiteUrl

    $updateValues = @{}

    $siteField = Get-ListFieldDefinition -ListName $ClientRegistryList -FieldName "SiteURL"
    if ($siteField) {
        $updateValues[$siteField.InternalName] = $SiteUrl
    }

    $lastSyncedField = Get-ListFieldDefinition -ListName $ClientRegistryList -FieldName "LastSynced"
    if ($lastSyncedField) {
        if ($lastSyncedField.TypeAsString -eq "DateTime") {
            $updateValues[$lastSyncedField.InternalName] = (Get-Date).ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ")
        }
    }

    if ($updateValues.Count -gt 0) {
        Set-PnPListItem -List $ClientRegistryList -Identity $ClientListItemId -Values $updateValues | Out-Null
        Write-Host "Client Registry metadata updated" -ForegroundColor Green
    }
}

# MAIN
$client = Get-Client $ClientLookupValue
if (-not $client) {
    throw "Client not found: $ClientLookupValue"
}

$clientName = [string]$client.FieldValues["ClientName"]
$clientListItemId = [int]$client.Id
$siteUrl = Get-ClientSiteUrl -Client $client

Write-Host "Client: $clientName" -ForegroundColor Cyan
Write-Host "Target Site: $siteUrl" -ForegroundColor Cyan

$mappings = @(Get-ServiceMappingsForClient -ClientListItemId $clientListItemId)
$contacts = @(Get-ContactsForClient -ClientListItemId $clientListItemId)

Wait-ForOverviewPage -SiteUrl $siteUrl
Connect-ZZ $siteUrl

$textComponents = @(Get-OverviewTextComponents)
$manualNotesHtml = Get-ExistingManualNotesHtml -TextComponents $textComponents

$clientSummaryHtml = Build-ClientSummaryHtml -Client $client
$servicesHtml = Build-ServicesHtml -Mappings $mappings
$contactsHtml = Build-ContactsHtml -Contacts $contacts

$newOverviewHtml = @"
$clientSummaryHtml

$servicesHtml

$contactsHtml

$manualNotesHtml
"@

Remove-OverviewAutoComponents -TextComponents $textComponents
Ensure-OverviewTextBlock -TextHtml $newOverviewHtml
Update-ClientRegistryMetadata -ClientListItemId $clientListItemId -SiteUrl $siteUrl

Write-Host "Update complete: $siteUrl" -ForegroundColor Green