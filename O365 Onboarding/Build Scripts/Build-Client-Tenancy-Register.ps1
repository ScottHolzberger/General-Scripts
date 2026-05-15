# ==========================================================
# ZaheZone – Client Tenancy Register
# SharePoint Build Script
# App-only Certificate Authentication
# ==========================================================

# -------------------------------
# AUTH
# -------------------------------
[string]$SiteUrl  = "https://zahe.sharepoint.com/sites/ZaheZoneClientTenancyRegister"
[string]$Tenant   = "zahe.onmicrosoft.com"
[string]$ClientId = "01e1b71f-cbcb-48df-a076-871aa4ba10d9"
[string]$PfxPath  = "$PSScriptRoot\ZaheZone-PnP-Projects.pfx"
[string]$PfxPasswordPlain = "UseA-LongRandomPasswordHere"

# -------------------------------
# CONNECT
# -------------------------------
Connect-PnPOnline `
    -Url $SiteUrl `
    -Tenant $Tenant `
    -ClientId $ClientId `
    -CertificatePath $PfxPath `
    -CertificatePassword (ConvertTo-SecureString $PfxPasswordPlain -AsPlainText -Force)

Write-Host "✅ Connected to SharePoint via app-only certificate auth"

# ==========================================================
# HELPERS
# ==========================================================

# -------------------------------
# Ensure List Exists
# -------------------------------
function Ensure-List {
    param(
        [string]$Title,
        [string]$Description
    )

    $list = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    if (-not $list) {
        Write-Host "Creating list: $Title"

        New-PnPList `
            -Title $Title `
            -Template GenericList `
            -EnableVersioning `
            -ErrorAction Stop | Out-Null

        Set-PnPList `
            -Identity $Title `
            -Description $Description `
            -ErrorAction Stop

        Write-Host "✅ List created: $Title"
    }
    else {
        Write-Host "List exists: $Title"
    }
}

# -------------------------------
# Ensure Standard Field Exists
# -------------------------------
function Ensure-Field {
    param(
        [string]$List,
        [string]$InternalName,
        [hashtable]$Options
    )

    $field = Get-PnPField -List $List -Identity $InternalName -ErrorAction SilentlyContinue
    if (-not $field) {
        try {
            Add-PnPField `
                -List $List `
                -InternalName $InternalName `
                @Options `
                -ErrorAction Stop

            Write-Host "  + Field added: $InternalName"
        }
        catch {
            Write-Error "❌ Failed to add field '$InternalName' to list '$List'"
            throw
        }
    }
}

# -------------------------------
# Ensure Lookup Field Exists (REST – works everywhere)
# -------------------------------
function Ensure-LookupField {
    param(
        [string]$List,
        [string]$InternalName,
        [string]$DisplayName,
        [string]$LookupList,
        [string]$LookupField = "Title",
        [bool]$Required = $false
    )

    # Already on list?
    $existing = Get-PnPField -List $List -Identity $InternalName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  + Lookup field exists: $InternalName"
        return
    }

    # Resolve list IDs
    $targetList   = Get-PnPList -Identity $List -ErrorAction Stop
    $lookupListId = (Get-PnPList -Identity $LookupList -ErrorAction Stop).Id.Guid

    $requiredValue = if ($Required) { "true" } else { "false" }

    # REST payload
    $payload = @{
        "__metadata" = @{ "type" = "SP.FieldLookup" }
        "Title"          = $DisplayName
        "InternalName"   = $InternalName
        "LookUpList"     = $lookupListId
        "LookupField"    = $LookupField
        "Required"       = $Required
    } | ConvertTo-Json -Depth 5

    $endpoint = "$($targetList.RootFolder.ServerRelativeUrl)/Fields"

    try {
        Invoke-PnPSPRestMethod `
            -Method POST `
            -Url $endpoint `
            -Content $payload `
            -ContentType "application/json;odata=verbose" `
            -ErrorAction Stop

        Write-Host "  + Lookup field created via REST: $InternalName"
    }
    catch {
        Write-Error "❌ Failed to create lookup field '$InternalName' via REST"
        throw
    }
}

# ==========================================================
# LIST 1 — Client Tenancies
# ==========================================================
Ensure-List `
    -Title "Client Tenancies" `
    -Description "Authoritative register of Microsoft 365 client tenants managed by ZaheZone."

Ensure-Field "Client Tenancies" "TenantID" @{
    Type        = "Text"
    DisplayName = "Tenant ID"
    Required    = $true
}

Ensure-Field "Client Tenancies" "PrimaryDomain" @{
    Type        = "Text"
    DisplayName = "Primary Domain"
}

Ensure-Field "Client Tenancies" "LicensingTier" @{
    Type        = "Choice"
    DisplayName = "Licensing Tier"
    Choices     = @("Basic","Business Premium","E3","E5","Mixed")
}

Ensure-Field "Client Tenancies" "O365Status" @{
    Type        = "Choice"
    DisplayName = "O365 Status"
    Choices     = @("Onboarding","BAU","Offboarded")
}

Ensure-Field "Client Tenancies" "OnboardingProject" @{
    Type        = "Text"
    DisplayName = "Onboarding Project ID"
}

Ensure-Field "Client Tenancies" "TechnicalOwner" @{
    Type        = "User"
    DisplayName = "Technical Owner (ZaheZone)"
}

Ensure-Field "Client Tenancies" "LastReviewed" @{
    Type        = "DateTime"
    DisplayName = "Last Reviewed"
}

# ==========================================================
# LIST 2 — Tenant Inventory (Snapshots)
# ==========================================================
Ensure-List `
    -Title "Tenant Inventory" `
    -Description "Point-in-time inventory snapshots of Microsoft 365 tenant configuration and security posture."

Ensure-LookupField `
    -List "Tenant Inventory" `
    -InternalName "ClientTenancy" `
    -DisplayName "Client Tenancy" `
    -LookupList "Client Tenancies" `
    -Required $true

Ensure-Field "Tenant Inventory" "InventoryDate" @{
    Type        = "DateTime"
    DisplayName = "Inventory Date"
    Required    = $true
}

Ensure-Field "Tenant Inventory" "SecurityDefaultsEnabled" @{
    Type        = "Boolean"
    DisplayName = "Security Defaults Enabled"
}

Ensure-Field "Tenant Inventory" "SSPREnabled" @{
    Type        = "Boolean"
    DisplayName = "SSPR Enabled"
}

Ensure-Field "Tenant Inventory" "DKIMStatus" @{
    Type        = "Choice"
    DisplayName = "DKIM Status"
    Choices     = @("Enabled","Partial","Disabled","Unknown")
}

Ensure-Field "Tenant Inventory" "PrivilegedAccountCount" @{
    Type        = "Number"
    DisplayName = "Privileged Account Count"
}

Ensure-Field "Tenant Inventory" "LegacyPartnerAccessDetected" @{
    Type        = "Boolean"
    DisplayName = "Legacy Partner Access Detected"
}

Ensure-Field "Tenant Inventory" "ConditionalAccessPresent" @{
    Type        = "Boolean"
    DisplayName = "Conditional Access Present"
}

Ensure-Field "Tenant Inventory" "InventorySummary" @{
    Type        = "Note"
    DisplayName = "Inventory Summary"
    RichText    = $false
}

# ==========================================================
# LIST 3 — Recommendations Register
# ==========================================================
Ensure-List `
    -Title "Recommendations Register" `
    -Description "Security, configuration, and governance recommendations derived from onboarding and audits."

Ensure-LookupField `
    -List "Recommendations Register" `
    -InternalName "ClientTenancy" `
    -DisplayName "Client Tenancy" `
    -LookupList "Client Tenancies" `
    -Required $true

Ensure-LookupField `
    -List "Recommendations Register" `
    -InternalName "InventoryReference" `
    -DisplayName "Inventory Reference" `
    -LookupList "Tenant Inventory"

Ensure-Field "Recommendations Register" "Category" @{
    Type        = "Choice"
    DisplayName = "Category"
    Choices     = @("Identity","Email","Security","Licensing","Governance")
}

Ensure-Field "Recommendations Register" "Severity" @{
    Type        = "Choice"
    DisplayName = "Severity"
    Choices     = @("Low","Medium","High")
    Required    = $true
}

Ensure-Field "Recommendations Register" "RiskType" @{
    Type        = "Choice"
    DisplayName = "Risk Type"
    Choices     = @("Security","Financial","Operational")
}

Ensure-Field "Recommendations Register" "AutoApplyEligible" @{
    Type        = "Boolean"
    DisplayName = "Auto-Apply Eligible"
}

Ensure-Field "Recommendations Register" "RecommendationStatus" @{
    Type        = "Choice"
    DisplayName = "Recommendation Status"
    Choices     = @("Applied","Accepted","Deferred")
    Required    = $true
}

Ensure-Field "Recommendations Register" "DecisionDate" @{
    Type        = "DateTime"
    DisplayName = "Decision Date"
}

Ensure-Field "Recommendations Register" "ApprovedBy" @{
    Type        = "User"
    DisplayName = "Approved By"
}

Ensure-Field "Recommendations Register" "LicensingDependency" @{
    Type        = "Choice"
    DisplayName = "Licensing Dependency"
    Choices     = @("None","Business Premium","E3","E5","Other")
}

Ensure-Field "Recommendations Register" "ServiceTier" @{
    Type        = "Choice"
    DisplayName = "Service Tier"
    Choices     = @("Included","Billable")
}

Ensure-Field "Recommendations Register" "Source" @{
    Type        = "Choice"
    DisplayName = "Source"
    Choices     = @("Onboarding","Audit","Secure Score","Manual")
}

Ensure-Field "Recommendations Register" "NextReviewDate" @{
    Type        = "DateTime"
    DisplayName = "Next Review Date"
}

Write-Host "✅ Client Tenancy Register build completed successfully."