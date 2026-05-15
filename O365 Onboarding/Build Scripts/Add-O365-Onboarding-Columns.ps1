# ==========================================================
# Add Missing Columns – O365 Onboarding Lists
# ==========================================================

# AUTH
[string]$SiteUrl  = "https://zahe.sharepoint.com/sites/ZaheZoneClientTenancyRegister"
[string]$Tenant   = "zahe.onmicrosoft.com"
[string]$ClientId = "01e1b71f-cbcb-48df-a076-871aa4ba10d9"
[string]$PfxPath  = "$PSScriptRoot\ZaheZone-PnP-Projects.pfx"
[string]$PfxPasswordPlain = "UseA-LongRandomPasswordHere"

Connect-PnPOnline `
    -Url $SiteUrl `
    -Tenant $Tenant `
    -ClientId $ClientId `
    -CertificatePath $PfxPath `
    -CertificatePassword (ConvertTo-SecureString $PfxPasswordPlain -AsPlainText -Force)

Write-Host "✅ Connected"

# -------------------------------
# Helper: Ensure Field
# -------------------------------
function Ensure-Field {
    param(
        [string]$List,
        [string]$InternalName,
        [string]$DisplayName,
        [string]$Type,
        [string[]]$Choices = $null,
        [bool]$Required = $false
    )

    $field = Get-PnPField -List $List -Identity $InternalName -ErrorAction SilentlyContinue
    if ($field) {
        Write-Host "Field exists: $DisplayName"
        return
    }

    if ($Type -eq "Choice") {
        Add-PnPField `
            -List $List `
            -DisplayName $DisplayName `
            -InternalName $InternalName `
            -Type Choice `
            -Choices $Choices `
            -Required:$Required `
            -ErrorAction Stop
    }
    elseif ($Type -eq "Boolean") {
        Add-PnPField `
            -List $List `
            -DisplayName $DisplayName `
            -InternalName $InternalName `
            -Type Boolean `
            -ErrorAction Stop
    }
    elseif ($Type -eq "DateTime") {
        Add-PnPField `
            -List $List `
            -DisplayName $DisplayName `
            -InternalName $InternalName `
            -Type DateTime `
            -ErrorAction Stop
    }

    Write-Host "✅ Created field: $DisplayName"
}

# ==========================================================
# RECOMMENDATIONS REGISTER
# ==========================================================

Ensure-Field "Recommendations Register" "RecommendationStatus" "Recommendation Status" "Choice" @("Applied","Accepted","Deferred") $true
Ensure-Field "Recommendations Register" "ServiceTier" "Service Tier" "Choice" @("Included","Billable")
Ensure-Field "Recommendations Register" "AutoApplyEligible" "Auto-Apply Eligible" "Boolean"
Ensure-Field "Recommendations Register" "Severity" "Severity" "Choice" @("Low","Medium","High") $true

# ==========================================================
# TENANT INVENTORY
# ==========================================================

Ensure-Field "Tenant Inventory" "SecurityDefaultsEnabled" "Security Defaults Enabled" "Boolean"
Ensure-Field "Tenant Inventory" "SSPREnabled" "SSPR Enabled" "Boolean"
Ensure-Field "Tenant Inventory" "DKIMStatus" "DKIM Status" "Choice" @("Enabled","Partial","Disabled","Unknown")

Write-Host "✅ All required columns created."