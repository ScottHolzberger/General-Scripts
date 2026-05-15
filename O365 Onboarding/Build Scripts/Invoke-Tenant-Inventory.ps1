# ==========================================================
# Invoke Tenant Inventory Collection
# ==========================================================

# -------------------------------
# CONFIG
# -------------------------------
[string]$SiteUrl  = "https://zahe.sharepoint.com/sites/ZaheZoneClientTenancyRegister"
[string]$Tenant   = "zahe.onmicrosoft.com"
[string]$ClientId = "01e1b71f-cbcb-48df-a076-871aa4ba10d9"
[string]$PfxPath  = "$PSScriptRoot\ZaheZone-PnP-Projects.pfx"
[string]$PfxPasswordPlain = "UseA-LongRandomPasswordHere"

# Target Customer (IMPORTANT)
[string]$ClientName = "ZaheZone"   # Must match Client Tenancies Title

# -------------------------------
# CONNECT SharePoint
# -------------------------------
Connect-PnPOnline `
    -Url $SiteUrl `
    -Tenant $Tenant `
    -ClientId $ClientId `
    -CertificatePath $PfxPath `
    -CertificatePassword (ConvertTo-SecureString $PfxPasswordPlain -AsPlainText -Force)

Write-Host "✅ Connected to SharePoint"

# -------------------------------
# CONNECT Microsoft Graph
# -------------------------------
try {
    Connect-MgGraph `
        -TenantId $Tenant `
        -ClientId $ClientId `
        -CertificatePath $PfxPath `
        -CertificatePassword (ConvertTo-SecureString $PfxPasswordPlain -AsPlainText -Force) `
        -NoWelcome

    Write-Host "✅ Connected to Microsoft Graph"
}
catch {
    Write-Warning "⚠️ Graph connection failed - running without Graph data"
}

# -------------------------------
# GET CLIENT TENANCY ITEM
# -------------------------------
$clientItem = Get-PnPListItem `
    -List "Client Tenancies" `
    -Query "<View><Query><Where>
                <Eq>
                    <FieldRef Name='Title'/>
                    <Value Type='Text'>$ClientName</Value>
                </Eq>
            </Where></Query></View>"

if (-not $clientItem) {
    Write-Error "❌ Client not found in Client Tenancies list"
    return
}

$clientId = $clientItem.Id

# -------------------------------
# COLLECT DATA
# -------------------------------

# Security Defaults
$securityDefaults = "Unknown"
try {
    $policy = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy
    $securityDefaults = $policy.IsEnabled
} catch {}

# SSPR (basic check)
$ssprEnabled = $false
try {
    $sspr = Get-MgPolicyAuthorizationPolicy
    $ssprEnabled = $sspr.DefaultUserRolePermissions.AllowedToUseSSPR
} catch {}

# DKIM (basic check placeholder)
$dkimStatus = "Unknown"

# -------------------------------
# WRITE TO SHAREPOINT
# -------------------------------
Add-PnPListItem `
    -List "Tenant Inventory" `
    -Values @{
        "Title"                   = "Inventory - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
        "Client_x0020_Tenancy"    = $clientId
        "InventoryDate"           = (Get-Date)
        "SecurityDefaultsEnabled" = $securityDefaults
        "SSPREnabled"             = $ssprEnabled
        "DKIMStatus"              = $dkimStatus
    }

Write-Host "✅ Inventory record created"