# ==========================================================
# Create Views — FIXED (Internal Names Corrected)
# ==========================================================

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

# SAFE FIELD NAMES
$F = @{
    ClientTenancy      = "Client_x0020_Tenancy"
    InventoryReference = "Inventory_x0020_Reference"
}

function Ensure-View {
    param(
        [string]$List,
        [string]$ViewName,
        [string[]]$Fields,
        [string]$Query = ""
    )

    $view = Get-PnPView -List $List -Identity $ViewName -ErrorAction SilentlyContinue

    if ($view) {
        Write-Host "Recreating view: $ViewName"

        Remove-PnPView `
            -List $List `
            -Identity $ViewName `
            -Force
    }

    Add-PnPView `
        -List $List `
        -Title $ViewName `
        -Fields $Fields `
        -Query $Query `
        -Paged `
        -RowLimit 50

    Write-Host "✅ Created view: $ViewName"
}


# ==========================================================
# RECOMMENDATIONS REGISTER
# ==========================================================

Ensure-View `
    -List "Recommendations Register" `
    -ViewName "High Severity - Deferred" `
    -Fields @("Title",$F.ClientTenancy,"Severity","RecommendationStatus","ServiceTier") `
    -Query "<Where>
                <And>
                    <Eq>
                        <FieldRef Name='Severity'/>
                        <Value Type='Choice'>High</Value>
                    </Eq>
                    <Eq>
                        <FieldRef Name='RecommendationStatus'/>
                        <Value Type='Choice'>Deferred</Value>
                    </Eq>
                </And>
            </Where>"

Ensure-View `
    -List "Recommendations Register" `
    -ViewName "Billable Work" `
    -Fields @("Title",$F.ClientTenancy,"Severity","ServiceTier") `
    -Query "<Where>
                <Eq>
                    <FieldRef Name='ServiceTier'/>
                    <Value Type='Choice'>Billable</Value>
                </Eq>
            </Where>"

Ensure-View `
    -List "Recommendations Register" `
    -ViewName "Auto Apply Candidates" `
    -Fields @("Title",$F.ClientTenancy,"Severity","AutoApplyEligible") `
    -Query "<Where>
                <Eq>
                    <FieldRef Name='AutoApplyEligible'/>
                    <Value Type='Boolean'>1</Value>
                </Eq>
            </Where>"

# ==========================================================
# TENANT INVENTORY
# ==========================================================

Ensure-View `
    -List "Tenant Inventory" `
    -ViewName "Security Gaps" `
    -Fields @("Title",$F.ClientTenancy,"SecurityDefaultsEnabled","SSPREnabled","DKIMStatus") `
    -Query "<Where>
                <Or>
                    <Eq>
                        <FieldRef Name='SecurityDefaultsEnabled'/>
                        <Value Type='Boolean'>0</Value>
                    </Eq>
                    <Eq>
                        <FieldRef Name='SSPREnabled'/>
                        <Value Type='Boolean'>0</Value>
                    </Eq>
                </Or>
            </Where>"

# ==========================================================
# CLIENT TENANCIES
# ==========================================================

Ensure-View `
    -List "Client Tenancies" `
    -ViewName "Active Tenancies" `
    -Fields @("Title","PrimaryDomain","O365Status","LicensingTier") `
    -Query "<Where>
                <Eq>
                    <FieldRef Name='O365Status'/>
                    <Value Type='Choice'>BAU</Value>
                </Eq>
            </Where>"

Write-Host "✅ Views fixed and updated."