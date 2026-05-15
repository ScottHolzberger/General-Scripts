Import-Module PnP.PowerShell
Disconnect-PnPOnline -ErrorAction SilentlyContinue

# ===============================
# AUTH CONFIG (APP-ONLY)
# ===============================
$SiteUrl  = "https://zahe.sharepoint.com/sites/Projects"
$Tenant   = "zahe.onmicrosoft.com"
$ClientId = "01e1b71f-cbcb-48df-a076-871aa4ba10d9"

$PfxPath = ".\ZaheZone-PnP-Projects.pfx"
$PfxPwd  = ConvertTo-SecureString -String "UseA-LongRandomPasswordHere" -AsPlainText -Force

Connect-PnPOnline `
    -Url $SiteUrl `
    -Tenant $Tenant `
    -ClientId $ClientId `
    -CertificatePath $PfxPath `
    -CertificatePassword $PfxPwd

# ===============================
# PROJECT REGISTER
# ===============================
$ListName = "Project Register"

$list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
if (-not $list) {
    Write-Host "Creating list: $ListName"
    $list = New-PnPList -Title $ListName -Template GenericList -OnQuickLaunch
} else {
    Write-Host "List already exists: $ListName"
}

# Helper: create field only if missing
function Ensure-Field {
    param (
        [string]$InternalName,
        [scriptblock]$CreateBlock
    )
    $field = Get-PnPField -List $ListName -Identity $InternalName -ErrorAction SilentlyContinue
    if (-not $field) {
        Write-Host "Creating field: $InternalName"
        & $CreateBlock
    }
}

Ensure-Field "ProjectID" {
    Add-PnPField -List $ListName -DisplayName "Project ID" -InternalName "ProjectID" -Type Text -AddToDefaultView
}

Ensure-Field "ProjectName" {
    Add-PnPField -List $ListName -DisplayName "Project Name" -InternalName "ProjectName" -Type Text -AddToDefaultView
}

Ensure-Field "ProjectType" {
    Add-PnPField -List $ListName -DisplayName "Project Type" -InternalName "ProjectType" `
        -Type Choice -Choices @("Internal","Customer") -AddToDefaultView
}

Ensure-Field "CustomerName" {
    Add-PnPField -List $ListName -DisplayName "Customer Name" -InternalName "CustomerName" -Type Text
}

Ensure-Field "Status" {
    Add-PnPField -List $ListName -DisplayName "Status" -InternalName "Status" `
        -Type Choice -Choices @("Discovery","Active","On Hold","Closed") -AddToDefaultView
}

Ensure-Field "Owner" {
    Add-PnPField -List $ListName -DisplayName "Owner" -InternalName "Owner" -Type User -AddToDefaultView
}

Ensure-Field "PrimaryDomains" {
    Add-PnPField -List $ListName -DisplayName "Primary Domains" -InternalName "PrimaryDomains" `
        -Type MultiChoice `
        -Choices @("Microsoft 365","Telephony","CRM","Infrastructure","Security","Networking","Automation")
}

Ensure-Field "ProjectHub" {
    Add-PnPField -List $ListName -DisplayName "Project Hub" -InternalName "ProjectHub" -Type URL
}

Ensure-Field "Notes" {
    Add-PnPField -List $ListName -DisplayName "Notes" -InternalName "Notes" -Type Note
}

Write-Host "✅ Project Register deployment complete"