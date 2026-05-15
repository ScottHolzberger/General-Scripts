Import-Module PnP.PowerShell
Disconnect-PnPOnline -ErrorAction SilentlyContinue

# ===============================
# AUTH (APP-ONLY – ALREADY WORKING)
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
# PROJECT MODULES LIST
# ===============================
$ListName = "Project Modules"
$ParentList = "Project Register"

$list = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
if (-not $list) {
    Write-Host "Creating list: $ListName"
    $list = New-PnPList -Title $ListName -Template GenericList -OnQuickLaunch
} else {
    Write-Host "List already exists: $ListName"
}

# Helper function
function Ensure-Field {
    param (
        [string]$InternalName,
        [scriptblock]$CreateScript
    )
    $field = Get-PnPField -List $ListName -Identity $InternalName -ErrorAction SilentlyContinue
    if (-not $field) {
        Write-Host "Creating field: $InternalName"
        & $CreateScript
    }
}

Ensure-Field "ModuleID" {
    Add-PnPField -List $ListName -DisplayName "Module ID" -InternalName "ModuleID" -Type Text -AddToDefaultView
}

Ensure-Field "ModuleName" {
    Add-PnPField -List $ListName -DisplayName "Module Name" -InternalName "ModuleName" -Type Text -AddToDefaultView
}

Ensure-Field "ParentProject" {

    $parentList = Get-PnPList -Identity "Project Register"
    $parentListId = $parentList.Id

    $fieldXml = @"
<Field
    Type="Lookup"
    DisplayName="Parent Project"
    Required="FALSE"
    List="$parentListId"
    ShowField="ProjectID"
    EnforceUniqueValues="FALSE"
    Indexed="TRUE"
    StaticName="ParentProject"
    Name="ParentProject" />
"@

    Add-PnPFieldFromXml `
        -List $ListName `
        -FieldXml $fieldXml

    # Add field to default view
    $defaultView = Get-PnPView -List $ListName | Where-Object { $_.DefaultView -eq $true }
    if ($defaultView -and ($defaultView.ViewFields -notcontains "ParentProject")) {
        Add-PnPViewField -List $ListName -Identity $defaultView -Fields "ParentProject"
    }
}


Ensure-Field "Status" {
    Add-PnPField -List $ListName -DisplayName "Status" -InternalName "Status" `
        -Type Choice `
        -Choices @("Discovery","Active","On Hold","Closed") `
        -AddToDefaultView
}

Ensure-Field "Owner" {
    Add-PnPField -List $ListName -DisplayName "Owner" -InternalName "Owner" -Type User -AddToDefaultView
}

Ensure-Field "PrimaryDomains" {
    Add-PnPField -List $ListName -DisplayName "Primary Domains" -InternalName "PrimaryDomains" `
        -Type MultiChoice `
        -Choices @(
            "Microsoft 365",
            "Telephony",
            "CRM",
            "Infrastructure",
            "Security",
            "Networking",
            "Automation"
        )
}

Ensure-Field "ModuleHub" {
    Add-PnPField -List $ListName -DisplayName "Module Hub" -InternalName "ModuleHub" -Type URL
}

Ensure-Field "Notes" {
    Add-PnPField -List $ListName -DisplayName "Notes" -InternalName "Notes" -Type Note
}

Write-Host "✅ Project Modules deployment complete"
