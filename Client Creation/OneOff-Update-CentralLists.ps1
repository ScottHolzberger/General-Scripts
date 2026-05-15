# ==========================================================
# One-off central list update/create
# Target: https://zahe.sharepoint.com/sites/ZZ-Control
# ==========================================================

$Tenant           = "zahe.onmicrosoft.com"
$ClientId         = "01e1b71f-cbcb-48df-a076-871aa4ba10d9"
$PfxPath          = ".\ZaheZone-PnP-Projects.pfx"
$PfxPasswordPlain = "UseA-LongRandomPasswordHere"
$PfxPassword      = ConvertTo-SecureString $PfxPasswordPlain -AsPlainText -Force
$ControlSiteUrl   = "https://zahe.sharepoint.com/sites/ZZ-Control"

function Connect-ZZ {
    param([string]$Url)

    Connect-PnPOnline `
        -Url $Url `
        -ClientId $ClientId `
        -Tenant $Tenant `
        -CertificatePath $PfxPath `
        -CertificatePassword $PfxPassword
}

function Ensure-List {
    param(
        [string]$ListName,
        [string]$Template = "GenericList"
    )

    try {
        Get-PnPList -Identity $ListName -ErrorAction Stop | Out-Null
        Write-Host "✅ List exists: $ListName"
    }
    catch {
        Write-Host "Creating list: $ListName"
        New-PnPList -Title $ListName -Template $Template | Out-Null
    }
}

function Ensure-Field {
    param(
        [string]$ListName,
        [string]$InternalName,
        [string]$DisplayName,
        [string]$Type,
        [string[]]$Choices = @()
    )

    $existing = @(Get-PnPField -List $ListName | Where-Object {
        $_.InternalName -eq $InternalName -or $_.Title -eq $DisplayName
    })

    if ($existing.Count -gt 0) {
        Write-Host "✅ Field exists: $ListName -> $DisplayName"
        return
    }

    Write-Host "Creating field: $ListName -> $DisplayName"

    switch ($Type) {
        "Text" {
            Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type Text | Out-Null
        }
        "Note" {
            Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type Note | Out-Null
        }
        "Boolean" {
            Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type Boolean | Out-Null
        }
        "Choice" {
            Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type Choice -Choices $Choices | Out-Null
        }
        "URL" {
            Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type URL | Out-Null
        }
        "DateTime" {
            Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type DateTime | Out-Null
        }
        default {
            throw "Unsupported field type: $Type"
        }
    }
}

function Ensure-LookupFieldFromXml {
    param(
        [string]$TargetList,
        [string]$DisplayName,
        [string]$InternalName,
        [Guid]$LookupListId,
        [string]$LookupFieldInternalName
    )

    $existing = @(Get-PnPField -List $TargetList | Where-Object {
        $_.InternalName -eq $InternalName -or $_.Title -eq $DisplayName
    })

    if ($existing.Count -gt 0) {
        Write-Host "✅ Lookup field exists: $TargetList -> $DisplayName"
        return
    }

    $xml = @"
<Field Type="Lookup"
       DisplayName="$DisplayName"
       Name="$InternalName"
       StaticName="$InternalName"
       Group="Custom Columns"
       Required="FALSE"
       List="{$LookupListId}"
       ShowField="$LookupFieldInternalName" />
"@

    Write-Host "Creating lookup field: $TargetList -> $DisplayName"
    Add-PnPFieldFromXml -List $TargetList -FieldXml $xml | Out-Null
}

Connect-ZZ $ControlSiteUrl

# ----------------------------------------------------------
# Ensure Client Registry metadata fields
# ----------------------------------------------------------
Ensure-Field -ListName "Client Registry" -InternalName "SiteURL" -DisplayName "SiteURL" -Type "URL"
Ensure-Field -ListName "Client Registry" -InternalName "LastSynced" -DisplayName "LastSynced" -Type "DateTime"

$clientRegistry = Get-PnPList -Identity "Client Registry"
$clientRegistryId = $clientRegistry.Id

# ----------------------------------------------------------
# Client Network (multi-record per client)
# ----------------------------------------------------------
Ensure-List -ListName "Client Network"

Ensure-LookupFieldFromXml -TargetList "Client Network" -DisplayName "Client" -InternalName "Client" `
    -LookupListId $clientRegistryId -LookupFieldInternalName "ClientName"

Ensure-Field -ListName "Client Network" -InternalName "RouterIP"     -DisplayName "Router IP"      -Type "Text"
Ensure-Field -ListName "Client Network" -InternalName "NetworkRange" -DisplayName "Network Range"  -Type "Text"
Ensure-Field -ListName "Client Network" -InternalName "SubnetMask"   -DisplayName "Subnet Mask"    -Type "Text"
Ensure-Field -ListName "Client Network" -InternalName "DNSServers"   -DisplayName "DNS Servers"    -Type "Note"
Ensure-Field -ListName "Client Network" -InternalName "DHCPServer"   -DisplayName "DHCP Server"    -Type "Text"
Ensure-Field -ListName "Client Network" -InternalName "NetworkType"  -DisplayName "Network Type"   -Type "Choice" -Choices @("Data","Voice","Guest","Management","Other")
Ensure-Field -ListName "Client Network" -InternalName "Location"     -DisplayName "Location"       -Type "Text"
Ensure-Field -ListName "Client Network" -InternalName "Notes"        -DisplayName "Notes"          -Type "Note"

# ----------------------------------------------------------
# Client Software (update existing list)
# Title = Software Name (keep default Title field)
# ----------------------------------------------------------
Ensure-List -ListName "Client Software"

Ensure-LookupFieldFromXml -TargetList "Client Software" -DisplayName "Client" -InternalName "Client" `
    -LookupListId $clientRegistryId -LookupFieldInternalName "ClientName"

Ensure-Field -ListName "Client Software" -InternalName "Vendor"        -DisplayName "Software Vendor"      -Type "Text"
Ensure-Field -ListName "Client Software" -InternalName "SupportedByZZ" -DisplayName "Support by ZZ"        -Type "Boolean"
Ensure-Field -ListName "Client Software" -InternalName "VendorPhone"   -DisplayName "Vendor Phone"         -Type "Text"
Ensure-Field -ListName "Client Software" -InternalName "VendorEmail"   -DisplayName "Vendor Email"         -Type "Text"
Ensure-Field -ListName "Client Software" -InternalName "Notes"         -DisplayName "Software Notes"       -Type "Note"

Write-Host "✅ Central list update/create complete." -ForegroundColor Green