# ---------- CONFIG ----------
$SiteUrl = "https://zahe.sharepoint.com/sites/ZZ-Control"

# Use whichever auth you prefer:
# Connect-PnPOnline -Url $SiteUrl -Interactive
# OR your app/cert method:
# Connect-PnPOnline -Url $SiteUrl -ClientId "<CLIENT_ID>" -Tenant "<TENANT_ID>" -CertificatePath ".\cert.pfx" -CertificatePassword (ConvertTo-SecureString -AsPlainText "<PASSWORD>" -Force)

# ---------- HELPERS ----------
function Ensure-LookupFieldFromXml {
    param(
        [Parameter(Mandatory)] [string] $TargetList,
        [Parameter(Mandatory)] [string] $DisplayName,
        [Parameter(Mandatory)] [string] $InternalName,
        [Parameter(Mandatory)] [Guid]   $LookupListId,
        [Parameter(Mandatory)] [string] $LookupFieldInternalName
    )

    $existing = Get-PnPField -List $TargetList | Where-Object { $_.InternalName -eq $InternalName -or $_.Title -eq $DisplayName }
    if ($existing) {
        Write-Host "Lookup field exists: $TargetList → $DisplayName"
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

    Write-Host "Creating lookup field: $TargetList → $DisplayName"
    Add-PnPFieldFromXml -List $TargetList -FieldXml $xml | Out-Null
}

# ---------- GET LIST IDS ----------
$clientRegistry     = Get-PnPList -Identity "Client Registry"
$serviceCatalogue   = Get-PnPList -Identity "Service Catalogue"

# ---------- ENSURE SOURCE FIELDS EXIST (sanity) ----------
# These should already exist from your script runs:
# Client Registry: ClientName
# Service Catalogue: ServiceName

# ---------- CREATE LOOKUPS ----------
# Client Service Mapping:
Ensure-LookupFieldFromXml -TargetList "Client Service Mapping" -DisplayName "Client" -InternalName "Client" `
    -LookupListId $clientRegistry.Id -LookupFieldInternalName "ClientName"

Ensure-LookupFieldFromXml -TargetList "Client Service Mapping" -DisplayName "Service" -InternalName "Service" `
    -LookupListId $serviceCatalogue.Id -LookupFieldInternalName "ServiceName"

# Client Contacts:
Ensure-LookupFieldFromXml -TargetList "Client Contacts" -DisplayName "Client" -InternalName "Client" `
    -LookupListId $clientRegistry.Id -LookupFieldInternalName "ClientName"

# Client Infrastructure:
Ensure-LookupFieldFromXml -TargetList "Client Infrastructure" -DisplayName "Client" -InternalName "Client" `
    -LookupListId $clientRegistry.Id -LookupFieldInternalName "ClientName"

# Client Software:
Ensure-LookupFieldFromXml -TargetList "Client Software" -DisplayName "Client" -InternalName "Client" `
    -LookupListId $clientRegistry.Id -LookupFieldInternalName "ClientName"

Write-Host "✅ Lookup repair complete."