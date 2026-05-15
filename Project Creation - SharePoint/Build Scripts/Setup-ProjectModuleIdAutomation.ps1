Import-Module PnP.PowerShell -Force
Disconnect-PnPOnline -ErrorAction SilentlyContinue

# ===== AUTH (APP-ONLY CERT) =====
$SiteUrl  = "https://zahe.sharepoint.com/sites/Projects"
$Tenant   = "zahe.onmicrosoft.com"
$ClientId = "01e1b71f-cbcb-48df-a076-871aa4ba10d9"

$PfxPath = ".\ZaheZone-PnP-Projects.pfx"
$PfxPwd  = ConvertTo-SecureString "UseA-LongRandomPasswordHere" -AsPlainText -Force

Connect-PnPOnline -Url $SiteUrl -Tenant $Tenant -ClientId $ClientId -CertificatePath $PfxPath -CertificatePassword $PfxPwd

# ===== LISTS =====
$ProjectsList = "Project Register"
$ModulesList  = "Project Modules"

function Ensure-TextField {
  param([string]$ListName,[string]$InternalName,[string]$DisplayName)
  $f = Get-PnPField -List $ListName -Identity $InternalName -ErrorAction SilentlyContinue
  if (-not $f) {
    Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type Text | Out-Null
  }
}

function Ensure-NumberField {
  param([string]$ListName,[string]$InternalName,[string]$DisplayName)
  $f = Get-PnPField -List $ListName -Identity $InternalName -ErrorAction SilentlyContinue
  if (-not $f) {
    Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type Number | Out-Null
  }
}

# 1) Ensure ProjectID exists as Text (stored value)
Ensure-TextField -ListName $ProjectsList -InternalName "ProjectID" -DisplayName "Project ID"

# 2) Ensure Module fields exist
Ensure-TextField   -ListName $ModulesList -InternalName "ModuleID"       -DisplayName "Module ID"
Ensure-NumberField -ListName $ModulesList -InternalName "ModuleSequence" -DisplayName "Module Sequence"

# 3) Make ParentProject required (scripted)
# Uses same concept as standard PowerShell required-field updates (Field.Required = $true) [6](https://www.sharepointdiary.com/2018/01/make-field-required-in-sharepoint-online-using-powershell.html)
$parentField = Get-PnPField -List $ModulesList -Identity "ParentProject" -ErrorAction Stop
$parentField.Required = $true
$parentField.Update()
Invoke-PnPQuery

Write-Host "✅ Setup complete: fields ensured + ParentProject set to required."
