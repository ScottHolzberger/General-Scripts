$ErrorActionPreference = "Stop"

# ===== CONFIG =====
$TenantId = "zahe.onmicrosoft.com"
$ClientId = "01e1b71f-cbcb-48df-a076-871aa4ba10d9"
$PfxPath  = Join-Path $PSScriptRoot "ZaheZone-PnP-Projects.pfx"
$PfxPasswordPlain = "UseA-LongRandomPasswordHere"

$SiteUrl = "https://zahe.sharepoint.com/sites/ZaheZoneClientTenancyRegister"
$Tenant  = "zahe.onmicrosoft.com"

# Must match Client Tenancies -> Title exactly
$ClientTenancyTitle = "ZaheZone"

# ===== Script paths =====
$GraphScript = Join-Path $PSScriptRoot "Collect-Tenant-Inventory-Graph.ps1"
$PnPScript   = Join-Path $PSScriptRoot "Write-Tenant-Inventory-SharePoint.ps1"

if (-not (Test-Path $GraphScript)) { throw "Missing file: $GraphScript" }
if (-not (Test-Path $PnPScript))   { throw "Missing file: $PnPScript" }
if (-not (Test-Path $PfxPath))     { throw "Missing file: $PfxPath" }

$tmpJson = Join-Path $env:TEMP ("tenant-inventory-{0}-{1}.json" -f ($ClientTenancyTitle -replace '\W',''), (Get-Date -Format yyyyMMddHHmmss))

Write-Host "== Step 1: Graph collection ==" -ForegroundColor Cyan
& pwsh -NoProfile -File $GraphScript `
  -TenantId $TenantId `
  -ClientId $ClientId `
  -PfxPath $PfxPath `
  -PfxPasswordPlain $PfxPasswordPlain `
  -OutJson $tmpJson

if ($LASTEXITCODE -ne 0) { throw "Graph step failed (exit code $LASTEXITCODE)" }
if (-not (Test-Path $tmpJson)) { throw "Graph output missing: $tmpJson" }

Write-Host "== Step 2: SharePoint write + rule engine ==" -ForegroundColor Cyan
& pwsh -NoProfile -File $PnPScript `
  -SiteUrl $SiteUrl `
  -Tenant $Tenant `
  -ClientId $ClientId `
  -PfxPath $PfxPath `
  -PfxPasswordPlain $PfxPasswordPlain `
  -ClientTenancyTitle $ClientTenancyTitle `
  -InventoryJsonPath $tmpJson

if ($LASTEXITCODE -ne 0) { throw "PnP step failed (exit code $LASTEXITCODE)" }

Write-Host "✅ Full inventory pipeline COMPLETE" -ForegroundColor Green
Write-Host "JSON used: $tmpJson"