param(
  [Parameter(Mandatory)][string]$TenantId,
  [Parameter(Mandatory)][string]$ClientId,
  [Parameter(Mandatory)][string]$PfxPath,
  [Parameter(Mandatory)][string]$PfxPasswordPlain,
  [Parameter(Mandatory)][string]$OutJson
)

$ErrorActionPreference = "Stop"

# Minimal Graph modules (avoid bringing in unnecessary assemblies)
Import-Module Microsoft.Graph.Authentication -ErrorAction Stop
Import-Module Microsoft.Graph.Identity.SignIns -ErrorAction Stop
Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

# Import PFX into CurrentUser cert store (only if needed), then connect by thumbprint
$pwd = ConvertTo-SecureString $PfxPasswordPlain -AsPlainText -Force
$cert = $null

try {
  # Try to load cert directly from PFX (no store requirement) and import into CurrentUser\My
  $cert = Import-PfxCertificate -FilePath $PfxPath -CertStoreLocation Cert:\CurrentUser\My -Password $pwd -ErrorAction Stop
} catch {
  # If already imported previously, attempt to find a matching cert in CurrentUser\My by thumbprint from file
  $tmp = Get-PfxCertificate -FilePath $PfxPath
  $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $tmp.Thumbprint } | Select-Object -First 1
  if (-not $cert) { throw }
}

Connect-MgGraph -TenantId $TenantId -ClientId $ClientId -CertificateThumbprint $cert.Thumbprint -NoWelcome | Out-Null

# ---- Collect signals (keep it strict + expandable) ----
$org = $null
try { $org = Get-MgOrganization | Select-Object -First 1 } catch {}

$securityDefaultsEnabled = $null
try {
  $sd = Get-MgPolicyIdentitySecurityDefaultEnforcementPolicy
  $securityDefaultsEnabled = $sd.IsEnabled
} catch {
  $securityDefaultsEnabled = $null
}

$ssprEnabled = $null
try {
  # NOTE: This can return null in some tenants; treat as Unknown rather than False
  $ap = Get-MgPolicyAuthorizationPolicy
  $ssprEnabled = $ap.DefaultUserRolePermissions.AllowedToUseSSPR
} catch {
  $ssprEnabled = $null
}

$skus = @()
try {
  $skus = Get-MgSubscribedSku | ForEach-Object { $_.SkuPartNumber }
} catch {
  $skus = @()
}

$result = [pscustomobject]@{
  CollectedAt            = (Get-Date).ToString("s")
  TenantDisplayName      = $org.DisplayName
  VerifiedDomains        = @($org.VerifiedDomains.Name)
  SecurityDefaultsEnabled= $securityDefaultsEnabled
  SSPREnabled            = $ssprEnabled
  SKUs                   = $skus
}

$result | ConvertTo-Json -Depth 6 | Out-File -Encoding utf8 $OutJson
Disconnect-MgGraph | Out-Null

Write-Host "✅ Graph inventory written to $OutJson"