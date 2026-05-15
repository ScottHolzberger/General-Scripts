param(
    [Parameter(Mandatory=$true)]
    [string]$PageUrl,

    [Parameter(Mandatory=$true)]
    [string]$Tenant,

    [Parameter(Mandatory=$true)]
    [string]$ClientId,

    [Parameter(Mandatory=$true)]
    [string]$Thumbprint
)

# =========================
# Derive Site URL + Folder
# =========================
$uri = [uri]$PageUrl
$parts = $uri.AbsolutePath.Trim("/").Split("/")

$SiteUrl = "$($uri.Scheme)://$($uri.Host)/$($parts[0])/$($parts[1])"
$PageFolder = "/" + ($parts[0..($parts.Length-2)] -join "/")

Write-Host "Site URL: $SiteUrl"
Write-Host "Folder: $PageFolder"

# =========================
# Connect (Certificate Auth)
# =========================
Connect-PnPOnline -Url $SiteUrl -Tenant $Tenant -ClientId $ClientId -Thumbprint $Thumbprint

# =========================
# Module Content Map
# =========================
$modules = @{
    "PRJ-001-001" = "<h2>Module Purpose</h2><p>Provision core infrastructure.</p>"
    "PRJ-001-002" = "<h2>Module Purpose</h2><p>API framework and routing.</p>"
    "PRJ-001-003" = "<h2>Module Purpose</h2><p>Database platform.</p>"
    "PRJ-001-004" = "<h2>Module Purpose</h2><p>Customer portals.</p>"
    "PRJ-001-005" = "<h2>Module Purpose</h2><p>Monitoring and logging.</p>"
    "PRJ-001-006" = "<h2>Module Purpose</h2><p>Deployment automation.</p>"
}

# =========================
# Process Each Module
# =========================
foreach ($moduleId in $modules.Keys) {

    $pagePath = "$PageFolder/$moduleId.aspx"
    Write-Host "Updating $moduleId"

    $components = @(Get-PnPPageComponent -Page $pagePath)

    $textComponent = $components | Where-Object {
        $_.Type -eq "Text" -and $_.Text -match "ZZ-AUTO-SCAFFOLD-MODULE"
    } | Select-Object -First 1

    if (-not $textComponent) {
        Write-Warning "No scaffold found for $moduleId"
        continue
    }

    # ✅ CLEAN HTML STRUCTURE (IMPORTANT)
    $html = @"
ZZ-AUTO-SCAFFOLD-MODULE

<h2>Module Overview</h2>

$($modules[$moduleId])

<h2>Scope</h2>
<ul>
<li>Defined per module</li>
</ul>

<h2>Current Focus</h2>
<ul>
<li>Initial deployment</li>
</ul>

<h2>Authoritative Artefacts</h2>
<ul>
<li>Runbook: $moduleId</li>
</ul>
"@

    Set-PnPPageTextPart -Page $pagePath -InstanceId $textComponent.InstanceId -Text $html

    Write-Host "$moduleId updated"
}

Write-Host "✅ All modules updated successfully"