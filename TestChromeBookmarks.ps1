param (
    [string]$BookmarkFolder = "Company Bookmarks",

    [string]$Bookmark1Name = "Google",
    [string]$Bookmark1URL = "https://google.com.au",

    [string]$Bookmark2Name = "",
    [string]$Bookmark2URL = "",

    [string]$Bookmark3Name = "",
    [string]$Bookmark3URL = "",

    [string]$Bookmark4Name = "",
    [string]$Bookmark4URL = "",

    [string]$Bookmark5Name = "",
    [string]$Bookmark5URL = ""
)

# Registry path for Chrome policies
$registryPath = 'HKLM:\Software\Policies\Google\Chrome'

# Ensure registry path exists
if (-not (Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force | Out-Null
}

# Enable bookmarks bar
Set-ItemProperty -Path $registryPath -Name "BookmarkBarEnabled" -Value 1 -Force

# Build bookmarks array
$bookmarks = @()
$bookmarks += @{ "toplevel_name" = "$BookmarkFolder" }

if ($Bookmark1Name -and $Bookmark1URL) {
    $bookmarks += @{ name = $Bookmark1Name; url = $Bookmark1URL }
}
if ($Bookmark2Name -and $Bookmark2URL) {
    $bookmarks += @{ name = $Bookmark2Name; url = $Bookmark2URL }
}
if ($Bookmark3Name -and $Bookmark3URL) {
    $bookmarks += @{ name = $Bookmark3Name; url = $Bookmark3URL }
}
if ($Bookmark4Name -and $Bookmark4URL) {
    $bookmarks += @{ name = $Bookmark4Name; url = $Bookmark4URL }
}
if ($Bookmark5Name -and $Bookmark5URL) {
    $bookmarks += @{ name = $Bookmark5Name; url = $Bookmark5URL }
}

# Convert to JSON
$bookmarkJson = $bookmarks | ConvertTo-Json -Compress

# Apply to registry
Set-ItemProperty -Path $registryPath -Name "ManagedBookmarks" -Value $bookmarkJson -Force


