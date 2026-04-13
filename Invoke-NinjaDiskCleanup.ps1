<#
.SYNOPSIS
  NinjaOne-safe endpoint disk cleanup + JSON output + ZaheZone TXT log + NinjaOne Custom Fields.

.DESCRIPTION
  Cleans:
   - Temp folders (system + all users)
   - Recycle Bin (best effort)
   - Delivery Optimization cache
   - Windows Update download cache
   - Windows Error Reporting (WER) cache
   - Crash dumps (Minidump + MEMORY.DMP)
   - Windows logs (CBS/DISM logs older than X days)
   - ZaheZone logs (C:\ZaheZone\Logs files older than X days)
   - DISM component cleanup (StartComponentCleanup) - safe, no ResetBase

  Produces:
   - Detailed NinjaOne log file of every step
   - ZaheZone TXT log in C:\ZaheZone\Logs named "Disk Cleanup {Date & Time started}.log"
   - Per-step reclaimed disk space (bytes)
   - Total reclaimed space as GB via $TotalReclaimed (decimal)
   - FINAL JSON output (single-line) for NinjaOne parsing

  Sets NinjaOne Custom Fields:
   - diskCleanUpLastDate        (Date/DateTime safe - writes Unix Epoch seconds under the hood)
   - diskCleanupReclaimedSpace  (decimal GB)
   - diskCleanUpFullResults     (WYSIWYG HTML <pre> of log contents)

.NOTES
  Recommended NinjaOne settings:
    - Run as: SYSTEM
    - 64-bit
    - Timeout: 10–20 minutes (DISM can take time)

  Important:
    - Ensure the custom fields exist in NinjaOne and allow Automation Read/Write.
    - If diskCleanUpLastDate is a Date/DateTime field, Ninja expects Unix Epoch seconds.
#>

[CmdletBinding()]
param(
    [string]$LogRoot = "C:\ProgramData\NinjaOne\DiskCleanup\Logs",
    [string]$DriveLetter = $env:SystemDrive,

    [int]$DaysToKeepWindowsLogs = 30,

    # ZaheZone logs cleanup settings
    [int]$DaysToKeepZaheZoneLogs = 90,
    [bool]$IncludeZaheZoneLogCleanup = $true,

    [bool]$IncludeTempCleanup = $true,
    [bool]$IncludeRecycleBinCleanup = $true,
    [bool]$IncludeDeliveryOptimizationCleanup = $true,
    [bool]$IncludeWindowsUpdateCleanup = $true,
    [bool]$IncludeWerCleanup = $true,
    [bool]$IncludeCrashDumpCleanup = $true,
    [bool]$IncludeWindowsLogCleanup = $true,
    [bool]$IncludeDismComponentCleanup = $true,

    [switch]$ReportOnly,

    # NinjaOne JSON output controls
    [bool]$OutputJson = $true,
    [switch]$JsonOnly,

    # WYSIWYG field safety (avoid size limits)
    [int]$MaxWysiwygChars = 30000,

    # Optional: attempt read-back after setting fields (helps troubleshooting)
    [bool]$EnableNinjaReadBack = $true
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

# -----------------------------
# Folder + Log initialization
# -----------------------------
function New-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Initialize-ZZTxtLog {
    param(
        [Parameter(Mandatory)][datetime]$StartTime,
        [string]$LogFolder = "C:\ZaheZone\Logs"
    )

    New-Folder -Path $LogFolder

    # File-system safe timestamp: yyyy-MM-dd HH-mm-ss (avoid colon in filenames)
    $safeStamp = $StartTime.ToString("yyyy-MM-dd HH-mm-ss")
    $logPath   = Join-Path $LogFolder ("Disk Cleanup {0}.log" -f $safeStamp)

    New-Item -Path $logPath -ItemType File -Force | Out-Null
    return $logPath
}

# Start time captured once for naming + JSON
$StartTime = Get-Date

# Create ZaheZone TXT log file
$script:ZZTxtLogFile = Initialize-ZZTxtLog -StartTime $StartTime

# Create NinjaOne log folder + file
New-Folder -Path $LogRoot
$script:LogFile = Join-Path $LogRoot ("DiskCleanup_{0}_{1}.log" -f $env:COMPUTERNAME, $StartTime.ToString("yyyyMMdd_HHmmss"))

# -----------------------------
# Logging
# -----------------------------
function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR","STEP")][string]$Level = "INFO"
    )

    $stamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line  = "[$stamp] [$Level] $Message"

    # Always write to NinjaOne log file
    Add-Content -Path $script:LogFile -Value $line

    # Always write to ZaheZone TXT log file
    if ($script:ZZTxtLogFile) {
        Add-Content -Path $script:ZZTxtLogFile -Value $line
    }

    # Optionally write to STDOUT (Ninja)
    if (-not $JsonOnly) {
        Write-Output $line
    }
}

# -----------------------------
# Admin check
# -----------------------------
function Test-IsAdmin {
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        $p  = New-Object Security.Principal.WindowsPrincipal($id)
        return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch {
        return $false
    }
}

# -----------------------------
# Drive / formatting helpers
# -----------------------------
function Get-FreeBytes {
    param([string]$Drive = $DriveLetter)
    try {
        $dl = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$Drive'"
        return [int64]$dl.FreeSpace
    } catch {
        return -1
    }
}

function Get-TotalBytes {
    param([string]$Drive = $DriveLetter)
    try {
        $dl = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$Drive'"
        return [int64]$dl.Size
    } catch {
        return -1
    }
}

function Format-Bytes {
    param([int64]$Bytes)
    if ($Bytes -lt 0) { return "N/A" }
    if ($Bytes -ge 1TB) { return ("{0:N2} TB" -f ($Bytes / 1TB)) }
    if ($Bytes -ge 1GB) { return ("{0:N2} GB" -f ($Bytes / 1GB)) }
    if ($Bytes -ge 1MB) { return ("{0:N2} MB" -f ($Bytes / 1MB)) }
    if ($Bytes -ge 1KB) { return ("{0:N2} KB" -f ($Bytes / 1KB)) }
    return ("{0} B" -f $Bytes)
}

# -----------------------------
# Service helpers
# -----------------------------
function Stop-ServicesSafe {
    param([string[]]$Names)
    foreach ($n in $Names) {
        try {
            $svc = Get-Service -Name $n -ErrorAction Stop
            if ($svc.Status -ne "Stopped") {
                Write-Log "Stopping service: $n" "INFO"
                if (-not $ReportOnly) { Stop-Service -Name $n -Force -ErrorAction SilentlyContinue }
            }
        } catch {
            Write-Log "Service not found or cannot query: $n" "WARN"
        }
    }
}

function Start-ServicesSafe {
    param([string[]]$Names)
    foreach ($n in $Names) {
        try {
            $svc = Get-Service -Name $n -ErrorAction Stop
            if ($svc.Status -ne "Running") {
                Write-Log "Starting service: $n" "INFO"
                if (-not $ReportOnly) { Start-Service -Name $n -ErrorAction SilentlyContinue }
            }
        } catch {
            Write-Log "Service not found or cannot query: $n" "WARN"
        }
    }
}

# -----------------------------
# Cleanup helpers
# -----------------------------
function Remove-Contents {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Label
    )

    if (-not (Test-Path $Path)) {
        Write-Log "Skip (not found): $Label [$Path]" "INFO"
        return
    }

    Write-Log "Cleaning: $Label [$Path]" "INFO"

    if ($ReportOnly) {
        Write-Log "ReportOnly enabled; not deleting contents of: $Label" "WARN"
        return
    }

    try {
        Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue | ForEach-Object {
            try { Remove-Item -LiteralPath $_.FullName -Recurse -Force -ErrorAction SilentlyContinue } catch {}
        }
    } catch {
        Write-Log "Failed to clean $Label ($Path): $($_.Exception.Message)" "WARN"
    }
}

function Remove-FilesOlderThan {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][int]$Days,
        [Parameter(Mandatory)][string]$Label,
        [string[]]$IncludePatterns = @("*.log","*.cab")
    )

    if (-not (Test-Path $Path)) {
        Write-Log "Skip (not found): $Label [$Path]" "INFO"
        return
    }

    $cutoff = (Get-Date).AddDays(-1 * [Math]::Abs($Days))
    Write-Log "Cleaning: $Label [$Path] older than $Days days" "INFO"

    if ($ReportOnly) {
        Write-Log "ReportOnly enabled; not deleting: $Label" "WARN"
        return
    }

    try {
        foreach ($pat in $IncludePatterns) {
            Get-ChildItem -LiteralPath $Path -Recurse -Force -File -Filter $pat -ErrorAction SilentlyContinue |
                Where-Object { $_.LastWriteTime -lt $cutoff } |
                ForEach-Object {
                    try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue } catch {}
                }
        }
    } catch {
        Write-Log "Failed cleanup for ${Label}: $($_.Exception.Message)" "WARN"
    }
}

function Remove-FilesOlderThanEx {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][int]$Days,
        [Parameter(Mandatory)][string]$Label,
        [string[]]$IncludePatterns = @("*.*"),
        [string[]]$ExcludeFullPaths = @()
    )

    if (-not (Test-Path $Path)) {
        Write-Log "Skip (not found): $Label [$Path]" "INFO"
        return
    }

    $cutoff = (Get-Date).AddDays(-1 * [Math]::Abs($Days))
    Write-Log "Cleaning: $Label [$Path] older than $Days days" "INFO"

    if ($ReportOnly) {
        Write-Log "ReportOnly enabled; not deleting: $Label" "WARN"
        return
    }

    try {
        $targets = @()
        foreach ($pat in $IncludePatterns) {
            $targets += Get-ChildItem -LiteralPath $Path -Recurse -Force -File -Filter $pat -ErrorAction SilentlyContinue
        }

        $targets = $targets |
            Sort-Object FullName -Unique |
            Where-Object { $_.LastWriteTime -lt $cutoff }

        $excludeNormalized = @()
        foreach ($e in $ExcludeFullPaths) { if ($e) { $excludeNormalized += $e.ToLowerInvariant() } }

        if ($excludeNormalized.Count -gt 0) {
            $targets = $targets | Where-Object { $_.FullName.ToLowerInvariant() -notin $excludeNormalized }
        }

        $count = ($targets | Measure-Object).Count
        Write-Log "$Label candidates: $count file(s)" "INFO"

        foreach ($f in $targets) {
            try { Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue } catch {}
        }
    } catch {
        Write-Log "Failed cleanup for ${Label}: $($_.Exception.Message)" "WARN"
    }
}

# -----------------------------
# NinjaOne Custom Fields helpers
# -----------------------------
function Get-NinjaCliPath {
    $p = "C:\ProgramData\NinjaRMMAgent\ninjarmm-cli.exe"
    if (Test-Path $p) { return $p }
    return $null
}

function Convert-ToUnixEpochSeconds {
    param(
        [Parameter(Mandatory)]$DateValue
    )
    # Manual epoch conversion (safe on older .NET/Windows PowerShell)
    $dtUtc = (Get-Date $DateValue).ToUniversalTime()
    $epochStartUtc = [DateTime]::SpecifyKind([DateTime]"1970-01-01 00:00:00", [DateTimeKind]::Utc)
    $seconds = (New-TimeSpan -Start $epochStartUtc -End $dtUtc).TotalSeconds
    return [int64][Math]::Floor($seconds)
}

function Invoke-NinjaPropertySet {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)]$Value
    )

    $cmd = Get-Command -Name "Ninja-Property-Set" -ErrorAction SilentlyContinue
    if (-not $cmd) { return $false }

    try {
        $hasNameParam  = $cmd.Parameters.ContainsKey("Name")
        $hasValueParam = $cmd.Parameters.ContainsKey("Value")

        $out = $null
        if ($hasNameParam -and $hasValueParam) {
            $out = Ninja-Property-Set -Name $Name -Value $Value 2>&1
        } else {
            $out = Ninja-Property-Set $Name $Value 2>&1
        }

        if ($out -and $out.Exception) { throw $out }
        return $true
    } catch {
        Write-Log "Ninja-Property-Set failed for '$Name': $($_.Exception.Message)" "WARN"
        return $false
    }
}

function Invoke-NinjaPropertyGet {
    param([Parameter(Mandatory)][string]$Name)

    $cmd = Get-Command -Name "Ninja-Property-Get" -ErrorAction SilentlyContinue
    if (-not $cmd) { return $null }

    try {
        return (Ninja-Property-Get $Name 2>$null)
    } catch {
        return $null
    }
}

function Invoke-NinjaCliSet {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)]$Value
    )

    $cli = Get-NinjaCliPath
    if (-not $cli) { return $false }

    try {
        & $cli set $Name "$Value" | Out-Null
        return ($LASTEXITCODE -eq 0)
    } catch {
        Write-Log "ninjarmm-cli failed for '$Name': $($_.Exception.Message)" "WARN"
        return $false
    }
}

function Set-NinjaCustomFieldSmart {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)]$Value,
        [string]$Type
    )

    # If Date/DateTime, send epoch seconds for Ninja date field compatibility.
    $finalValue = $Value
    $didEpoch = $false
    if ($Type -and ($Type -eq "Date" -or $Type -eq "Date And Time")) {
        $finalValue = Convert-ToUnixEpochSeconds -DateValue $Value
        $didEpoch = $true
    }

    $ok = Invoke-NinjaPropertySet -Name $Name -Value $finalValue
    if (-not $ok) {
        $ok = Invoke-NinjaCliSet -Name $Name -Value $finalValue
    }

    # If we tried epoch and failed, retry as string (covers field being Text unexpectedly)
    if (-not $ok -and $didEpoch) {
        Write-Log "Retrying '$Name' as plain text value (field may not be Date/DateTime)..." "WARN"
        $ok = Invoke-NinjaPropertySet -Name $Name -Value "$Value"
        if (-not $ok) { $ok = Invoke-NinjaCliSet -Name $Name -Value "$Value" }
    }

    return $ok
}

function Convert-LogToWysiwygHtml {
    param(
        [Parameter(Mandatory)][string]$Path,
        [int]$MaxChars = 30000
    )

    if (-not (Test-Path $Path)) {
        return "<pre style='font-family:Consolas,monospace;white-space:pre-wrap;'>Log file not found: $Path</pre>"
    }

    $raw = ""
    try {
        $raw = Get-Content -LiteralPath $Path -Raw -ErrorAction Stop
    } catch {
        $raw = "Unable to read log file: $Path`r`n$($_.Exception.Message)"
    }

    if ($MaxChars -gt 0 -and $raw.Length -gt $MaxChars) {
        $startIndex = [Math]::Max(0, $raw.Length - $MaxChars)
        $length = [Math]::Min($MaxChars, $raw.Length)
        $raw = $raw.Substring($startIndex, $length)
        $raw = "[TRUNCATED: showing last $MaxChars characters]`r`n`r`n" + $raw
    }

    $encoded = [System.Net.WebUtility]::HtmlEncode($raw)
    return "<pre style='font-family:Consolas,monospace;white-space:pre-wrap;'>$encoded</pre>"
}

# -----------------------------
# Step runner (per-step reclaimed space)
# -----------------------------
$StepResults = New-Object System.Collections.Generic.List[object]

function Invoke-Step {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$Action
    )

    $before = Get-FreeBytes
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Log "---- STEP START: $Name ----" "STEP"

    $status = "OK"
    try { & $Action }
    catch {
        $status = "ERROR: $($_.Exception.Message)"
        Write-Log "Step error in [$Name]: $($_.Exception.Message)" "ERROR"
    }

    $sw.Stop()
    $after = Get-FreeBytes

    $delta = 0
    if ($before -ge 0 -and $after -ge 0) { $delta = ($after - $before) }

    $row = [pscustomobject]@{
        Step           = $Name
        BeforeBytes    = $before
        AfterBytes     = $after
        ReclaimedBytes = $delta
        DurationSec    = [Math]::Round($sw.Elapsed.TotalSeconds, 2)
        Status         = $status
    }

    $StepResults.Add($row) | Out-Null

    Write-Log ("STEP END: {0} | Before: {1} | After: {2} | Reclaimed: {3} | Duration: {4}s | Status: {5}" -f `
        $Name, (Format-Bytes $before), (Format-Bytes $after), (Format-Bytes $delta), $row.DurationSec, $status) "STEP"
}

# -----------------------------
# Start
# -----------------------------
Write-Log "==== NinjaOne Disk Cleanup started ====" "INFO"
Write-Log "Computer: $env:COMPUTERNAME | RunAs: $env:USERNAME | Drive: $DriveLetter | ReportOnly: $ReportOnly | JsonOnly: $JsonOnly" "INFO"
Write-Log "Ninja log file: $script:LogFile" "INFO"
Write-Log "ZaheZone TXT log file: $script:ZZTxtLogFile" "INFO"

$IsAdmin = Test-IsAdmin
if (-not $IsAdmin) {
    Write-Log "Not running elevated. In NinjaOne set 'Run as: SYSTEM'. Exiting." "ERROR"

    if ($OutputJson) {
        $failObj = [pscustomobject]@{
            schemaVersion = 1
            computerName  = $env:COMPUTERNAME
            drive         = $DriveLetter
            startTime     = $StartTime.ToString("o")
            endTime       = (Get-Date).ToString("o")
            success       = $false
            exitCode      = 2
            error         = "Not running elevated"
            logFile       = $script:LogFile
            zzTxtLogFile  = $script:ZZTxtLogFile
            steps         = @()
        }
        Write-Output ($failObj | ConvertTo-Json -Depth 8 -Compress)
    }

    exit 2
}

$FreeStart  = Get-FreeBytes
$TotalBytes = Get-TotalBytes
$PctStart   = if ($FreeStart -ge 0 -and $TotalBytes -gt 0) { [Math]::Round(($FreeStart / $TotalBytes) * 100, 2) } else { -1 }

Write-Log ("Drive size: {0} | Free START: {1} ({2}%)" -f (Format-Bytes $TotalBytes), (Format-Bytes $FreeStart), $PctStart) "INFO"

# -----------------------------
# Cleanup Modules
# -----------------------------
if ($IncludeTempCleanup) {
    Invoke-Step -Name "Temp Cleanup (System + All Users)" -Action {
        Remove-Contents -Path "$env:windir\Temp" -Label "Windows TEMP"
        Remove-Contents -Path "C:\Windows\Temp" -Label "C:\Windows\Temp"

        if ($env:TEMP) { Remove-Contents -Path $env:TEMP -Label "Current context TEMP" }

        $usersRoot = "C:\Users"
        if (Test-Path $usersRoot) {
            Get-ChildItem -Path $usersRoot -Directory -Force -ErrorAction SilentlyContinue |
                Where-Object { $_.Name -notin @("Public","Default","Default User","All Users") } |
                ForEach-Object {
                    $p = Join-Path $_.FullName "AppData\Local\Temp"
                    Remove-Contents -Path $p -Label "User TEMP ($($_.Name))"
                }
        }
    }
}

if ($IncludeRecycleBinCleanup) {
    Invoke-Step -Name "Recycle Bin Cleanup (Best Effort)" -Action {
        try {
            if ($ReportOnly) {
                Write-Log "ReportOnly enabled; not clearing Recycle Bin." "WARN"
            } else {
                Clear-RecycleBin -Force -ErrorAction SilentlyContinue | Out-Null
            }
        } catch {
            Write-Log "Clear-RecycleBin failed (non-fatal): $($_.Exception.Message)" "WARN"
        }

        $rb = "C:\`$Recycle.Bin"
        if (Test-Path $rb) {
            Remove-Contents -Path $rb -Label "Recycle Bin Folder (fallback)"
        }
    }
}

if ($IncludeDeliveryOptimizationCleanup) {
    Invoke-Step -Name "Delivery Optimization Cache Cleanup" -Action {
        Stop-ServicesSafe -Names @("dosvc")
        $doCache = "C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache"
        Remove-Contents -Path $doCache -Label "Delivery Optimization Cache"
        Start-ServicesSafe -Names @("dosvc")
    }
}

if ($IncludeWindowsUpdateCleanup) {
    Invoke-Step -Name "Windows Update Download Cache Cleanup" -Action {
        $svc = @("wuauserv","bits")
        Stop-ServicesSafe -Names $svc
        Remove-Contents -Path "C:\Windows\SoftwareDistribution\Download" -Label "Windows Update Download Cache"
        Start-ServicesSafe -Names $svc
    }
}

if ($IncludeWerCleanup) {
    Invoke-Step -Name "Windows Error Reporting (WER) Cache Cleanup" -Action {
        $wer1 = "C:\ProgramData\Microsoft\Windows\WER\ReportArchive"
        $wer2 = "C:\ProgramData\Microsoft\Windows\WER\ReportQueue"
        Remove-Contents -Path $wer1 -Label "WER ReportArchive"
        Remove-Contents -Path $wer2 -Label "WER ReportQueue"
    }
}

if ($IncludeCrashDumpCleanup) {
    Invoke-Step -Name "Crash Dump Cleanup" -Action {
        Remove-Contents -Path "C:\Windows\Minidump" -Label "Windows Minidumps"

        if (Test-Path "C:\Windows\MEMORY.DMP") {
            Write-Log "Deleting: C:\Windows\MEMORY.DMP" "INFO"
            if ($ReportOnly) {
                Write-Log "ReportOnly enabled; not deleting MEMORY.DMP" "WARN"
            } else {
                try { Remove-Item -LiteralPath "C:\Windows\MEMORY.DMP" -Force -ErrorAction SilentlyContinue } catch {}
            }
        } else {
            Write-Log "MEMORY.DMP not present." "INFO"
        }
    }
}

if ($IncludeWindowsLogCleanup) {
    Invoke-Step -Name "Windows Logs Cleanup (CBS/DISM older than $DaysToKeepWindowsLogs days)" -Action {
        Remove-FilesOlderThan -Path "C:\Windows\Logs\CBS"  -Days $DaysToKeepWindowsLogs -Label "CBS Logs"  -IncludePatterns @("*.log","*.cab")
        Remove-FilesOlderThan -Path "C:\Windows\Logs\DISM" -Days $DaysToKeepWindowsLogs -Label "DISM Logs" -IncludePatterns @("*.log")
    }
}

# ZaheZone Logs cleanup (older than 90 days by default; excludes current run log)
if ($IncludeZaheZoneLogCleanup) {
    Invoke-Step -Name "ZaheZone Logs Cleanup (older than $DaysToKeepZaheZoneLogs days)" -Action {
        $zzFolder = "C:\ZaheZone\Logs"
        $exclude = @($script:ZZTxtLogFile)

        Remove-FilesOlderThanEx -Path $zzFolder `
            -Days $DaysToKeepZaheZoneLogs `
            -Label "ZaheZone Logs" `
            -IncludePatterns @("*.log","*.txt") `
            -ExcludeFullPaths $exclude
    }
} else {
    Write-Log "ZaheZone log cleanup disabled." "INFO"
}

if ($IncludeDismComponentCleanup) {
    Invoke-Step -Name "DISM Component Store Cleanup (StartComponentCleanup)" -Action {
        if ($ReportOnly) {
            Write-Log "ReportOnly enabled; not running DISM." "WARN"
            return
        }
        $args = "/Online /Cleanup-Image /StartComponentCleanup /Quiet"
        Write-Log "Running: dism.exe $args" "INFO"

        try {
            $p = Start-Process -FilePath "dism.exe" -ArgumentList $args -Wait -PassThru -WindowStyle Hidden
            Write-Log "DISM exit code: $($p.ExitCode)" "INFO"
        } catch {
            Write-Log "DISM failed (non-fatal): $($_.Exception.Message)" "WARN"
        }
    }
}

# -----------------------------
# Summary (TotalReclaimed is GB decimal)
# -----------------------------
$EndTime  = Get-Date
$FreeEnd  = Get-FreeBytes
$PctEnd   = if ($FreeEnd -ge 0 -and $TotalBytes -gt 0) { [Math]::Round(($FreeEnd / $TotalBytes) * 100, 2) } else { -1 }

$TotalReclaimedBytes = if ($FreeStart -ge 0 -and $FreeEnd -ge 0) { [Math]::Max(($FreeEnd - $FreeStart), 0) } else { 0 }
$TotalReclaimed = [Math]::Round(($TotalReclaimedBytes / 1GB), 2)

Write-Log ("Free END: {0} ({1}%)" -f (Format-Bytes $FreeEnd), $PctEnd) "INFO"
Write-Log ("TOTAL reclaimed (whole run): {0}" -f (Format-Bytes $TotalReclaimedBytes)) "INFO"
Write-Log ("TOTAL reclaimed (GB): {0} GB" -f $TotalReclaimed) "INFO"
Write-Log "Ninja log file saved to: $script:LogFile" "INFO"
Write-Log "ZaheZone TXT log file saved to: $script:ZZTxtLogFile" "INFO"

# -----------------------------
# Set NinjaOne Custom Fields (Date/DateTime epoch-safe)
# -----------------------------
$customFieldResults = [ordered]@{}

try {
    # Requested format (string for logging / consistency)
    $diskCleanUpLastDateString = $EndTime.ToString("yyyy-MM-dd HH:mm:ss")

    # Reclaimed space as decimal string (invariant culture)
    $diskCleanupReclaimedSpace = $TotalReclaimed.ToString("0.00", [System.Globalization.CultureInfo]::InvariantCulture)

    # WYSIWYG HTML <pre> of log contents
    $diskCleanUpFullResults = Convert-LogToWysiwygHtml -Path $script:ZZTxtLogFile -MaxChars $MaxWysiwygChars

    # Date/DateTime field: send epoch seconds (Ninja requirement).
    $customFieldResults.diskCleanUpLastDate = Set-NinjaCustomFieldSmart -Name "diskCleanUpLastDate" -Value $diskCleanUpLastDateString -Type "Date And Time"
    $customFieldResults.diskCleanupReclaimedSpace = Set-NinjaCustomFieldSmart -Name "diskCleanupReclaimedSpace" -Value $diskCleanupReclaimedSpace -Type "Decimal"
    $customFieldResults.diskCleanUpFullResults = Set-NinjaCustomFieldSmart -Name "diskCleanUpFullResults" -Value $diskCleanUpFullResults -Type "Text"

    Write-Log ("Custom Field set: diskCleanUpLastDate = {0} (success={1})" -f $diskCleanUpLastDateString, $customFieldResults.diskCleanUpLastDate) "INFO"
    Write-Log ("Custom Field set: diskCleanupReclaimedSpace = {0} (success={1})" -f $diskCleanupReclaimedSpace, $customFieldResults.diskCleanupReclaimedSpace) "INFO"
    Write-Log ("Custom Field set: diskCleanUpFullResults = (HTML content) (success={0})" -f $customFieldResults.diskCleanUpFullResults) "INFO"

    # Optional: read-back validation
    if ($EnableNinjaReadBack) {
        $rb = Invoke-NinjaPropertyGet -Name "diskCleanUpLastDate"
        if ($null -ne $rb) {
            Write-Log "Read-back diskCleanUpLastDate: $rb" "INFO"
        } else {
            Write-Log "Read-back diskCleanUpLastDate not available (Ninja-Property-Get missing or returned null)." "WARN"
        }
    }
} catch {
    Write-Log "Custom field update block failed: $($_.Exception.Message)" "WARN"
}

Write-Log "==== NinjaOne Disk Cleanup finished ====" "INFO"

# -----------------------------
# JSON Output for NinjaOne (single line at end)
# -----------------------------
$exitCode = 0
if ($FreeStart -lt 0 -or $FreeEnd -lt 0) { $exitCode = 1 }

$stepErrors = @($StepResults | Where-Object { $_.Status -like "ERROR*" })
$success = ($exitCode -eq 0 -and $stepErrors.Count -eq 0)

if ($OutputJson) {
    $jsonObj = [pscustomobject]@{
        schemaVersion                = 1
        computerName                 = $env:COMPUTERNAME
        drive                        = $DriveLetter
        startTime                    = $StartTime.ToString("o")
        endTime                      = $EndTime.ToString("o")
        durationSec                  = [Math]::Round(($EndTime - $StartTime).TotalSeconds, 2)
        reportOnly                   = [bool]$ReportOnly
        success                      = $success
        exitCode                     = $exitCode

        logFile                      = $script:LogFile
        zzTxtLogFile                 = $script:ZZTxtLogFile

        driveSizeBytes               = $TotalBytes
        freeStartBytes               = $FreeStart
        freeEndBytes                 = $FreeEnd
        freeStartPercent             = $PctStart
        freeEndPercent               = $PctEnd

        totalReclaimedBytes          = $TotalReclaimedBytes
        totalReclaimedGB             = $TotalReclaimed
        totalReclaimedHuman          = (Format-Bytes $TotalReclaimedBytes)

        daysToKeepWindowsLogs        = $DaysToKeepWindowsLogs
        includeZaheZoneLogCleanup    = $IncludeZaheZoneLogCleanup
        daysToKeepZaheZoneLogs       = $DaysToKeepZaheZoneLogs
        maxWysiwygChars              = $MaxWysiwygChars

        customFields                 = $customFieldResults

        steps = @(
            $StepResults | ForEach-Object {
                [pscustomobject]@{
                    name           = $_.Step
                    status         = $_.Status
                    durationSec    = $_.DurationSec
                    beforeBytes    = $_.BeforeBytes
                    afterBytes     = $_.AfterBytes
                    reclaimedBytes = $_.ReclaimedBytes
                }
            }
        )

        warnings = @()
    }

    if ($stepErrors.Count -gt 0) {
        $jsonObj.warnings += ("{0} step(s) reported errors." -f $stepErrors.Count)
    }

    Write-Output ($jsonObj | ConvertTo-Json -Depth 12 -Compress)
}

exit $exitCode