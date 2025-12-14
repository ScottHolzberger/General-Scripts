
    # Check OS version
    $osVersion = [System.Environment]::OSVersion.Version
    $major = $osVersion.Major
    $minor = $osVersion.Minor
    $build = $osVersion.Build

    Write-Host "Detected OS Version: $major.$minor.$build"

    # Windows 10 starts at major=10, build >= 17763 for Storage Sense (1809)
    if ($major -eq 10 -and $build -ge 17763) {
        Write-Host "Compatible OS detected. Applying Storage Sense settings..."
    }
    elseif ($major -ge 11) {
        Write-Host "Windows 11 detected. Applying Storage Sense settings..."
    }
    else {
        Write-Host "Storage Sense is not supported on this OS version." -ForegroundColor Yellow
        Exit
    }

    # Registry paths
    $basePath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense"
    $policyPath = "$basePath\Parameters\StoragePolicy"

    # Ensure keys exist
    if (-not (Test-Path $policyPath)) {
        New-Item -Path $policyPath -Force | Out-Null
        Write-Host "Created registry path: $policyPath"
    }

    $StorageSenseKeys = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\StorageSense\Parameters\StoragePolicy\'
    Set-ItemProperty -Path $StorageSenseKeys -name '01' -value '1' -Type DWord  -Force  # Enable Storage Sense
    Set-ItemProperty -Path $StorageSenseKeys -name '02' -value '1' -Type DWord  -Force  	
    Set-ItemProperty -Path $StorageSenseKeys -name '04' -value '1' -Type DWord -Force   # Delete temporary files that my apps aren’t using
    Set-ItemProperty -Path $StorageSenseKeys -name '08' -value '1' -Type DWord -Force   # Delete files in my recycle bin if they have been there for over
    Set-ItemProperty -Path $StorageSenseKeys -name '32' -value '1' -Type DWord -Force   # Delete files in my Downloads folder if they have been there for ovr
    Set-ItemProperty -Path $StorageSenseKeys -name '128' -value '0' -Type DWord -Force
    Set-ItemProperty -Path $StorageSenseKeys -name '256' -value '14' -Type DWord -Force # Number of days (Recycle Bin)
    Set-ItemProperty -Path $StorageSenseKeys -name '512' -value '30' -Type DWord -Force # Number of days (Downloads)
    Set-ItemProperty -Path $StorageSenseKeys -name '1024' -value '1' -Type DWord -Force
    Set-ItemProperty -Path $StorageSenseKeys -name '2048' -value '1' -Type DWord -Force # Storage Sense frequency (run schedule)
    Set-ItemProperty -Path $StorageSenseKeys -name 'CloudfilePolicyConsent' -value '1' -Type DWord -Force
    
        $CurrentUserSID = ([System.Security.Principal.WindowsIdentity]::GetCurrent()).User.Value
        $CurrentSites = Get-ItemProperty 'HKCU:\SOFTWARE\Microsoft\OneDrive\Accounts\Business1\ScopeIdToMountPointPathCache' -ErrorAction SilentlyContinue | Select-Object -Property * -ExcludeProperty PSPath, PsParentPath, PSChildname, PSDrive, PsProvider
        foreach ($OneDriveSite in $CurrentSites.psobject.properties.name) {
            New-Item "$($StorageSenseKeys)/OneDrive!$($CurrentUserSID)!Business1|$($OneDriveSite)" -Force
            New-ItemProperty "$($StorageSenseKeys)/OneDrive!$($CurrentUserSID)!Business1|$($OneDriveSite)" -Name '02' -Value '1' -type DWORD -Force
            New-ItemProperty "$($StorageSenseKeys)/OneDrive!$($CurrentUserSID)!Business1|$($OneDriveSite)" -Name '128' -Value '14' -type DWORD -Force # Number of days (OneDrive Sync)
        }


