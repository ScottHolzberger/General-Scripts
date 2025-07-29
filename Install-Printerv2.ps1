$DriverUnpackPath        = "C:\InTune\Kyocera\" # set this to the top level of the unpacked driver, it will search recursively for INFs.
$DriverName              = 'Kyocera ECOSYS M6230cidn KX' # $DriverName must be EXACTLY correct (you'll have to read the driver's .inf to find it out)
$PrinterIconName         = 'Kyocera TASKalfa 400ci KX'
$PortName                = '192.168.15.21' # change me
$printprocessor          = 'winprint'
$Datatype                = 'RAW'   
$PortNumber              = 9100  

# Test version of PNPUtil.exe
    $pnputilIsOld = (& pnputil /? | Select-String 'This usage screen' -quiet)
    if ($pnputilIsOld) {
        Write-Host "Note: PNPUtil.exe is the old version"
    }
# Find which INF to install with PNPUtil.exe
    $INFs = Get-ChildItem $DriverUnpackPath -Recurse -Filter "*.inf" -ErrorAction Stop 
    $FoundIt = $false
    foreach ($INF in $INFs) {
        $test = Get-Content $INF.FullName
        if (($test | Select-String $DriverName -quiet) -and ($test | Select-String -Pattern "Class.*=.*Printer" -quiet)) {
            $FoundIt = $true
            Write-Host "Found the correct INF: $($inf.FullName)"
            if ($pnputilIsOld) {
                Start-Process PNPUtil.exe -ArgumentList "-i -a $($inf.fullname)" -Wait
            } else {
                Start-Process PNPUtil.exe -ArgumentList "/add-driver $($inf.fullname) /install" -Wait
            }
        }
    }
    if ($false -eq $FoundIt) {
        Write-Host 'Aborting: was not able to install a driver with PNPUtil.exe.'
        Break Script
    }
# all set, add the driver for real:
Add-PrinterDriver -Name $DriverName -ErrorAction Stop -Verbose
# add the "icon" instance:
Add-Printer -Name $PrinterIconName -DriverName $DriverName -PortName $PortName -PrintProcessor $PrintProcessor -Datatype $Datatype -Verbose