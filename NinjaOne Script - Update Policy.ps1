#Connect-NinjaOne -Instance 'oc' -ClientId 'b6ju_20q5KthapaP9wEXoP4Ldrs' -ClientSecret 'HMCmapaddFMvaJsG660ZEZWMkAbE8FFDmdB3nqvHMPtyHOG_smFUsA' -UseClientAuth




$devices = Get-NinjaOneDevices -organisationId 7

foreach ($device in $devices)
{
    #Write-Host "Device ID: "$device.id
    $Check = Get-NinjaOneDeviceCustomFields -deviceId $device.id
    if ($Check.patchPilot) 
    {
        if ($device.PolicyId -eq 75) 
        {
           #Write-Host "Pilot Policy is Enabled - " $device.PolicyId
        }
        else 
        {
           if ($device.nodeClass -eq "WINDOWS_SERVER") 
           {
                #Write-Host "This is a server - no change"
           }
           else 
           {
           # Write-Host "Policy has been updated - " $device.PolicyId
            Set-NinjaOneDevice -deviceId $device.id -deviceInformation @{ policyId = 75 }
           }
        }
    }
    
    Else {
        Write-Host "Patch Pilot is not enabled"
        if ($device.nodeClass -eq "WINDOWS_SERVER") 
           {
                Write-Host "This is a server - no change"
           }
        elseif ($device.PolicyId -eq 73){
            Write-Host "Standard Policy is Enabled - " $device.PolicyId
            
           else 
           {
           # Write-Host "Policy has been updated - " $device.PolicyId
            Set-NinjaOneDevice -deviceId $device.id -deviceInformation @{ policyId = 75 }
           }
        }
        Else {
            #Write-Host "Policy has been updated - " $device.PolicyId
            Set-NinjaOneDevice -deviceId $device.id -deviceInformation @{ policyId = 73 }
        }
    }
    

}