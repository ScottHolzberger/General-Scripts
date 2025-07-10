Start-Transcript -Path F:\SQLStats.txt -Append

[string] $Server="GPSI-SVR-SBS-06"
[string] $Database= "GPSI_DATA"
[string] $SQLQuery= $("SELECT TOP (1) [ID], [SampleTime],[WaitingSBSA],[WaitingSBSB],[WaitingNotification] FROM [GPSI_DATA].[dbo].[PerformanceHighRes] ORDER BY [SampleTime] Desc")
[string] $user = "sa"
[string] $pass = "@MAz1ngh0rs3$"

    $alert = $false
    $alertText = $null

  $connectionString = "Server=$server;Database=$database;Uid=$user;Pwd=$pass"
  $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection -ArgumentList $connectionString
  

    $Connection.Open()
    $Command = New-Object System.Data.SQLClient.SQLCommand
    $Command.Connection = $Connection
    $Command.CommandText = $SQLQuery
    $Reader = $Command.ExecuteReader()
    while ($Reader.Read()) {
         $time = $Reader.GetValue(1)
         $SBSA = $Reader.GetValue(2)
         $SBSB = $Reader.GetValue(3)
         $WN  = $Reader.GetValue(4)
         
         Write-Host "Sample Time: " $time
         Write-Host "SBSA: " $sbsa
         Write-Host "SBSB: " $sbsb
         Write-Host "Notifications: " $WN

         $Reader
    }
    $Connection.Close()


    if ($sbsa -gt 20000){
        #Write-Host "SBSA Triggered"
        $alertText += "SBSA "
        $alert = $true
        }
    if ($sbsb -gt 20000){
        $alert = $true
        #Write-Host "SBSB Triggered"
        $alertText += "SBSB "
    }
    if ($WN -gt 10) {
        $alert = $true
        #Write-Host "WN Triggered"
        $alertText += "Notifications "
    }
    $run = (get-date).AddMinutes(-5)
    if ($time -le $run){
        $alert = $true
        #Write-Host "Time Triggered"
        $alertText += "Values not updated "
        }


    if ($alert){
        Write-Host "Alert Raised"
        $alertBody = @{Text=$alertText} | ConvertTo-Json
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-RestMethod -Method post -ContentType 'Application/Json' -Body $alertBody -Uri "https://prod-17.australiasoutheast.logic.azure.com:443/workflows/68ca483cc1df483f849e51c0d0683ad9/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=_MLnRnJ6mQLTT9zaSPM9G6UaFgHJ4iArovsmGyfj9pw"
        }
    else {
        Write-Host "All Checks Passed - No Alert"
        }
        
Stop-Transcript
    