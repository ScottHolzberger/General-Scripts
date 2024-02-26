if ($null -eq (Get-Module -Name MR.ModuleHelper)) {
    if ($null -eq (Get-PSRepository | Where-Object Name -eq "DucentisRepo")) {
        Write-Host "Registering DucentisRepo..."
        Register-PSRepository -SourceLocation "\\dsiis02\PSRepo$" -PublishLocation "\\dsiis02\PSRepo$" -InstallationPolicy Trusted -Name "DucentisRepo"
    }
    Write-Host "Adding MR.ModuleHelper"
    Install-Module MR.ModuleHelper
}
else {
    #TODO - Compare current version and update if different
}
"Test   ="

Import-Module MR.ModuleHelper
$SecretID = "emwlOdCbTXU2F5Y9ZRV4"
$SecretKey = "W2r0ZNqGte9bhSTIuidz8o7amCyfLg"

$MRAPPToken = Get-MRAPPToken -SecretID $SecretID -SecretKey $SecretKey
$MRAPPURI = "https://mrapprest.ip2me.com.au/api"
$MRAPPHeader = @{
    Authorization = "Bearer $($MRAPPToken.token)"
}
function BeginSearch {
    Write-Host "Let's find a customer, do you want to search by JIM2Code, or Company Name?"
    Write-Host "[1] JIM2 Code"
    Write-Host "[2] Company Name"
    Write-Host "[Q] to quit"
    $Prompt = Read-Host "Enter the number"
    
    if ($Prompt.ToLower() -notin "1","2","q") {
        Write-Host "Invalid selection $Prompt - exiting"
        return
    }
    if ($Prompt.ToLower() -eq "q") {
        return
    }
    
    #JimCode
    
    if ($Prompt -eq 1) {
        $JimCode = Read-Host "Enter in the JIM2Code"
        $Company = Invoke-RestMethod -Method Get -URI "$MRAPPURI/v2/company?config=JIMCode=$JimCode" -Headers $MRAPPHeader -SkipHttpErrorCheck
        if ($Company.status -eq 404) {
            Write-Host "Jim Code not found"  -BackgroundColor Red
            BeginSearch
            return
        }
        if ($Company.statusCode -eq "200") {
            if ($Company.data.Count -gt 1) {
                Write-Host "More than one company found with that code, this should never occur - please inform Matt before you proceed"
                foreach ($comp in $Company.data) {
                    Write-Host "$($Comp.ID) - $($Comp.Name)"
                }
                return $null
            }
            elseif ($Company.data.Count -eq 0) {
                #This shouldn't be hit, but just in case it does get hit
                Write-Host "No Companies Found with the code $JimCode" -BackgroundColor Red
                BeginSearch
                return
            }
            elseif ($Company.data.count -eq 1) {
                Write-Host "Found Company - $($Company.Data[0].Name)"
                return $Company.Data[0]
            }
        }
    }
    elseif ($Prompt -eq 2) {
        $CompanyName = Read-Host "Enter in a partial, or full customer name"
        $Company = Invoke-RestMethod -Method GET -URI "$MRAPPURI/v2/company?name=$CompanyName" -Headers $MRAPPHeader -SkipHttpErrorCheck
        if ($Company.Status -eq 404 -or $Company -eq "") {
            Write-Host "Company not found" -BackgroundColor Red
            BeginSearch
            return
        }
        if ($Company.statusCode -eq "200") {
            if ($Company.data.Count -gt 1) {
                Write-Host "More than one company found with that name, please enter which company you wanted to search for"
                for ($i = 1; $i -le $Company.data.Count; $i++) {
                    Write-Host "[$($i)] - $($Company.data[$i - 1].Name)"
                }
                Write-Host "[Q] - Quit"
                $CompanySelection = Read-Host "Which Company is correct?"
                if ($CompanySelection.ToLower() -eq "q") {
                    return
                }
                elseif ($CompanySelection -notin 1..$($Company.data.Count)) {
                    Write-Host "Invalid Selection... $CompanySelection" -BackgroundColor Red
                    BeginSearch
                }
                return $Company.Data[$CompanySelection - 1]
            }
            return $Company.Data
        }
    }
}

$CloudID = "AVMZKORET9FG45X56A3I"
$CloudSecret = "N1DrHih/4j0mr4cfNuutRRifpGF0k/kayL6mllZp"

$MRAPPBody = @{
    "AccessKey" = $CloudID
    "SecretKey" = $CloudSecret
}

$Company = BeginSearch
Write-Host "Proceeding with $($Company.Name)"
Write-Host "Provisioning... Please wait..."
$Details = Invoke-RestMethod -Uri "$MRAPPURI/v2/company/$($Company.id)/onboarding/veeam" -Method POST -Body $($MRAPPBody | ConvertTo-Json) -Headers $MRAPPHeader -ContentType "application/json" -SkipHttpErrorCheck
if ($Details.isSuccesfull -eq $false) {
    Write-Host "Error - $($Details.message)" -BackgroundColor Red
}
else {
    Write-Host "Done"
}