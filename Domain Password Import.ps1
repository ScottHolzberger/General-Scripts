
$checkDomain = Import-CSV C:\Temp\Domains.csv
$DomainPassword = New-Object -TypeName System.Collections.ArrayList
$domainPassword.Add("DomainPassword")


foreach ($name in $checkDomain){

    $method = "https://theconsole.tppwholesale.com.au/api/query.pl?SessionID=u1mASnPIKBD4tddmfDGlOjWGF&Type=Domains&Object=Domain&Action=Details&Domain="+$name.DomainPassword
    $domain = Invoke-RestMethod $method
    $splitDpmain = $domain.Split()
    $text = $splitDpmain[3] -replace "DomainPassword="
    $DomainPassword.Add($text)
}



