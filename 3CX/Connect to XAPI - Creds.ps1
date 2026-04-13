$URL = "https://3cx.zahezone.com.au"

$Username = "zahezone@zahezone.com.au"
$plainPassword = "Cte946da!@#"
$MFA = ""

$body = @{
    SecurityCode = $MFA
    Username     = $Username
    Password     = $plainPassword
}

$bodyJson = $body | ConvertTo-Json -Depth 3 -Compress

$response = Invoke-RestMethod `
    -Uri "$URL/webclient/api/Login/GetAccessToken" `
    -Method POST `
    -ContentType "application/json" `
    -Body $bodyJson `
    -TimeoutSec 30

# ----- SAFE token extraction (StrictMode compatible) -----
$token = $null
foreach ($name in 'access_token','AccessToken','Token') {
    if ($response.PSObject.Properties[$name]) {
        $token = $response.$name
        break
    }
}

if (-not $token) {
    throw "Login succeeded but no token was returned.`n$($response | ConvertTo-Json -Depth 5)"
}

#$headers = @{
#    Authorization = "Bearer $token"
#}

$headers = @{ Authorization = "Bearer $($token.access_token)" }

Invoke-RestMethod `
    -Uri "$URL/xapi/v1/systemstatus" `
    -Headers $headers

#$sys = Invoke-RestMethod -Uri "$URL/xapi/v1/systemstatus" -Headers $headers
#$sys | ConvertTo-Json -Depth 50 | Out-File "systemstatus-$($sys.FQDN).json" -Encoding utf8
#$sys