$base = "https://zahezone.3cx.com.au" # prefer 443 base first
$clientId = "zzmonitoring"
$clientSecret = "0Snw9pTrcSyKpud3Ilv8whQ5I8JnJ1Ae"

$token = Invoke-RestMethod -Method POST `
  -Uri "$base/connect/token" `
  -ContentType "application/x-www-form-urlencoded" `
  -Body @{
    grant_type    = "client_credentials"
    client_id     = $clientId
    client_secret = $clientSecret
  }

$headers = @{ Authorization = "Bearer $($token.access_token)" }


Invoke-RestMethod -Method GET `
  -Uri "$base/xapi/v1/SystemStatus" `
  -Headers $headers
