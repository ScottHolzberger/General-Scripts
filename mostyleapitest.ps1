# ===== CONFIG =====
$MosyleApiKey  = "745d7f77c5a25d727c477a7275553e33a7d7e22a334c7f7f52a6656f37755dae"
$MosyleEmail   = "support@zahezone.com.au"
$MosylePassword= "*Gx4xT0@&G51xpkZUGhpLj^*"

$Url = "https://businessapi.mosyle.com/v1/login"

# ===== CLEAN INPUT =====
$MosyleApiKey   = $MosyleApiKey.Trim()
$MosyleEmail    = $MosyleEmail.Trim()
$MosylePassword = $MosylePassword.TrimEnd("`r","`n")

# ===== HEADERS =====
$Headers = @{
    "accessToken"  = $MosyleApiKey
    "Content-Type" = "application/json"
    "Accept"       = "application/json"
}

# ===== BODY =====
$Body = @{
    email    = $MosyleEmail
    password = $MosylePassword
} | ConvertTo-Json -Depth 3

# ===== EXECUTE =====
Write-Host "Testing Mosyle Login..." -ForegroundColor Cyan

try {
    $Response = Invoke-WebRequest -Method Post -Uri $Url -Headers $Headers -Body $Body -UseBasicParsing -ErrorAction Stop

    Write-Host "SUCCESS ✅" -ForegroundColor Green

    $AuthHeader = $Response.Headers["Authorization"]

    if ($AuthHeader) {
        Write-Host "Bearer Token Returned:" -ForegroundColor Yellow
        Write-Host $AuthHeader
    } else {
        Write-Host "WARNING: No Authorization header returned" -ForegroundColor Yellow
    }

    Write-Host "`nStatus Code:" $Response.StatusCode
}
catch {
    Write-Host "FAILED ❌" -ForegroundColor Red
    Write-Host $_.Exception.Message

    if ($_.Exception.Response) {
        try {
            $stream = $_.Exception.Response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream)
            $body   = $reader.ReadToEnd()
            $reader.Close()

            Write-Host "`nResponse Body:" -ForegroundColor Yellow
            Write-Host $body
        } catch {}
    }
}