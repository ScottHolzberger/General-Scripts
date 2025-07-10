$ApplicationId         = '10726a6e-39a1-46df-b178-1f49a9cd89ac'
$ApplicationSecret     = 'fWLly5c1Xzs0wVxw1JN9N9r4rXQuR8x1SHzemhkswo4=' | Convertto-SecureString -AsPlainText -Force
$TenantID              = 'caa9ec21-bcab-4544-af18-bfc520f7d18b'
$RefreshToken          = '0.AWcAIeypyqu8REWvGL_FIPfRi25qchChOd9GsXgfSanNiaxnACY.AgABAAEAAADnfolhJpSnRYB1SVj-Hgd8AgDs_wUA9P_40Y5sg2OroptpR2lhlftfJV4xtJGO2lKbxCQMICcSMNVb1u0078c5o_6LDBL4CCPrVlwD71kBOe1KbkhwAprrQto-JsjyACAy25x4YdkHTT05ks4Z81OUrUg3rfuTE1--MM5FdD1CJk_T_eFRdeDHEULGRkHCrqYCx5cvd3zF6okaZrrNyicM3lgjexfHNuuFfwnvW4l1CALKeIZAWNTt0p-oZrmJdjgmY3C4g831FFSnBE3RJ-1IT2aJxGl9BWwzUhuW3ltjqgBGLWXdofgh2fNqnHzTj9qYlcmWdcWGRO5BXxrmWLXdOG5-MfHc7R7imYrJ0vb4HOH5w8A4XTxiAmCTx0Y1uVk1ImS3NdcftLyKVBZhyMKgHiEU_NBRhZV_zf-l4JoDWG24155RzYEhsbMV_2s9kd0s9PkPxLHlJaDBSpCQqkXu6ArMqAlhJo8CfQcLhQjYkK-9qT6DPDj5s-kY6TKKXN9DZhX6MMl6qwEkY2Fec8AGvhyOEFUcmWBNkEOmCf_nXawdQ9IU2hqUnlhIoemnPDcEkW-6sQKl5-fRrVq3y0QA5RUOxcLaRcTtrNM8T-hlwLA0BTm_ei2LqBPO6E1KwFlXFHhuc5Wppp8ZvPorJFefC0bSIzFKYlXswv4l_YDAA4rgjruFrwQKgIvc8lkU0Y0eT_ImJETkmogqjpxkjJW4kI8HnfOL1vFTLqgqVIMfccQWwIcgJ1OknAnOK6TPbCbd8Jboe_wN3TBslMPre57lalTX2faZjuypSouTF9d8m2vYIRj0scItmuBXd0AO5jOSxzcYaImxV5LaqUZwMKrTkhf3j-mh6Uqg4YmDz7ettbGqhtxF3GFXRNw-weZXMdbeIhhz3KvfG18uIj7TosgorAT6ya1RwxCa1BO7FJpc_Yfh8_V5MHMPKgimeWa2xlbqk-dOZ7Y'
$ExchangeRefreshToken  = '0.AWcAIeypyqu8REWvGL_FIPfRixY8x6Djp2RFmpUr30c4NxZnACY.AgABAAEAAADnfolhJpSnRYB1SVj-Hgd8AgDs_wUA9P_bh2Hmxpc84EIMNPpJtWL2dPU-7ZXorrci6hAfYv79eTtTQbtZmVpd4uoOL0OBrRTK-Sj7i0QH1XoQ_GNjHvE9YfJHnRxm-JJcaLWxqU3tJFJvjg33LjzyZH6gWz7dn3mtK532f9ySNLDMhd-iNfk413acWMGoMRYHHSEP_MjUrctUWVa2fuFYYoR5Pp0q1JLWIvEKHRiV7ccFDOkS94uv-Gmsfapgqnl6NWDnCL3toLllcZnFhCUZziGWyjbIv3SOaX5AEuGXvhqCBHawWocjHlS5GAckB-4sdWLd5xxeVUbgr2kQ1RvZb3cErDTHfUglfv-XcNrAQbHaAY9Y9tD5-djUxkqcY0jQlqU5k5jqI15ZyghT9q_5KEc6T25enMaiqhRQdZQPk6j41IDVHvDqkLN3gQIxGW9EAl8i60xRa_xxEaA3SurgNP5CgY955Imlpnl8wLfy4Zx3JeazDEWhWpD-BscV5xcX0F1853NcbaYflm8xjEJ5VtDHJMtJOZ5DuDlZypWQUQEy3PbjSlB4WmcgWnFtwTnevdx5tvB7L3QTzZ7SZ5yUC37Y9xPdJd-rcpQso-9clOusPqSLvAs1f54hGFMaEuFZHgGNHqq658Ur8Wm8AHA7kLeGSf6TbdEXuAzUoAFTtov42_J0fVhn8lkjrNL9n0iP3toEvmiU4Ad_xb6SJMZqY8YkG2qAK-O7MBUrvGIkDEEIyVpd8jQP5J7bSW_nsnhSoDp-gxSYBJjdeGX1hMss45zve3NjSXo0V818hWIfjCWMFC-6pI5fNSrvp33C7SM'
$upn                   = 'cipp@zahe.onmicrosoft.com'

$credential = New-Object System.Management.Automation.PSCredential($ApplicationId, $ApplicationSecret)

$aadGraphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.windows.net/.default' -ServicePrincipal -Tenant $tenantID
$graphToken = New-PartnerAccessToken -ApplicationId $ApplicationId -Credential $credential -RefreshToken $refreshToken -Scopes 'https://graph.microsoft.com/.default' -ServicePrincipal -Tenant $tenantID

Connect-MsolService -AdGraphAccessToken $aadGraphToken.AccessToken -MsGraphAccessToken $graphToken.AccessToken
$customers = Get-MsolPartnerContract -All
foreach ($customer in $customers) {

    $customerId = $customer.DefaultDomainName

  <#  write-host "Connecting to the Security Center for client $($customer.name)"
    $SCCToken = New-PartnerAccessToken -ApplicationId 'a0c73c16-a7e3-4564-9a95-2bdf47383716'-RefreshToken $ExchangeRefreshToken -Scopes 'https://outlook.office365.com/.default'
    $SCCTokenValue = ConvertTo-SecureString "Bearer $($SCCToken.AccessToken)" -AsPlainText -Force
    $SCCcredential = New-Object System.Management.Automation.PSCredential($upn, $SCCTokenValue)
    $SccSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.compliance.protection.outlook.com/powershell-liveid?BasicAuthToOAuthConversion=true&DelegatedOrg=$($customerId)" -Credential $SCCcredential -AllowRedirection -Authentication Basic
    import-session $SccSession -disablenamechecking -allowclobber
    #YourCommands here

    #/End of Commands

    Remove-session $SccSession
    write-host "Connecting to the Exchange managed console for client $($customer.name)" #>

    Write-host "Enabling all settings for $($Customer.Name)" -ForegroundColor Green
    $token = New-PartnerAccessToken -ApplicationId 'a0c73c16-a7e3-4564-9a95-2bdf47383716'-RefreshToken $ExchangeRefreshToken -Scopes 'https://outlook.office365.com/.default' -Tenant $customer.TenantId
    $tokenValue = ConvertTo-SecureString "Bearer $($token.AccessToken)" -AsPlainText -Force
    $credentialExchange = New-Object System.Management.Automation.PSCredential($upn, $tokenValue)

    $ExchangeOnlineSession = Connect-ExchangeOnline -ConfigurationName Microsoft.Exchange -ConnectionUri "https://ps.outlook.com/powershell-liveid?DelegatedOrg=$($customerId)&BasicAuthToOAuthConversion=true" -Credential $credentialExchange -Authentication  -AllowRedirection -erroraction Stop
    Import-PSSession -Session $ExchangeOnlineSession -AllowClobber -DisableNameChecking
    #YourCommands here

    #/End of Commands
    Remove-PSSession $ExchangeOnlineSession
}