
Connect-PnPOnline `
  -Url "https://zahe.sharepoint.com/sites/ZZ-Control" `
  -ClientId "01e1b71f-cbcb-48df-a076-871aa4ba10d9" `
  -Tenant "zahe.onmicrosoft.com" `
  -CertificatePath ".\ZaheZone-PnP-Projects.pfx" `
  -CertificatePassword (ConvertTo-SecureString "UseA-LongRandomPasswordHere" -AsPlainText -Force)
