$Groups = Get-DistributionGroup -ResultSize Unlimited

$Groups | ForEach-Object {
$group = $_
Get-DistributionGroupMember -Identity $group.Name -ResultSize Unlimited | ForEach-Object {
      New-Object -TypeName PSObject -Property @{
       Group = $group.DisplayName
       Member = $_.Name
       EmailAddress = $_.PrimarySMTPAddress
       RecipientType= $_.RecipientType
}}} | Export-CSV "C:\Temp\Office365GroupMembers.csv" -NoTypeInformation -Encoding UTF8