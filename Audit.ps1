import-module ActiveDirectory
$Header = @"
<style>
TABLE {border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
TH {border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;}
TD {border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
P {font-size:24px; font-weight:bold;}
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
</style>
<title>
AD Audit
</title>
"@
$Pre = get-date
$after = "<p>Total Users: "+(Get-ADUser -Filter "Enabled -eq 'True' -AND PasswordNeverExpires -eq 'False'").count+"</p>"
Get-ADUser -Filter "Enabled -eq 'True' -AND PasswordNeverExpires -eq 'False'" -Properties Department,PasswordLastSet,PasswordNeverExpires,PasswordExpired | Select DistinguishedName,Name,Department,pass*,@{Name="PasswordAge"; Expression={(Get-Date)-$_.PasswordLastSet}} |sort Name | ConvertTo-Html -Head $Header -PreContent $Pre -PostContent $after | Out-File C:\ADaudit.html
