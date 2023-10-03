<#
Code requirements:
Using Active Directory Module, find all disabled users where surname is like 'sta*' or 'rams*'.
Return SAMAccountName, DisplayName and LastLogonDate sorted by DisplayName, export to OGV.
#>

# Get all disabled users where surname is like 'sta*' or 'rams*'
$users = Get-ADUser -LDAPFilter "(&(userAccountControl:1.2.840.113556.1.4.803:=2)(|(samAccountName=stah*)(samaccountname=rams*)))" -Properties LAstLogonDate, DisplayName | Select-Object SAMAccountName, surname, enabled, LAstLogonDate, DisplayName

# Return SAMAccountName, DisplayName and LastLogonDate sorted by DisplayName, export to OGV.
$users | Sort-Object DisplayName | Select-Object SAMAccountName, DisplayName, LastLogonDate | Out-GridView
