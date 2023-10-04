break
<#  Assuming that the module is installed, if not, read this:
    https://learn.microsoft.com/en-us/powershell/module/activedirectory/?view=windowsserver2022-ps
    
#>
# What does the module give us?
Get-Command -Module ActiveDirectory 
Get-Command -Module ActiveDirectory | Out-GridView
Get-Command -Module ActiveDirectory | Out-GridView -Title "AD Cmdlets"
Get-Command -Module ActiveDirectory -Verb Get | Out-GridView
Get-Command -Module ActiveDirectory -Noun ADUser | Out-GridView

# examples
Get-ADUser crews2
Get-ADUser crews2 -Properties Description, Title
Get-ADUser -Filter {samaccountname -like "crew*"} -Properties Department, Title | Select-Object Name, Enabled, Department, Title
Get-ADUser -LDAPFilter "(&(title=*)(samaccountname=crew*))" -Properties Department, Title | Select-Object Name, Enabled, Department, Title

Search-ADAccount -LockedOut
Get-ADPrincipalGroupMembership crews2 | Sort-Object Name | Select-Object Name