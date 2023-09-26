<#  Assuming that the module is installed, if not, read this:
    https://learn.microsoft.com/en-us/powershell/module/activedirectory/?view=windowsserver2022-ps
    
#>

Get-Command -Module ActiveDirectory | Out-GridView
Get-Command -Module ActiveDirectory -Verb Get | Out-GridView
Get-Command -Module ActiveDirectory -Noun ADUser | Out-GridView