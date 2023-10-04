<#
Code requirements:
Using Active Directory Module, find all disabled users where surname is like 'sta*' or 'rams*'.
Return SAMAccountName, DisplayName and LastLogonDate sorted by DisplayName, export to OGV.
#>

# Example on writing code.
copilot -inputPrompt "Using PowerShell, create a LDAPFilter and then use it to get all disabled users with a last name like 'sta*' "

