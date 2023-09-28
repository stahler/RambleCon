break
<# Most important cmdlets!
    Get-Help
    Get-Command
    Get-Member

    Note Verb-Noun convention
#>

# Get-Help
Update-Help # With some frequency (I do it weekly), update the help
Get-Help Get-Process
Get-Help Get-Process -Examples
Get-Help Get-Process -Online
Get-Help Get-Process -ShowWindow

# Get-Command
Get-Command *IPC*
Get-Command -Noun Process
Get-Command -verb Get 
Get-Command | Measure-Object
Get-Command -Module ActiveDirectory | Measure-Object
Get-Command -Module ActiveDirectory -Noun ADUser

# Get-Member - lists the properties and methods of an object
Get-ADUser stah06 | Get-Member
$user = Get-ADUser stah06
$user | Get-Member
$user.Surname # Show a property

notepad.exe
$process = Get-Process notepad
$process | Get-Member -MemberType Method # Show a method
$process.Kill()

# Exporting data, better ways ahead
Get-Process | Select-Object ProcessName, Id | Export-Csv -Path C:\temp\process.csv
Invoke-Item C:\temp\process.csv