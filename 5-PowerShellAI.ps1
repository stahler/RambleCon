break
# Get the module https://github.com/dfinke/PowerShellAI

# may need to get/set the key
Set-OpenAIKey -Key (Get-Secret -Name ChatGPT)

# simple request for translation
Get-GPT4Completion "Hello friends in Russian, Japanese and French"

# example of getting specific info and dictating format
Get-GPT4Completion  "List planets name,size and number of moons as json"

# building off last example, automate creating a spreadsheet
New-SpreadSheet -prompt "List planets name,size and number of moons"

# Example on writing code.
copilot -inputPrompt "Using PowerShell, get all disabled users with a last name like 'sta*' "

# get distinct domains from email, group and show where count is > 100
copilot "using PowerShell ActiveDirectory Module. Split Domain from email attribute then group by domain and show only where domain count is greater then 100"

$prompt = @'
You are a SQL Server expert.
Can you give me the SQL code to generate a table called Employees that has the following:
UniqueID
FirstName
LastName
AddressLine1
AddressLine2
City
State
Zip
'@

copilot $prompt

# allows for pipeline passing
ai "cheat sheet for regular expressions as a spreadsheet" | ConvertFrom-GPTMarkdownTable | Export-Excel

# exploring errors
Get-ADUser NoExist

# Get info on the error
Invoke-AIErrorHelper

# use the .net error for Try/Catch
try {
  $user = 'NoExist'  
  Get-ADUser $user -ErrorAction Stop
} 
catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
  Write-Warning -message "$user doesn't exist in Active Directory"
}

# Add Code editing example (adding Help)
$function = @'
function Get-DomainName {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   Position=0)]
        [Alias("email")] 
        [string[]]$eMailAddress
    )

    Process {
        foreach ($mail in $eMailAddress) {
            ($mail -split "@")[-1]
        }
    }
}
'@

$result = Get-OpenAIEdit -InputText $function -Instruction 'add comment-based help detailed description'

# demo explain for a one-liner
$users = Get-ADUser -LDAPFilter "(&(userAccountControl:1.2.840.113556.1.4.803:=2)(|(samAccountNAmwe=stah*)(samaccountname=rams*)))" | Select-Object SAMAccountName, surname, enabled
explain