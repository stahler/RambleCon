break

# Install module to manage secrets
Install-Module -Name Microsoft.PowerShell.SecretManagement -Repository PSGallery

# Install local store extension.  There are others
Install-Module -Name Microsoft.PowerShell.SecretStore -Repository PSGallery

Get-SecretVault # see if any vaults are registered

# Set up default vault
Register-SecretVault -Name MySecrets -ModuleName Microsoft.PowerShell.SecretStore -DefaultVault

# create a secret
Set-Secret -Name ExampleSecret -Secret "Super Secure Password"

# You may be asked to set up the password for the vault if you haven't already

# let's get the secret
Get-Secret -Name ExampleSecret -AsPlainText

Set-Secret -Name ChatGPT -Secret "Nope not telling you" -WhatIf
Set-OpenAIKey -Key (Get-Secret -Name ChatGPT)