<# Get the Module from the PowerShell Gallery
    https://www.powershellgallery.com/packages/ImportExcel/7.8.5
    https://github.com/dfinke/ImportExcel
    
    Doug Finke Microsoft MVP since forever...
    Check out his GitHub repo: https://github.com/dfinke
#>

# Let's find the module
Find-Module ImportExcel

# Install the module
Find-Module ImportExcel | Import-Module ImportExcel

# Import the module into our session
Import-Module ImportExcel

# Take a peek at the module 
Get-Module ImportExcel | Format-List

# Explore the various cmdlets it has to offer
Get-Command -Module ImportExcel | Out-GridView


