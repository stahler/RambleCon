break
<# Get the Module from the PowerShell Gallery
    https://www.powershellgallery.com/packages/ImportExcel/7.8.5
    https://github.com/dfinke/ImportExcel
    
    Doug Finke Microsoft MVP since forever and a great guy...
    Check out his GitHub repo: https://github.com/dfinke
#>

#region - metadata about the module ##########################################################
# Let's find the module
Find-Module ImportExcel

# Install the module
Find-Module ImportExcel | Install-Module ImportExcel

# Import the module into our session
Import-Module ImportExcel

# Take a peek at the module 
Get-Module ImportExcel | Format-List

# Explore the various cmdlets it has to offer
Get-Command -Module ImportExcel | Out-GridView
#endregion

#region - Conditional Formatting    ##########################################################
# Let's kick out some data to Excel (simple example)
Get-ADUser -Filter {samaccountname -like "Stah*"} -Properties Department, Title | 
Select-Object Name, Enabled, Department, Title | Export-Excel

# Simple Conditional Formatting
$disabled = New-ConditionalText False #White Red
$enabled = New-ConditionalText True white Green
Get-ADUser -Filter {samaccountname -like "Stah*"} -Properties Department, Title | 
Select-Object Name, Enabled, Department, Title | Export-Excel -ConditionalText $disabled #, $enabled

# Example of IconSets
# https://gist.github.com/stahler/97bb228b9d43c3ba4f1194cb7ea95880#file-add-iconset-ps1

# Example comparing groups
$Mark = Get-ADPrincipalGroupMembership ambl01 | Sort-Object Name | Select-Object @{N='Mark';E={$PSItem.Name}}
$Wes =  Get-ADPrincipalGroupMembership stah06 | Sort-Object Name | Select-Object @{N='Wes';E={$PSItem.Name}}

$xlfile = "c:\temp\groups.xlsx"
Remove-Item $xlfile -ErrorAction SilentlyContinue

$wsName = "Unique Values"
$c1 = New-ConditionalText -ConditionalType UniqueValues -Range '$A:$C' 

$Mark | Export-Excel $xlfile -WorksheetName $wsName
$Wes | Export-Excel $xlfile -WorksheetName $wsName -StartColumn 3 -ConditionalText $c1 -Show
#endregion

#region - Charts                    ##########################################################
# Add a Chart
$path = "$env:TEMP\ExampleChart.xlsx"
Remove-Item -Path $path -ErrorAction SilentlyContinue

# create some fake data
$data = ConvertFrom-Csv @"
Name,Weight
Shawn,380
Thom,215
West,210
"@

# define our chart parameters
$chartParameters = @{
    ChartType = 'ColumnClustered'
    XRange    = 'Name'
    YRange    = 'Weight'
    Title     = 'Example Chart Title'
    TitleBold = $true
    NoLegend  = $true
    Column    = 3
    Row       = 0
    Width     = 300 
    Height    = 200    
}
$chart = New-ExcelChartDefinition @chartParameters

# define excel parameters
$ExcelParameters = @{
    InputObject          = $data
    Show                 = $true
    AutoSize             = $true
    ExcelChartDefinition = $chart
    AutoNameRange        = $true
}
Export-Excel @ExcelParameters

# Multiple Charts
# create some "data", note the formula in th 4th column
$data = ConvertFrom-Csv @"
Name,Weight, Height,Ratio
Shawn,350,76,"=b2/c2"
Wes,210,69,"=b3/c3"
Thom, 205,73,"=b4/c4"
Sam,105,69,"=b5/c5"
"@

# paramters for our pie chart
$pie = @{
    ChartType    = 'Pie3D'
    XRange       = 'name'
    YRange       = 'weight'
    Height       = 200
    ShowCategory = $true
    NoLegend     = $true
    Width        = 225
    Row          = 5
    Column       = 3
    Title        = "Mmmm Pie"
}
$pie = New-ExcelChart @pie

# parameters for our bar chart
$bar = @{
    XRange   = 'Name'
    YRange   = 'weight'
    Height   = 200
    Row      = 5
    Column   = 0
    NoLegend = $true 
    Width    = 190
    Title    = "Bar hopping"
}
New-ExcelChart @bar

# pass our data with the two chart definitions
$data | Export-Excel -AutoSize -AutoNameRange -ExcelChartDefinition $bar, $pie
#endregion

#region - Pivot tables              ##########################################################
Get-ADUser -Filter {samaccountname -like "Stah*"} -Properties Department, Title | 
Select-Object Name, Enabled, Department, Title | 
Export-Excel -IncludePivotTable -PivotRows Enabled -PivotData @{Enabled='Count'}

# Create Multiple Pivot Tables
$xlFile = "C:\TEMP\MultiplePivotTables.xlsx"
Remove-Item -Path $xlFile -ErrorAction SilentlyContinue

# create some data 
$data = ConvertFrom-Csv @"
Name,State,City
Thom,Ohio,Columbus
West,Ohio,Dayton
Chris,Ohio,Columbus
Gern,Michigan,Holland
Hari,Michigan,Holland
Shawn,Michigan,Grand Rapids
Todd,Michigan,Flint
"@

# parameters for Data worksheet {Splatting}
$params = @{
    Path          = $xlFile
    WorkSheetName = 'Data'
    InputObject   = $data 
    AutoSize      = $true
    BoldTopRow    = $true
    FreezeTopRow  = $true
    PassThru      = $true
}

$xl = Export-Excel @params 

$null = Add-Worksheet -ExcelPackage $xl -WorkSheetname "PivotTable"

# parameters for pivotTable (By State)
$pivotTableParams = @{
    PivotTableName  = 'ByState'
    Address         = $xl.PivotTable.cells["A1"]
    SourceWorkSheet = $xl.Data
    PivotRows       = 'State'
    PivotData       = @{'Name' = 'Count' }
    PivotTableStyle = 'Light20'
}

# add first pivotTable (By City)
$null = Add-PivotTable @pivotTableParams

# modify pivotTable parameters and add the second one.
$pivotTableParams.Address = $xl.PivotTable.cells["D1"]
$pivotTableParams.PivotTableName = 'ByCity'
$pivotTableParams.PivotRows = 'City'
$null = Add-PivotTable @pivotTableParams

Close-ExcelPackage $xl -Show
#endregion                          ##########################################################

#region - other examples            ##########################################################
# adding sparklines - https://gist.github.com/stahler/7b5c0c8b6347b010bf9bce64b3c29326#file-add-sparkline-ps1
# adding hyperlinks - https://gist.github.com/stahler/47b0842350f1177da8c26308809ee769
# adding formulas - https://gist.github.com/stahler/51e84413e4c05d9a563dc7cac766b5cc
# Keeping leading zeros - https://gist.github.com/stahler/72e36020b5c24d41ed9988f3f5a7dd48#file-keep-leadingzero-ps1

# Example of creating multiple worksheets
$groups =   'Domain Admins',
            'Group Policy Creator Owners', 
            'EEI-OU-ADMIN',
            'accounts.change.limited',
            'Level 2 Support'

$dt = Get-Date -Format yyyyMMdd
$path = "C:\temp\PriviledgedGroups$dt.xlsx"
Remove-Item -Path $path -ErrorAction SilentlyContinue

foreach ($group in $groups) {
    Get-ADGroupMember $group | 
    Select-Object SAMAccountName, Name, DistinguishedName | 
    Export-Excel -Path $path -WorksheetName "$($group)" -AutoSize
}

Invoke-Item -Path $path
#endregion                          ##########################################################


