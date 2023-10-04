<#  Putting it all together....
    Scenario:
    Your manager gives you a spreadsheet and asks you to add additional data fields.
    You know these extra fields exist in a SQL table and in Active Directory

    She then asks that you export the data out to Excel and make sure it is cosmetically pleasing.
    
    We will be using the following modules:
    Native cmdlets,
    ActiveDirectory,
    ImportExcel,
    SQLServer
#>

# Retrieve the data from the manager supplied spreadsheet
$data = Import-Excel -Path C:\temp\ManagerDeliveredData.xlsx

# Add the EmployeeID Number from a SQL database
$revisedData = $data | Add-Member -MemberType ScriptProperty -Name GroupCount -Value {
    Get-ADPrincipalGroupMembership $this.ID | Measure-Object | Select-Object -ExpandProperty Count
} -PassThru | Add-Member -MemberType ScriptProperty -Name EMPID -Value {
    $id = "'{0}'" -f $this.ID 
    $sql = "SELECT [OSUPSEmplID] FROM [OCIOIdentityDB].[dbo].[IDM_PEOPLE] WHERE [OSUMedCenterID] = $id"
    (Invoke-Sqlcmd -Encrypt Optional -Query $sql -ServerInstance 'EnterpriseDirDB.osumc.edu\tp').OSUPSEmplID
} -PassThru


# define our chart parameters (introducing splatting)
$chartParameters = @{
    ChartType = 'Pie'
    XRange    = 'ID'
    YRange    = 'GroupCount'
    Title     = 'Group Count'
    TitleBold = $true
    Column    = 1
    Row       = 8
    Width     = 400 
    Height    = 300
    NoLegend  = $false    
}
$chart = New-ExcelChartDefinition @chartParameters

# define excel parameters
$ExcelParameters = @{
    InputObject          = $revisedData  
    Show                 = $true
    AutoSize             = $true
    ExcelChartDefinition = $chart
    AutoNameRange        = $true
}
Export-Excel @ExcelParameters


