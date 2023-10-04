<#  https://learn.microsoft.com/en-us/sql/powershell/sql-server-powershell?view=sql-server-ver16&viewFallbackFrom=sqlserver-ps
    https://learn.microsoft.com/en-us/powershell/module/sqlserver/invoke-sqlcmd?view=sqlserver-ps
    
#>

# A simple SELECT
$sql = @'
SELECT  [OSUPSEmplID], 
        [OSUNameNActive], 
        [OSUmedcenterID] 
FROM    [OCIOIdentityDB].[dbo].[IDM_PEOPLE] 
WHERE   [OSUMedCenterID] in ('stah06','mazi01')
'@

Invoke-Sqlcmd -Encrypt Optional -Query $sql -ServerInstance 'EnterpriseDirDB.osumc.edu\tp'


