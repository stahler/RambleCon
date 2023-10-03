# A simple SELECT
$sql = "SELECT [OSUPSEmplID], [OSUNameNActive], [OSUmedcenterID] 
FROM [OCIOIdentityDB].[dbo].[IDM_PEOPLE] 
WHERE [OSUMedCenterID] in ('stah06','mazi01')"
Invoke-Sqlcmd -Encrypt Optional -Query $sql -ServerInstance 'EnterpriseDirDB.osumc.edu\tp'


