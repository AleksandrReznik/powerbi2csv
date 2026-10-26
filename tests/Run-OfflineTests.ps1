# Offline tests of powerbi2csv.ps1: the whole pipeline runs against a fake tenant (tests\lib\TestWorld.ps1), nothing
# touches the network. Works on Windows PowerShell 5.1 and PowerShell 7.
#   .\tests\Run-OfflineTests.ps1                 this PowerShell
#   .\tests\Run-OfflineTests.ps1 -AllEditions    powershell.exe and pwsh
param([switch] $AllEditions, [switch] $KeepOutput)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot
$script = Join-Path $root 'powerbi2csv.ps1'

if ($AllEditions) {
    $failed = 0
    foreach ($exe in @('powershell.exe', 'pwsh')) {
        if (-not (Get-Command $exe -ErrorAction SilentlyContinue)) { Write-Host ($exe + ' not found - skipped') -ForegroundColor Yellow; continue }
        Write-Host ('=== ' + $exe + ' ===') -ForegroundColor Cyan
        & $exe -NoProfile -ExecutionPolicy Bypass -File $PSCommandPath
        if ($LASTEXITCODE -ne 0) { $failed++ }
    }
    exit $failed
}

. (Join-Path $PSScriptRoot 'lib\Assert.ps1')
. $script
. (Join-Path $PSScriptRoot 'lib\TestWorld.ps1')
. (Join-Path $PSScriptRoot 'lib\FakeTransport.ps1')

$global:TWOut = Join-Path ([IO.Path]::GetTempPath()) ('pbi2csv-tests-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $global:TWOut -Force | Out-Null

function Get-TWCoverage { param($Run, [string] $Area) return @($Run.Coverage | Where-Object { $_['Area'] -eq $Area })[0] }

Write-Host 'Static checks' -ForegroundColor Cyan
Test-Case 'script parses without errors' {
    $errs = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($script, [ref]$null, [ref]$errs)
    Assert-Equal 0 @($errs).Count
}
Test-Case 'source is ASCII only (Windows PowerShell reads BOM-less files as ANSI)' {
    $bad = @([IO.File]::ReadAllBytes($script) | Where-Object { $_ -gt 127 })
    Assert-Equal 0 $bad.Count
}
Test-Case 'every catalog entry has a dataset function, a report function and known datasets' {
    foreach ($r in $script:A2CReportCatalog) {
        Assert-True (Get-Command ('New-A2CReport' + $r.Name) -ErrorAction SilentlyContinue) ('report function of ' + $r.Name)
        foreach ($d in $r.Needs) { Assert-True $script:A2CDatasetCatalog.Contains($d) ('dataset ' + $d) }
    }
    foreach ($d in $script:A2CDatasetCatalog.Keys) { Assert-True (Get-Command ('Get-A2CDataset' + $d) -ErrorAction SilentlyContinue) ('dataset function of ' + $d) }
}
if (Get-Module -ListAvailable -Name PSScriptAnalyzer) {
    Test-Case 'PSScriptAnalyzer reports nothing' {
        $found = @(Invoke-ScriptAnalyzer -Path $script -Settings (Join-Path $root 'PSScriptAnalyzerSettings.psd1'))
        Assert-Equal 0 $found.Count (($found | Select-Object -First 5 | ForEach-Object { '{0}:{1}' -f $_.Line, $_.RuleName }) -join ', ')
    }
}
else { Write-Host '  SKIP  PSScriptAnalyzer not installed' -ForegroundColor Yellow }

Write-Host 'Options' -ForegroundColor Cyan
Test-Case 'unknown report name is refused' { Assert-Throws { New-A2COptions -Params @{ Reports = @('Nope') } } 'Unknown report' }
Test-Case 'comma list from pwsh -File is put together again' {
    $p = [ordered]@{ Reports = @('PowerBI,'); RemainingArguments = @('Fabric') }
    Assert-Equal 'PowerBI Fabric' ((New-A2COptions -Params $p).Reports -join ' ')
}
Test-Case 'profiles: Standard has no opt-in reports, Full has all but the export' {
    $std = Resolve-A2CSelection (New-A2COptions -Params @{})
    Assert-Equal 'PowerBI Fabric FabricSql Dependencies PowerBIAdmin' ($std -join ' ')
    $full = Resolve-A2CSelection (New-A2COptions -Params @{ ReportProfile = 'Full' })
    Assert-True ($full -contains 'PowerBIModels'); Assert-True ($full -contains 'PowerBIActivity'); Assert-False ($full -contains 'PowerBIExport')
}

Write-Host 'Standard run as a workspace member (no admin role)' -ForegroundColor Cyan
Reset-TWState
$global:TW.Throttle.Add([pscustomobject]@{ Pattern = '/groups\?'; Remaining = 1; RetryAfter = 2 })
$run = Invoke-TWRun -Params @{} -KeepWorld
Test-Case 'every request had a fixture (no wrong URL)' { Assert-Equal '' ($global:TWUnrouted -join ' | ') }
Test-Case 'no row has a column that its schema lacks' { Assert-Equal 0 $run.SchemaWarnings.Count (@($run.SchemaWarnings.Keys) -join ', ') }
Test-Case 'no errors, folder and log are named after the tool' {
    Assert-Equal 0 $run.Errors.Count ($run.Errors -join ' | ')
    Assert-Match $run.RunName '^powerbi2csv_\d{8}-\d{6}_11111111'
    Assert-True (Test-Path -LiteralPath $run.LogPath)
}
Test-Case 'throttling (429) is retried after Retry-After' {
    Assert-Equal 1 $run.Stats.Throttled
    Assert-True ($global:TWSleeps -contains 2.0 -or @($global:TWSleeps | Where-Object { $_ -ge 2 -and $_ -lt 3.1 }).Count -gt 0)
}
Test-Case 'workspaces: Power BI and Fabric lists are merged, Fabric paging is followed, My workspace is left out' {
    $ws = Import-TWCsv $run 'PowerBI/workspaces.csv'
    Assert-Equal 'Finance DEV, Sales' (($ws | ForEach-Object { $_.Workspace }) -join ', ')
    $sales = $ws | Where-Object { $_.Workspace -eq 'Sales' }
    Assert-Equal 'cap-prod' $sales.Capacity; Assert-Equal 'F64' $sales.CapacitySku; Assert-Equal 'West Europe' $sales.CapacityRegion
    Assert-Equal '2' $sales.Reports; Assert-Equal '2' $sales.SemanticModels; Assert-Equal '9' $sales.FabricItems; Assert-Equal 'anna@contoso.com' $sales.Admins
    Assert-Equal 'ConnectedAndInitialized' $sales.GitState
    Assert-Equal 'finance sandbox' ($ws | Where-Object { $_.Workspace -eq 'Finance DEV' }).Description
}
Test-Case 'a refused call makes the workspace Partial with the reason, the rest is still collected' {
    $ws = Import-TWCsv $run 'PowerBI/workspaces.csv'
    $sales = $ws | Where-Object { $_.Workspace -eq 'Sales' }
    Assert-Equal 'Partial' $sales.PowerBIStatus
    Assert-Match $sales.PowerBIProblems 'semantic model users: 1 call\(s\) failed \(HTTP 403 Forbidden'
    Assert-Equal 'Collected' $sales.FabricStatus
    Assert-Match ($ws | Where-Object { $_.Workspace -eq 'Finance DEV' }).PowerBIProblems 'workspace users: 1 call\(s\) failed \(HTTP 401 PowerBINotAuthorizedException'
    Assert-Equal 'Partial' (Get-TWCoverage $run 'Power BI workspace Sales')['Status']
}
Test-Case 'reports, pages, paginated report data sources' {
    $r = Import-TWCsv $run 'PowerBI/reports.csv'
    Assert-Equal 2 $r.Count
    Assert-Equal 'Sales model' ($r | Where-Object { $_.ReportId -eq 'r1' }).Model
    Assert-Equal 2 (Import-TWCsv $run 'PowerBI/reportPages.csv').Count
    Assert-Equal 'sql1.contoso.com' (Import-TWCsv $run 'PowerBI/reportDatasources.csv')[0].Server
}
Test-Case 'semantic models: data sources, last refresh, schedule, history with duration' {
    $m = Import-TWCsv $run 'PowerBI/semanticModels.csv'
    $d1 = $m | Where-Object { $_.ModelId -eq 'd1' }
    Assert-Equal 'Failed' $d1.LastRefreshStatus; Assert-Equal 'False' $d1.ScheduleEnabled; Assert-Equal '1' $d1.Reports; Assert-Equal 'Sql' $d1.DatasourceTypes
    Assert-Equal '2025-01-02T03:04:05Z' $d1.CreatedDate
    $h = Import-TWCsv $run 'PowerBI/refreshHistory.csv'
    Assert-Equal 2 $h.Count; Assert-Equal '90' $h[0].DurationSec; Assert-Match $h[0].Error 'CredentialsNotSpecified'
    Assert-Equal 'Monday, Friday' (Import-TWCsv $run 'PowerBI/refreshSchedules.csv')[0].Days
    Assert-Equal 'Sales' (Import-TWCsv $run 'PowerBI/semanticModelDatasources.csv')[0].Database
}
Test-Case 'dashboards, dataflows, lineage, apps, pipelines, features' {
    Assert-Equal 'Revenue' (Import-TWCsv $run 'PowerBI/dashboardTiles.csv')[0].Tile
    Assert-Equal 'Success' (Import-TWCsv $run 'PowerBI/dataflows.csv')[0].LastRunStatus
    $l = Import-TWCsv $run 'PowerBI/lineage.csv'
    Assert-True (@($l | Where-Object { $_.ItemType -eq 'Report' -and $_.DependsOnId -eq 'd1' }).Count -eq 1)
    Assert-True (@($l | Where-Object { $_.ItemType -eq 'SemanticModel' -and $_.DependsOnId -eq 'df1' }).Count -eq 1)
    Assert-Equal 'Sales' (Import-TWCsv $run 'PowerBI/apps.csv')[0].Workspace
    Assert-Equal 2 (Import-TWCsv $run 'PowerBI/pipelineStages.csv').Count
    Assert-Equal 4 (Import-TWCsv $run 'PowerBI/pipelineStageItems.csv').Count
    Assert-Equal 'anna@contoso.com' (Import-TWCsv $run 'PowerBI/pipelineOperations.csv')[0].PerformedBy
    Assert-Equal 'embedTrial' (Import-TWCsv $run 'PowerBI/availableFeatures.csv')[0].Feature
}
Test-Case 'capacities and gateways come from both APIs' {
    $c = Import-TWCsv $run 'PowerBI/capacities.csv'
    Assert-Equal 2 $c.Count
    Assert-Equal 'Power BI; Fabric' ($c | Where-Object { $_.Capacity -eq 'cap-prod' }).Source
    Assert-Equal '1' ($c | Where-Object { $_.Capacity -eq 'cap-prod' }).Workspaces
    Assert-Equal 'Fabric' (Import-TWCsv $run 'PowerBI/gateways.csv')[0].Source
}
Test-Case 'Fabric: items, roles, git, spark, lakehouse tables, shortcuts, schedules, connections' {
    $i = Import-TWCsv $run 'PowerBI/fabricItems.csv'
    Assert-Equal 9 $i.Count
    Assert-Equal '1' ($i | Where-Object { $_.ItemId -eq 'lh1' }).Shortcuts
    Assert-Match ($i | Where-Object { $_.ItemId -eq 'lh1' }).Properties 'sqlEndpointProperties'
    Assert-Equal 'app-123' (Import-TWCsv $run 'PowerBI/workspaceRoleAssignments.csv' | Where-Object { $_.PrincipalType -eq 'ServicePrincipal' }).AppId
    Assert-Equal 'sales' (Import-TWCsv $run 'PowerBI/workspaceGit.csv' | Where-Object { $_.Workspace -eq 'Sales' }).Repository
    Assert-Equal 'Starter Pool' (Import-TWCsv $run 'PowerBI/workspaceSparkSettings.csv')[0].DefaultPool
    Assert-Equal 2 @(Import-TWCsv $run 'PowerBI/lakehouseTables.csv' | Where-Object { $_.Lakehouse -eq 'Bronze' }).Count
    Assert-Equal 'AdlsGen2' (Import-TWCsv $run 'PowerBI/shortcuts.csv')[0].TargetType
    Assert-Equal 'RunNotebook' (Import-TWCsv $run 'PowerBI/itemSchedules.csv')[0].JobType
    Assert-Equal 'sql1.contoso.com;Sales' (Import-TWCsv $run 'PowerBI/itemConnections.csv')[0].Path
    Assert-Equal 'Basic' (Import-TWCsv $run 'PowerBI/connections.csv')[0].CredentialType
}
Test-Case 'SQL endpoints: objects per lakehouse and warehouse, counts, tables of a lakehouse with schemas' {
    $o = Import-TWCsv $run 'PowerBI/sqlObjects.csv'
    Assert-Equal 11 $o.Count
    Assert-Equal 'Function, StoredProcedure, Table, View, View' ((($o | Where-Object { $_.Item -eq 'sales-dw' } | ForEach-Object { $_.ObjectType }) | Sort-Object) -join ', ')
    Assert-Equal '2026-01-02T03:04:05Z' $o[0].Created
    $w = Import-TWCsv $run 'PowerBI/warehouses.csv'
    Assert-Equal 'sales-dw 1 2 1 1 Read' ('{0} {1} {2} {3} {4} {5}' -f $w[0].Warehouse, $w[0].Tables, $w[0].Views, $w[0].StoredProcedures, $w[0].Functions, $w[0].SqlStatus)
    $l = Import-TWCsv $run 'PowerBI/lakehouses.csv'
    $bronze = $l | Where-Object { $_.Lakehouse -eq 'Bronze' }; $gold = $l | Where-Object { $_.Lakehouse -eq 'Gold' }
    Assert-Equal '2 Fabric API 1' ('{0} {1} {2}' -f $bronze.Tables, $bronze.TablesSource, $bronze.Views)
    Assert-Equal '3 SQL endpoint 0' ('{0} {1} {2}' -f $gold.Tables, $gold.TablesSource, $gold.Views)
    $t = @(Import-TWCsv $run 'PowerBI/lakehouseTables.csv' | Where-Object { $_.Lakehouse -eq 'Gold' })
    Assert-Equal 'gold.dim_date, gold.fact_sales, silver.sales' ((($t | ForEach-Object { $_.Schema + '.' + $_.Table }) | Sort-Object) -join ', ')
    Assert-Equal 'SQL endpoint' $t[0].Source
    Assert-Equal 'Collected' (Import-TWCsv $run 'PowerBI/workspaces.csv' | Where-Object { $_.Workspace -eq 'Sales' }).FabricStatus 'a lakehouse with schemas is not a problem'
    Assert-Equal 3 @($global:TWRequests | Where-Object { $_.Kind -eq 'sql' }).Count
}
Test-Case 'SQL endpoints: columns, code of views and procedures, references between databases; a refused catalog query is reported' {
    $c = Import-TWCsv $run 'PowerBI/sqlColumns.csv'
    Assert-Equal 'Amount decimal 18 2 YES' ('{0} {1} {2} {3} {4}' -f $c[0].Column, $c[0].DataType, $c[0].Precision, $c[0].Scale, $c[0].Nullable)
    $m = Import-TWCsv $run 'PowerBI/sqlModules.csv'
    $v = $m | Where-Object { $_.Object -eq 'vSales' }
    Assert-Equal 'View Bronze' ($v.ObjectType + ' ' + $v.ReferencedDatabases); Assert-Match $v.Definition 'CREATE VIEW dbo.vSales AS SELECT'
    $r = Import-TWCsv $run 'PowerBI/sqlReferences.csv'
    Assert-Equal 'Bronze dbo orders Lakehouse Sales' ('{0} {1} {2} {3} {4}' -f $r[0].ReferencedDatabase, $r[0].ReferencedSchema, $r[0].ReferencedObject, $r[0].ReferencedItemType, $r[0].ReferencedWorkspace)
    $o = Import-TWCsv $run 'PowerBI/sqlObjects.csv' | Where-Object { $_.Object -eq 'FactSales' }
    Assert-Equal '2' $o.Columns
    $cov = Get-TWCoverage $run 'Fabric SQL endpoints'
    Assert-Equal 'Partial' $cov['Status']; Assert-Match $cov['Detail'] 'Gold modules: The SELECT permission was denied'
}
Test-Case 'run history: runs per notebook and pipeline, last run on the item' {
    $r = Import-TWCsv $run 'PowerBI/itemRuns.csv'
    Assert-Equal 2 $r.Count
    $f = $r | Where-Object { $_.Status -eq 'Failed' }
    Assert-Equal '60' $f.DurationSec; Assert-Match $f.FailureReason 'Table not found'
    $i = Import-TWCsv $run 'PowerBI/fabricItems.csv'
    $nb = $i | Where-Object { $_.ItemId -eq 'nb1' }
    Assert-Equal '2 Failed 2026-09-28T02:00:00Z' ('{0} {1} {2}' -f $nb.Runs, $nb.LastRunStatus, $nb.LastRunStart)
    Assert-Equal '0' ($i | Where-Object { $_.Item -eq 'PL_Daily' }).Runs
    Assert-Equal '' ($i | Where-Object { $_.ItemId -eq 'lh1' }).Runs
}
Test-Case 'folders, tags, network policy, Git item status, Spark sessions, OneLake roles, deployment pipelines, domains' {
    $f = Import-TWCsv $run 'PowerBI/folders.csv'
    Assert-Equal 'ETL 0|ETL/Bronze 1' (($f | ForEach-Object { $_.Path + ' ' + $_.Items }) -join '|')
    $nb = Import-TWCsv $run 'PowerBI/fabricItems.csv' | Where-Object { $_.ItemId -eq 'nb1' }
    Assert-Equal 'ETL/Bronze l1 Certified source' ('{0} {1} {2}' -f $nb.FolderPath, $nb.LogicalId, $nb.Tags)
    Assert-Match (Import-TWCsv $run 'PowerBI/fabricItems.csv' | Where-Object { $_.Item -eq 'Import files' }).Properties 'isParametric'
    $n = Import-TWCsv $run 'PowerBI/workspaceNetwork.csv'
    Assert-Equal 'Sales Allow Deny' ('{0} {1} {2}' -f $n[0].Workspace, $n[0].InboundDefaultAction, $n[0].OutboundDefaultAction); Assert-Equal 1 @($n).Count
    $g = Import-TWCsv $run 'PowerBI/workspaceGit.csv' | Where-Object { $_.Workspace -eq 'Sales' }
    Assert-Equal 'Read 1 1 0 abc999' ('{0} {1} {2} {3} {4}' -f $g.ItemStatus, $g.UncommittedItems, $g.IncomingItems, $g.Conflicts, $g.RemoteCommitHash)
    Assert-Equal 'Load bronze Modified|New report ' ((Import-TWCsv $run 'PowerBI/gitItemStatus.csv' | ForEach-Object { $_.Item + ' ' + $_.WorkspaceChange }) -join '|')
    $s = Import-TWCsv $run 'PowerBI/sparkSessions.csv'
    Assert-Equal 'Notebook Load bronze Succeeded 120' ('{0} {1} {2} {3}' -f $s[0].ItemType, $s[0].Item, $s[0].State, $s[0].TotalDurationSec)
    $r = Import-TWCsv $run 'PowerBI/onelakeRoles.csv'
    Assert-Equal 'Bronze SalesReaders /Tables/customers Read grp-1' ('{0} {1} {2} {3} {4}' -f $r[0].Lakehouse, $r[0].Role, $r[0].Paths, $r[0].Actions, $r[0].EntraMembers)
    $st = Import-TWCsv $run 'PowerBI/deploymentStages.csv'
    Assert-Equal 'Development  |Production Sales 2' (($st | ForEach-Object { '{0} {1} {2}' -f $_.Stage, $_.Workspace, $_.Items }) -join '|')
    Assert-Equal 'Lakehouse, Notebook' (((Import-TWCsv $run 'PowerBI/deploymentStageItems.csv') | ForEach-Object { $_.ItemType } | Sort-Object) -join ', ')
    Assert-Equal 'Admin' (Import-TWCsv $run 'PowerBI/deploymentRoleAssignments.csv')[0].Role
    Assert-Equal 'Sales domain' (Import-TWCsv $run 'PowerBI/domains.csv')[0].Domain
}
Test-Case 'shortcut target as columns' {
    $s = (Import-TWCsv $run 'PowerBI/shortcuts.csv')[0]
    Assert-Equal 'https://stor.dfs.core.windows.net /raw' ($s.TargetLocation + ' ' + $s.TargetSubpath)
}
Test-Case 'pipeline activities: nested ones, referenced items, procedure and parameters, items out of scope' {
    $a = Import-TWCsv $run 'PowerBI/pipelineActivities.csv'
    Assert-Equal 'Run transform, Load model, Refresh model, On failure, Mail' (($a | ForEach-Object { $_.Activity }) -join ', ')
    $n = $a | Where-Object { $_.Activity -eq 'Run transform' }
    Assert-Equal 'Notebook Transform Sales' ('{0} {1} {2}' -f $n.RefType, $n.RefItem, $n.RefWorkspace)
    $p = $a | Where-Object { $_.Activity -eq 'Load model' }
    Assert-Equal '[dm].[load_data_model] eo_start=1000 Run transform' ('{0} {1} {2}' -f $p.StoredProcedure, $p.Parameters, $p.DependsOn)
    Assert-Match ($a | Where-Object { $_.Activity -eq 'Refresh model' }).References '^SemanticModel: dddddddd-0000-0000-0000-00000000000d \(not in scope\)'
    Assert-Equal 'On failure' ($a | Where-Object { $_.Activity -eq 'Mail' }).Parent
    Assert-NotMatch ($a | Where-Object { $_.Activity -eq 'On failure' }).TypeProperties 'Office365Email'
}
Test-Case 'notebook dependencies: lakehouses from the metadata, tables, paths, calls and URLs from the code, comments skipped' {
    $d = @(Import-TWCsv $run 'PowerBI/notebookDependencies.csv' | Where-Object { $_.Notebook -eq 'Transform' })
    $got = ($d | ForEach-Object { $_.Relation + '=' + $_.Target } | Sort-Object) -join '; '
    Assert-Equal ('AttachedLakehouse=Bronze; DefaultLakehouse=Gold; OneLakePath=aaaaaaaa-0000-0000-0000-000000000001/lh1/Tables/; ReadsTable=gold.dim_date; ReadsTable=silver.sales; ' +
        'RunsNotebook=Load bronze; Url=https://api.contoso.com/rates; WritesTable=gold.fact_sales') $got
    Assert-Equal 'Bronze' ($d | Where-Object { $_.Relation -eq 'OneLakePath' }).TargetItem
    Assert-Equal 'nb1' ($d | Where-Object { $_.Relation -eq 'RunsNotebook' }).TargetItemId
}
Test-Case 'dataflow sources: file and site, Fabric items by SQL endpoint, destinations, gen1' {
    $s = Import-TWCsv $run 'PowerBI/dataflowSources.csv'
    $b = $s | Where-Object { $_.Query -eq 'Budget' }
    Assert-Equal 'SharePoint.Files https://contoso.sharepoint.com/sites/Finance Budget.xlsx SharePoint External' ('{0} {1} {2} {3} {4}' -f $b.Connector, $b.Target, $b.File, $b.Kind, $b.Scope)
    Assert-Equal 1 @($b).Count 'Excel.Workbook around the file is not a second source'
    $o = $s | Where-Object { $_.Query -eq 'Orders' }
    Assert-Equal 'Warehouse sales-dw Fabric (this tenant)' ('{0} {1} {2}' -f $o.TargetItemType, $o.TargetItem, $o.Scope)
    Assert-Equal 'Destination' ($s | Where-Object { $_.Query -eq 'Budget_DataDestination' }).Role
    $g1 = $s | Where-Object { $_.Generation -eq 'Gen1' }
    Assert-Equal 'Customers Web.Contents https://api.contoso.com/customers' ('{0} {1} {2}' -f $g1.Query, $g1.Connector, $g1.Target)
}
Test-Case 'semantic model definition: partitions and modes, DAX of measures and calculated columns, M, roles, sources' {
    $p = Import-TWCsv $run 'PowerBI/modelPartitions.csv'
    Assert-Equal 'Date entity directLake | Sales m import' ((($p | ForEach-Object { '{0} {1} {2}' -f $_.Table, $_.SourceType, $_.Mode }) | Sort-Object) -join ' | ')
    Assert-Match ($p | Where-Object { $_.Table -eq 'Sales' }).Source 'Source = Reference'
    Assert-Match ($p | Where-Object { $_.Table -eq 'Date' }).Source 'entityName: dim_date'
    $e = Import-TWCsv $run 'PowerBI/modelExpressions.csv'
    Assert-Equal 'SUM(Sales[Amount]) #,0 KPI' ('{0} {1} {2}' -f ($e | Where-Object { $_.Name -eq 'Total Sales' }).Expression, ($e | Where-Object { $_.Name -eq 'Total Sales' }).FormatString, ($e | Where-Object { $_.Name -eq 'Total Sales' }).DisplayFolder)
    Assert-Equal 'VAR s = [Total Sales] RETURN s * 0.2' ($e | Where-Object { $_.Name -eq 'Margin' }).Expression
    Assert-Equal 'CalculatedColumn' ($e | Where-Object { $_.Name -eq 'Double' }).Kind
    Assert-Equal 0 @($e | Where-Object { $_.Name -eq 'Amount' }).Count 'a plain column is not an expression'
    Assert-Match ($e | Where-Object { $_.Kind -eq 'Expression (M)' }).Expression 'Sql\.Databases'
    $r = Import-TWCsv $run 'PowerBI/modelRoles.csv'
    Assert-Equal 'Sales [Email] = USERPRINCIPALNAME()' (($r | Where-Object { $_.Table }) | ForEach-Object { $_.Table + ' ' + $_.FilterExpression })
    $s = Import-TWCsv $run 'PowerBI/modelSources.csv'
    Assert-Equal 'Reference Sql.Databases Warehouse sales-dw' ('{0} {1} {2} {3}' -f $s[0].Query, $s[0].Connector, $s[0].TargetItemType, $s[0].TargetItem)
    $m = Import-TWCsv $run 'PowerBI/semanticModels.csv' | Where-Object { $_.ModelId -eq 'd1' }
    Assert-Equal 'directLake (1); import (1) 2 1' ('{0} {1} {2}' -f $m.StorageModes, $m.Measures, $m.RlsRoles)
}
Test-Case 'item dependencies and external targets in one place' {
    $d = Import-TWCsv $run 'PowerBI/itemDependencies.csv'
    $got = ($d | ForEach-Object { '{0} {1} {2} {3} {4}' -f $_.ItemType, $_.Item, $_.Relation, $_.DependsOnType, $_.DependsOn })
    foreach ($x in @('DataPipeline PL_Daily Runs Notebook Transform', 'Notebook Transform Uses Lakehouse Gold', 'Notebook Transform Runs Notebook Load bronze', 'Dataflow Import files ReadsFrom Warehouse sales-dw',
            'SemanticModel Sales model ReadsFrom Warehouse sales-dw', 'Report Sales overview Uses SemanticModel Sales model', 'SemanticModel Sales model ReadsFrom Dataflow Customers')) {
        Assert-True ($got -contains $x) ('missing: ' + $x + ' in ' + ($got -join ' | '))
    }
    $e = Import-TWCsv $run 'PowerBI/externalDependencies.csv'
    $kinds = (($e | ForEach-Object { $_.Kind + ' ' + $_.Target }) | Sort-Object -Unique) -join ' | '
    foreach ($x in @('SharePoint https://contoso.sharepoint.com/sites/Finance', 'Web https://api.contoso.com/rates', 'Web https://api.contoso.com/customers', 'Azure Storage https://stor.dfs.core.windows.net', 'SQL sql1.contoso.com')) {
        Assert-Match $kinds ([regex]::Escape($x))
    }
    Assert-Equal 'Budget.xlsx' (($e | Where-Object { $_.Kind -eq 'SharePoint' })[0].Detail -split ' \| ')[-1]
    $f13 = Import-TWCsv $run 'findings.csv' | Where-Object { $_.Id -eq 'PBI-13' }
    Assert-True ([int]$f13.Count -ge 5) 'one per distinct target'; Assert-Match $f13.Detail 'Web \(\d\)'
}
Test-Case 'admin report is skipped with the reason and writes no admin files' {
    Assert-Equal 'Skipped' (Get-TWCoverage $run 'Power BI admin (tenant-wide)')['Status']
    Assert-Match (Get-TWCoverage $run 'Power BI admin (tenant-wide)')['Detail'] 'Fabric administrator'
    Assert-False (Test-Path -LiteralPath (Join-Path $run.OutDir 'PowerBI\adminWorkspaces.csv'))
}
Test-Case 'findings' {
    $f = Import-TWCsv $run 'findings.csv'
    $ids = ($f | ForEach-Object { $_.Id } | Sort-Object) -join ' '
    Assert-Equal 'PBI-01 PBI-02 PBI-03 PBI-04 PBI-05 PBI-06 PBI-10 PBI-12 PBI-13' $ids
    Assert-Match ($f | Where-Object { $_.Id -eq 'PBI-01' }).Detail 'Sales / Sales model'
    Assert-Match ($f | Where-Object { $_.Id -eq 'PBI-05' }).Detail 'gone@contoso.com'
    Assert-Match ($f | Where-Object { $_.Id -eq 'PBI-06' }).Detail 'Finance DEV'
}
Test-Case 'run.json is valid and counts the API calls' {
    $j = Get-Content -LiteralPath (Join-Path $run.OutDir ($run.RunStamp + '__run.json')) -Raw | ConvertFrom-Json
    Assert-Equal 'powerbi2csv' $j.tool; Assert-Equal 2 $j.workspaces.inScope; Assert-Equal 3 $j.workspaces.visible
    Assert-True ($j.apiCalls.pbi -gt 20 -and $j.apiCalls.fabric -gt 10)
}
Test-Case 'CSV files have a UTF-8 BOM and quote every field' {
    $b = [IO.File]::ReadAllBytes((Join-Path $run.OutDir 'PowerBI\reports.csv'))
    Assert-Equal '239 187 191' ($b[0..2] -join ' ')
    Assert-Match ([IO.File]::ReadAllLines((Join-Path $run.OutDir 'PowerBI\reports.csv'))[0]) '^"Workspace","WorkspaceId","Report"'
}

Write-Host 'Scope' -ForegroundColor Cyan
Test-Case '-Workspace with a wildcard, -ExcludeWorkspace, -PbiIncludePersonalWorkspaces' {
    $r1 = Invoke-TWRun -Params @{ Reports = 'Fabric'; Workspace = @('Sal*') }
    Assert-Equal 'Sales' ((Import-TWCsv $r1 'PowerBI/workspaces.csv' | ForEach-Object { $_.Workspace }) -join ', ')
    $r2 = Invoke-TWRun -Params @{ Reports = 'Fabric'; ExcludeWorkspace = @($global:TWW1) }
    Assert-Equal 'Finance DEV' ((Import-TWCsv $r2 'PowerBI/workspaces.csv' | ForEach-Object { $_.Workspace }) -join ', ')
    $r3 = Invoke-TWRun -Params @{ PbiIncludePersonalWorkspaces = $true; SkipReports = 'PowerBIAdmin' }
    Assert-Equal 3 (Import-TWCsv $r3 'PowerBI/workspaces.csv').Count
    Assert-Equal '' ($global:TWUnrouted -join ' | ')
    Assert-True (@($global:TWRequests | Where-Object { $_.Uri -eq 'https://api.powerbi.com/v1.0/myorg/reports' }).Count -eq 1) 'My workspace is read without /groups'
}
Test-Case 'Quick profile (Fabric only) still writes workspaces.csv and makes no Power BI item calls' {
    $r = Invoke-TWRun -Params @{ ReportProfile = 'Quick' }
    Assert-Equal 2 (Import-TWCsv $r 'PowerBI/workspaces.csv').Count
    Assert-Equal 0 @($global:TWRequests | Where-Object { $_.Uri -match '/datasets' }).Count
    Assert-False (Test-Path -LiteralPath (Join-Path $r.OutDir 'PowerBI\reports.csv'))
}

Test-Case 'SQL endpoint not reachable: the run goes on, the files say why the counts are missing' {
    Reset-TWState
    foreach ($k in @($global:TW.Sql.Keys)) { $global:TW.Sql[$k] = 'A network-related or instance-specific error occurred' }
    $r = Invoke-TWRun -Params @{ Reports = 'Fabric,FabricSql' } -KeepWorld
    Assert-Equal 0 $r.Errors.Count
    Assert-Equal 'Skipped' (Get-TWCoverage $r 'Fabric SQL endpoints')['Status']
    $w = Import-TWCsv $r 'PowerBI/warehouses.csv'
    Assert-Equal '' $w[0].Views; Assert-Match $w[0].SqlStatus '^Not read: A network-related'
    Assert-Equal '' (Import-TWCsv $r 'PowerBI/lakehouses.csv' | Where-Object { $_.Lakehouse -eq 'Gold' }).Tables
    $q = Invoke-TWRun -Params @{ Reports = 'Fabric' }
    Assert-Match (Import-TWCsv $q 'PowerBI/warehouses.csv')[0].SqlStatus 'Not selected'
    Assert-Equal 0 @($global:TWRequests | Where-Object { $_.Kind -eq 'sql' }).Count
}

Test-Case 'environment matrix: same solution in DEV and PROD, what is missing where' {
    Reset-TWState
    $global:TW.PbiRoutes.Insert(0, (New-TWRoute GET '^/groups\?' { New-TWList @(@{ id = $global:TWW1; name = 'Sales'; type = 'Workspace'; capacityId = $global:TWCap }, @{ id = $global:TWW2; name = 'Sales - DEV'; type = 'Workspace' }) }))
    $global:TW.FabricRoutes.Insert(0, (New-TWRoute GET "^/workspaces/$($global:TWW2)/items$" { New-TWList @(@{ id = 'x1'; displayName = 'Transform_dev'; type = 'Notebook' }, @{ id = 'x2'; displayName = 'Experiment'; type = 'Notebook' }) }))
    $global:TW.FabricRoutes.Insert(0, (New-TWRoute GET "^/workspaces/$($global:TWW2)/items/x\d/(connections|jobs/instances|jobs/\w+/schedules)$" { New-TWList @() }))
    $global:TW.FabricRoutes.Insert(0, (New-TWRoute GET '^/workspaces\?continuationToken=' { @{ value = @(@{ id = $global:TWW2; displayName = 'Sales - DEV'; type = 'Workspace' }) } }))
    $r = Invoke-TWRun -Params @{ Reports = 'Fabric' } -KeepWorld
    Assert-Equal '' ($global:TWUnrouted -join ' | ')
    $m = Import-TWCsv $r 'PowerBI/environmentMatrix.csv'
    $t = $m | Where-Object { $_.Item -eq 'Transform' }
    Assert-Equal 'Sales Notebook Transform_dev Transform 2' ('{0} {1} {2} {3} {4}' -f $t.Solution, $t.ItemType, $t.DEV, $t.PROD, $t.Environments)
    Assert-Equal 'PROD' ($m | Where-Object { $_.Item -eq 'Experiment' }).MissingIn
    Assert-Equal 'DEV' ($m | Where-Object { $_.Item -eq 'PL_Daily' }).MissingIn
    Assert-True (@(Import-TWCsv $r 'findings.csv' | Where-Object { $_.Id -eq 'PBI-14' }).Count -eq 1)
}
Test-Case '-WorkspaceListFile (.csv and .xlsx): found, spelled differently, not visible, not on the list; columns carried over' {
    $csv = Join-Path $global:TWOut 'list.csv'
    Set-Content -LiteralPath $csv -Encoding UTF8 -Value @('Workspace,To migrate,owner', 'Sales,Y,Anna', 'Finanse  DEV,Y,Bob', 'Marketing,N,Eve')
    $r = Invoke-TWRun -Params @{ Reports = 'Fabric'; Workspace = @('Sales'); WorkspaceListFile = $csv }
    $p = Import-TWCsv $r 'PowerBI/workspacePlan.csv'
    Assert-Equal 'Found|Found (spelled differently)|Not visible to this account' (($p | Where-Object { $_.ListedWorkspace } | ForEach-Object { $_.Status }) -join '|')
    $f = $p | Where-Object { $_.ListedWorkspace -eq 'Finanse  DEV' }
    Assert-Equal 'Finance DEV False Bob Y' ('{0} {1} {2} {3}' -f $f.Workspace, $f.InScope, $f.Listed_owner, $f.Listed_To_migrate)
    Assert-Equal 0 @($p | Where-Object { $_.Status -eq 'Not on the list' }).Count
    Assert-Match (Import-TWCsv $r 'findings.csv' | Where-Object { $_.Id -eq 'PBI-11' }).Detail 'Marketing'
    # the same list as a workbook: shared strings, a second sheet named Workspaces
    $xlsx = Join-Path $global:TWOut 'list.xlsx'
    Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
    $z = [IO.Compression.ZipFile]::Open($xlsx, [IO.Compression.ZipArchiveMode]::Create)
    $entries = [ordered]@{
        'xl/workbook.xml' = '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Intro" sheetId="1" r:id="rId1"/><sheet name="Workspaces" sheetId="2" r:id="rId2"/></sheets></workbook>'
        'xl/_rels/workbook.xml.rels' = '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="t" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="t" Target="worksheets/sheet2.xml"/></Relationships>'
        'xl/sharedStrings.xml' = '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>Workspace</t></si><si><t>owner</t></si><si><t>Sales</t></si><si><r><t>An</t></r><r><t>na</t></r></si><si><t>Marketing</t></si></sst>'
        'xl/worksheets/sheet1.xml' = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>readme</t></is></c></row></sheetData></worksheet>'
        'xl/worksheets/sheet2.xml' = '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="C1" t="s"><v>1</v></c></row><row r="2"><c r="A2" t="s"><v>2</v></c><c r="C2" t="s"><v>3</v></c></row><row r="3"><c r="A3" t="s"><v>4</v></c></row></sheetData></worksheet>'
    }
    try { foreach ($k in $entries.Keys) { $sw = New-Object IO.StreamWriter($z.CreateEntry($k).Open()); try { $sw.Write($entries[$k]) } finally { $sw.Dispose() } } } finally { $z.Dispose() }
    $t = Read-A2CTable -Path $xlsx -Sheet 'Workspaces'
    Assert-Equal 'Sales/Anna|Marketing/' (($t | ForEach-Object { $_['Workspace'] + '/' + $_['owner'] }) -join '|')
    Assert-Throws { New-A2COptions -Params @{ WorkspaceListFile = (Join-Path $global:TWOut 'nope.csv') } } 'not found'
}
Test-Case 'OneLake sizes (paging, schemas, a refused item) and Power Automate flows' {
    $r = Invoke-TWRun -Params @{ Reports = 'FabricStorage,PowerAutomate' }
    Assert-Equal '' ($global:TWUnrouted -join ' | ')
    $t = Import-TWCsv $r 'PowerBI/onelakeTableSizes.csv'
    $c = $t | Where-Object { $_.Table -eq 'customers' }; $o = $t | Where-Object { $_.Table -eq 'orders' }; $g = $t | Where-Object { $_.Table -eq 'fact_sales' }
    Assert-Equal '2 1048676 1' ('{0} {1} {2}' -f $c.Files, $c.Bytes, $c.MB)
    Assert-Equal '2 2000 2026-09-02T10:00:00Z' ('{0} {1} {2}' -f $o.Files, $o.Bytes, $o.LastModified)
    Assert-Equal 'gold 3000' ($g.Schema + ' ' + $g.Bytes)
    $s = Import-TWCsv $r 'PowerBI/onelakeStorage.csv'
    $lh1 = $s | Where-Object { $_.ItemId -eq 'lh1' }
    Assert-Equal '2 4 1050676 1 2000 Read' ('{0} {1} {2} {3} {4} {5}' -f $lh1.Tables, $lh1.TableFiles, $lh1.TableBytes, $lh1.FilesAreaFiles, $lh1.FilesAreaBytes, $lh1.Status)
    Assert-Match ($s | Where-Object { $_.ItemId -eq 'dw1' }).Status '^Not read: HTTP 403'
    Assert-Equal 'Partial' (Get-TWCoverage $r 'OneLake storage')['Status']
    $f = Import-TWCsv $r 'PowerBI/flows.csv'
    Assert-Equal 'Refresh sales model Started Owned Recurrence True' ('{0} {1} {2} {3} {4}' -f $f[0].Flow, $f[0].State, $f[0].Scope, $f[0].Triggers, $f[0].UsesPowerBI)
    Assert-Equal '1' (Import-TWCsv $r 'PowerBI/flowEnvironments.csv')[0].Flows
    Assert-Match (Get-TWCoverage $r 'Power Automate')['Detail'] 'Power Platform administrator'
}
Test-Case 'long running operations are started together and polled together' {
    Reset-TWState
    $res = Invoke-A2CFabricLroBatch -Paths @("/workspaces/$($global:TWW1)/items/nb1/getDefinition", "/workspaces/$($global:TWW1)/items/lh1/getDefinition", "/workspaces/$($global:TWW1)/items/$($global:TWNb2)/getDefinition")
    Assert-Equal 'True False True' (($res | ForEach-Object { $_.Ok }) -join ' ')
    Assert-Equal 'OperationNotSupportedForItem' $res[1].Error.Code
    Assert-Equal 'POST POST POST GET GET' (($global:TWRequests | ForEach-Object { $_.Method }) -join ' ')
}

Write-Host 'Full run as a Fabric administrator' -ForegroundColor Cyan
Reset-TWState
$global:TW.Admin = $true
$run = Invoke-TWRun -Params @{ ReportProfile = 'Full'; PbiActivityDays = 2 } -KeepWorld
Test-Case 'no unrouted requests, no schema warnings, no errors' {
    Assert-Equal '' ($global:TWUnrouted -join ' | ')
    Assert-Equal 0 $run.SchemaWarnings.Count (@($run.SchemaWarnings.Keys) -join ', ')
    Assert-Equal 0 $run.Errors.Count ($run.Errors -join ' | ')
}
Test-Case 'admin: all workspaces of the tenant with users' {
    $w = Import-TWCsv $run 'PowerBI/adminWorkspaces.csv'
    Assert-Equal 5 $w.Count
    Assert-Equal '0' ($w | Where-Object { $_.Workspace -eq 'Orphaned' }).Users
    Assert-Equal '3' ($w | Where-Object { $_.Workspace -eq 'Sales' }).Items
    Assert-Equal 2 @(Import-TWCsv $run 'PowerBI/adminWorkspaceUsers.csv').Count
}
Test-Case 'admin: Scanner API result (items of every type, access, data sources, expressions, raw JSON)' {
    $i = Import-TWCsv $run 'PowerBI/adminItems.csv'
    Assert-Equal 'datasets, Lakehouse, reports' ((($i | ForEach-Object { $_.ItemType }) | Sort-Object -Unique) -join ', ')
    Assert-Equal 'dsi1' ($i | Where-Object { $_.ItemType -eq 'datasets' -and $_.Workspace -eq 'Sales' }).DatasourceIds
    $u = Import-TWCsv $run 'PowerBI/adminItemUsers.csv'
    Assert-Equal 'Read' $u[0].AccessRight
    Assert-Equal 'sql1.contoso.com' ((Import-TWCsv $run 'PowerBI/adminDatasourceInstances.csv')[0].ConnectionDetails | ConvertFrom-Json).server
    Assert-Match (Import-TWCsv $run 'PowerBI/adminModelTables.csv')[0].Source 'Sql\.Database'
    Assert-Equal 'Expression, Measure' (((Import-TWCsv $run 'PowerBI/adminModelExpressions.csv') | ForEach-Object { $_.Kind }) -join ', ')
    Assert-True (Test-Path -LiteralPath (Join-Path $run.OutDir 'PowerBI\scan\scanResult_001.json'))
    $post = @($global:TWRequests | Where-Object { $_.Method -eq 'POST' -and $_.Uri -match 'getInfo' })
    Assert-Equal 1 $post.Count; Assert-Match $post[0].Body ('^\{"workspaces":\["' + $global:TWW1)
}
Test-Case 'admin: tenant settings, domains, widely shared' {
    $s = Import-TWCsv $run 'PowerBI/tenantSettings.csv'
    Assert-Equal 2 $s.Count; Assert-Equal 'Marketing' ($s | Where-Object { $_.SettingName -eq 'PublishToWeb' }).EnabledSecurityGroups
    Assert-Equal 'Sales domain' (Import-TWCsv $run 'PowerBI/adminDomains.csv')[0].Domain
    Assert-Equal 'Published to web' (Import-TWCsv $run 'PowerBI/widelyShared.csv')[0].Kind
    Assert-True (@(Import-TWCsv $run 'findings.csv' | Where-Object { $_.Id -eq 'PBI-08' -and $_.Severity -eq 'High' }).Count -eq 1)
}
Test-Case 'activity log: one request per UTC day, continuation followed' {
    $e = Import-TWCsv $run 'PowerBI/activityEvents.csv'
    Assert-Equal 'e1 e2 e3' (($e | ForEach-Object { $_.Id }) -join ' ')
    Assert-Equal 'Sales' $e[0].Workspace
    $days = @($global:TWRequests | Where-Object { $_.Uri -match 'activityevents\?startDateTime' })
    Assert-Equal 2 $days.Count
    Assert-Match $days[1].Uri "startDateTime='2026-09-29T00:00:00.000Z'&endDateTime='2026-09-29T23:59:59.999Z'"
}
Test-Case 'models: DAX INFO.VIEW results (BOM tolerated), a refused model makes the dataset Partial' {
    Assert-Equal 'Date, Sales' (((Import-TWCsv $run 'PowerBI/modelTables.csv') | ForEach-Object { $_.Table } | Sort-Object) -join ', ')
    $m = Import-TWCsv $run 'PowerBI/modelMeasures.csv'
    Assert-Equal 'Total Sales' $m[0].Measure; Assert-Equal 'SUM(Sales[Amount])' $m[0].Expression
    Assert-Equal 'Sales' (Import-TWCsv $run 'PowerBI/modelRelationships.csv')[0].FromTable
    Assert-Equal 'Partial' $run.DatasetStatus['PowerBIModels']
    Assert-Equal 0 @(Get-ChildItem -LiteralPath $run.OutDir -Recurse -Filter *.pbix).Count 'report files are not part of Full'
    Assert-True (Test-Path -LiteralPath (Join-Path $run.OutDir 'PowerBI\Sales\Notebook\Transform\notebook-content.py')) 'the code download is'
}

Write-Host 'Code download' -ForegroundColor Cyan
Test-Case 'FabricCode: notebooks, pipelines, dataflows, models and SQL as files, no report files; -PbiNotebookFormat Ipynb' {
    $r = Invoke-TWRun -Params @{ Reports = 'Fabric,FabricSql,FabricCode' }
    Assert-Equal '' ($global:TWUnrouted -join ' | ')
    $base = Join-Path $r.OutDir 'PowerBI\Sales'
    Assert-Match ([IO.File]::ReadAllText((Join-Path $base 'Notebook\Transform\notebook-content.py'))) 'saveAsTable'
    Assert-True (Test-Path -LiteralPath (Join-Path $base 'DataPipeline\PL_Daily\pipeline-content.json'))
    Assert-True (Test-Path -LiteralPath (Join-Path $base 'Dataflow\Import files\mashup.pq'))
    Assert-True (Test-Path -LiteralPath (Join-Path $base 'SemanticModel\Sales model\definition\tables\Sales.tmdl'))
    Assert-True (Test-Path -LiteralPath (Join-Path $base 'dataflows\Customers.json'))
    Assert-True (Test-Path -LiteralPath (Join-Path $base 'sql\sales-dw\dm.load_data_model.sql'))
    Assert-Equal 0 @(Get-ChildItem -LiteralPath $r.OutDir -Recurse -Filter *.pbix).Count
    Assert-False (Test-Path -LiteralPath (Join-Path $base 'Report'))
    $x = Import-TWCsv $r 'PowerBI/exportIndex.csv'
    Assert-Equal 0 @($x | Where-Object { $_.ItemType -eq 'Report' }).Count
    Assert-Equal 'PowerBI/Sales/Notebook/Transform' (Import-TWCsv $r 'PowerBI/fabricItems.csv' | Where-Object { $_.Item -eq 'Transform' }).ExportedTo
    Reset-TWState
    $r2 = Invoke-TWRun -Params @{ Reports = 'FabricCode'; PbiNotebookFormat = 'Ipynb' } -KeepWorld
    Assert-True (@($global:TWRequests | Where-Object { $_.Uri -match 'nb1/getDefinition\?format=ipynb$' }).Count -eq 1)
    Assert-True (@($global:TWRequests | Where-Object { $_.Uri -match "$($global:TWPl1)/getDefinition$" }).Count -eq 1) 'only notebooks get the format'
}

Test-Case 'admin scope: per-workspace scope, content, last change, admins, data source types and gateway' {
    $w = Import-TWCsv $run 'PowerBI/adminWorkspaces.csv'
    $got = ($w | Sort-Object Workspace | ForEach-Object { '{0}={1}' -f $_.Workspace, $_.Scope }) -join ' | '
    Assert-Equal 'Legacy reports=Shared, with content | Old project=Not active (Deleted) | Orphaned=Shared, empty | PersonalWorkspace Bob=Personal, empty | Sales=Shared, with content' $got
    $s = $w | Where-Object { $_.Workspace -eq 'Sales' }
    Assert-Equal 'anna@contoso.com False True True 1 Sql True 1' ('{0} {1} {2} {3} {4} {5} {6} {7}' -f $s.Admins, $s.NoAdmin, $s.HasContent, $s.ChangedLast30Days, $s.RecentActivities, $s.DatasourceTypes, $s.UsesGateway, $s.Datasources)
    $l = $w | Where-Object { $_.Workspace -eq 'Legacy reports' }
    Assert-Equal 'True False 2024-03-01T00:00:00Z SharePointList False 0' ('{0} {1} {2} {3} {4} {5}' -f $l.NoAdmin, $l.ChangedLast30Days, $l.LastItemChange, $l.DatasourceTypes, $l.UsesGateway, $l.RecentActivities)
    Assert-Equal '' ($w | Where-Object { $_.IsPersonal -eq 'True' }).NoAdmin 'a personal workspace has an owner, not an admin'
    Assert-Equal '1' (Import-TWCsv $run 'PowerBI/adminDatasourceInstances.csv' | Where-Object { $_.DatasourceId -eq 'dsi2' }).Workspaces
}
Test-Case 'workspaceScopeSummary.csv and findings PBI-17 / PBI-18' {
    $sum = @{}; foreach ($r in (Import-TWCsv $run 'PowerBI/workspaceScopeSummary.csv')) { $sum[$r.Measure] = $r.Count }
    $want = [ordered]@{ 'Workspaces in the tenant' = '5'; 'Not active' = '1'; 'Personal ("My workspace")' = '1'; 'Personal, empty' = '1'; 'Shared' = '3'; 'Shared, empty' = '1'
        'Shared with content' = '2'; 'Shared with content, on a capacity' = '1'; 'Shared with content, using a gateway' = '1'; 'Shared with content, without an admin' = '1'
        'Shared with content, changed in the last 30 days' = '1'; 'Shared with content, with activity in the activity log' = '1'; 'Shared, not scanned' = '0'
        'Shared with content, connected to Git' = '1'; 'Capacities' = '1' }
    foreach ($k in $want.Keys) { Assert-Equal $want[$k] $sum[$k] $k }
    $f = Import-TWCsv $run 'findings.csv'
    Assert-Match ($f | Where-Object { $_.Id -eq 'PBI-17' }).Detail 'Legacy reports'
    $p18 = $f | Where-Object { $_.Id -eq 'PBI-18' }
    Assert-Equal '2' $p18.Count; Assert-Equal '5 workspaces: 1 personal (1 empty), 3 shared (1 empty), 1 not active' $p18.Detail
}

Write-Host 'Share-safe scope run (-ClientScope)' -ForegroundColor Cyan
Test-Case '-ClientScope: admin APIs only, the share-safe files only, nothing sensitive in the output' {
    $csv = Join-Path $global:TWOut 'scope.csv'
    Set-Content -LiteralPath $csv -Encoding UTF8 -Value @('Workspace,owner', 'Sales,Anna', 'Orphaned,Eve', 'Marketing,Max', 'PersonalWorkspace Bob,Bob')
    Reset-TWState
    $global:TW.Admin = $true
    $r = Invoke-TWRun -Params @{ ClientScope = $true; WorkspaceListFile = $csv; Reports = 'PowerBI,FabricCode' } -KeepWorld
    Assert-Equal '' ($global:TWUnrouted -join ' | ')
    Assert-Equal 0 $r.Errors.Count ($r.Errors -join ' | ')
    Assert-Equal 'PowerBIAdmin' ($r.Selected -join ',') 'other reports are ignored'
    $files = (Get-ChildItem -LiteralPath (Join-Path $r.OutDir 'PowerBI') -Recurse -File | ForEach-Object { $_.Name } | Sort-Object) -join ', '
    Assert-Equal 'adminCapacities.csv, adminDatasourceInstances.csv, adminWorkspaces.csv, workspacePlan.csv, workspaceScopeSummary.csv' $files
    Assert-True (Test-Path -LiteralPath (Join-Path $global:TWOut ($r.RunName + '.zip'))) 'zip is on'
    $reqs = @($global:TWRequests | ForEach-Object { $_.Uri })
    Assert-Equal 0 @($reqs | Where-Object { $_ -match 'tenantsettings|/admin/apps|widelyShared|/admin/domains|/admin/activityevents' }).Count 'no settings, apps, sharing links or activity'
    Assert-Equal 0 @($reqs | Where-Object { $_ -notmatch '/admin/' }).Count 'only admin APIs'
    $scan = @($global:TWRequests | Where-Object { $_.Uri -match 'getInfo' })[0].Uri
    Assert-Match $scan 'lineage=True&datasourceDetails=True$'
    $text = (Get-ChildItem -LiteralPath $r.OutDir -Recurse -File | Where-Object { $_.Extension -in '.csv', '.json', '.txt' } | ForEach-Object { [IO.File]::ReadAllText($_.FullName) }) -join "`n"
    foreach ($secret in @('sql1.contoso.com', 'contoso.sharepoint.com', 'Sales team workspace', 'ben@contoso.com', 'vic@contoso.com', 'Sales overview', 'SUM(Sales[Amount])')) {
        Assert-False ($text.Contains($secret)) ('found in the output: ' + $secret)
    }
    Assert-True ($text.Contains('anna@contoso.com')) 'workspace admins are kept'
    Assert-Match ([IO.File]::ReadAllText($r.LogPath)) '-ClientScope: only the scope of the tenant is collected'
    $plan = Import-TWCsv $r 'PowerBI/workspacePlan.csv'
    $got = ($plan | ForEach-Object { '{0}|{1}|{2}|{3}' -f $_.ListedWorkspace, $_.Status, $_.Check, $_.Workspace }) -join ' / '
    Assert-Equal ('Sales|Found||Sales / Orphaned|Found|On the list but empty|Orphaned / Marketing|Not in the tenant|| / PersonalWorkspace Bob|Found|On the list but a personal workspace|PersonalWorkspace Bob / ' +
        '|Not on the list|Shared with content, not on the list|Legacy reports') $got
    Assert-Equal 'Admin API Anna' ($plan[0].Source + ' ' + $plan[0].Listed_owner)
    $sum = @{}; foreach ($x in (Import-TWCsv $r 'PowerBI/workspaceScopeSummary.csv')) { $sum[$x.Measure] = $x.Count }
    Assert-Equal '4 1 2 1 not available from the admin APIs' ('{0} {1} {2} {3} {4}' -f $sum['On the workspace list'], $sum['On the list, not found in the tenant'], $sum['On the list, but empty, personal or not active'],
        $sum['Shared with content, not on the list'], ((Import-TWCsv $r 'PowerBI/workspaceScopeSummary.csv') | Where-Object { $_.Measure -eq 'Shared with content, connected to Git' }).Note)
    $ids = ((Import-TWCsv $r 'findings.csv') | ForEach-Object { $_.Id } | Sort-Object) -join ' '
    Assert-Equal 'PBI-11 PBI-15 PBI-16 PBI-17 PBI-18' $ids
}
Test-Case '-ClientScope without the administrator role: a clear skip, no error, no scope files' {
    $r = Invoke-TWRun -Params @{ ClientScope = $true }
    Assert-Equal 0 $r.Errors.Count
    Assert-Equal 'Skipped' (Get-TWCoverage $r 'Power BI admin (tenant-wide)')['Status']
    Assert-Match ($r.Warnings -join ' | ') '-ClientScope needs an account with the Fabric administrator role'
    Assert-False (Test-Path -LiteralPath (Join-Path $r.OutDir 'PowerBI\workspaceScopeSummary.csv'))
}

Write-Host 'Export' -ForegroundColor Cyan
$run = Invoke-TWRun -Params @{ Reports = 'PowerBI,Fabric,PowerBIExport' }
Test-Case 'report file, item definition (long running operation), dataflow definition' {
    Assert-Equal '' ($global:TWUnrouted -join ' | ')
    $x = Import-TWCsv $run 'PowerBI/exportIndex.csv'
    $pbix = $x | Where-Object { $_.Kind -eq 'pbix' -and $_.Status -eq 'Exported' }
    Assert-Equal 'PowerBI/Sales/reports/Sales overview.pbix' $pbix.Path
    Assert-Equal 'PBIX-BYTES' ([IO.File]::ReadAllText((Join-Path $run.OutDir $pbix.Path)))
    Assert-Equal 'Failed' ($x | Where-Object { $_.Kind -eq 'rdl' }).Status
    $nb = $x | Where-Object { $_.ItemId -eq 'nb1' }
    Assert-Equal 'Exported' $nb.Status; Assert-Equal '2' $nb.Files
    Assert-Equal 'print("bronze")' ([IO.File]::ReadAllText((Join-Path $run.OutDir 'PowerBI\Sales\Notebook\Load bronze\notebook-content.py')))
    Assert-True (Test-Path -LiteralPath (Join-Path $run.OutDir 'PowerBI\Sales\Notebook\Load bronze\_\.platform')) 'a ".." segment cannot leave the item folder'
    Assert-Match ($x | Where-Object { $_.ItemType -eq 'Lakehouse' }).Detail 'OperationNotSupportedForItem'
    Assert-Match ([IO.File]::ReadAllText((Join-Path $run.OutDir 'PowerBI\Sales\dataflows\Customers.json'))) '"entities"'
}
Test-Case 'export together with Dependencies and FabricSql: definitions fetched once, SQL code as files' {
    $r = Invoke-TWRun -Params @{ Reports = 'Fabric,FabricSql,Dependencies,PowerBIExport' }
    Assert-Equal '' ($global:TWUnrouted -join ' | ')
    Assert-Equal 1 @($global:TWRequests | Where-Object { $_.Method -eq 'POST' -and $_.Uri -match "$($global:TWPl1)/getDefinition" }).Count
    Assert-True (Test-Path -LiteralPath (Join-Path $r.OutDir 'PowerBI\Sales\DataPipeline\PL_Daily\pipeline-content.json'))
    Assert-True (Test-Path -LiteralPath (Join-Path $r.OutDir 'PowerBI\Sales\SemanticModel\Sales model\definition\cultures\en-US.tmdl')) 'parts that are not parsed are exported too'
    Assert-Match ([IO.File]::ReadAllText((Join-Path $r.OutDir 'PowerBI\Sales\sql\sales-dw\dbo.vSales.sql'))) 'CREATE VIEW'
    $x = Import-TWCsv $r 'PowerBI/exportIndex.csv'
    Assert-Equal 'Exported 2' (($x | Where-Object { $_.Kind -eq 'sql' }) | ForEach-Object { $_.Status + ' ' + $_.Files })
    Assert-Equal 1 @($x | Where-Object { $_.ItemId -eq 'nb1' }).Count
    Assert-True (@(Import-TWCsv $r 'PowerBI/pipelineActivities.csv').Count -eq 5)
}
Test-Case 'ExportedTo columns point to the exported files' {
    Assert-Equal 'PowerBI/Sales/reports/Sales overview.pbix' (Import-TWCsv $run 'PowerBI/reports.csv' | Where-Object { $_.ReportId -eq 'r1' }).ExportedTo
    Assert-Equal 'PowerBI/Sales/Notebook/Load bronze' (Import-TWCsv $run 'PowerBI/fabricItems.csv' | Where-Object { $_.ItemId -eq 'nb1' }).ExportedTo
}
Test-Case '-PbiExportMaxSizeMB stops report downloads, -Zip packs the run folder' {
    Reset-TWState
    $global:TW.PbiRoutes.Insert(0, (New-TWRoute GET '/reports/r1/Export$' { @{ StatusCode = 200; Content = ''; Headers = @{}; File = ('x' * 1100000) } }))
    $r = Invoke-TWRun -Params @{ Reports = 'PowerBIExport'; PbiExportMaxSizeMB = 1; Zip = $true } -KeepWorld
    $x = Import-TWCsv $r 'PowerBI/exportIndex.csv'
    Assert-Equal 'Skipped' ($x | Where-Object { $_.ItemId -eq 'r2' }).Status
    Assert-True (Test-Path -LiteralPath (Join-Path $global:TWOut ($r.RunName + '.zip')))
}

$failed = Complete-Tests
if (-not $KeepOutput) { Remove-Item -LiteralPath $global:TWOut -Recurse -Force -ErrorAction SilentlyContinue } else { Write-Host ('Output kept in ' + $global:TWOut) }
exit $failed
