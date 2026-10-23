# A small fake Power BI / Fabric tenant for the offline tests.
#   Sales        on capacity cap1: 2 reports, 2 models, dashboard, dataflow, lakehouse, notebook; model users refused
#   Finance DEV  shared capacity, empty, workspace users refused
#   My workspace personal (only listed by Fabric)
# $global:TW.Admin = $true makes the account a Fabric administrator.

$global:TWW1 = 'aaaaaaaa-0000-0000-0000-000000000001'
$global:TWW2 = 'aaaaaaaa-0000-0000-0000-000000000002'
$global:TWWP = 'aaaaaaaa-0000-0000-0000-00000000000f'
$global:TWCap = 'cccccccc-0000-0000-0000-000000000001'

function New-TWRoute {
    param([string] $Method, [string] $Pattern, [scriptblock] $Handler)
    return @{ Method = $Method; Pattern = $Pattern; Handler = $Handler }
}

function New-TWSqlObject {
    param([string] $Schema, [string] $Name, [string] $Type)
    return [ordered]@{ schema = $Schema; name = $Name; type = $Type; create_date = [DateTime]'2026-01-02T03:04:05'; modify_date = [DateTime]'2026-02-03T04:05:06' }
}

function New-TWDefinition {
    # Answer of getDefinition: path -> text
    param($Parts)
    $list = foreach ($k in $Parts.Keys) { @{ path = $k; payload = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([string]$Parts[$k])); payloadType = 'InlineBase64' } }
    return @{ definition = @{ parts = @($list) } }
}

$global:TWNb2 = 'bbbbbbbb-0000-0000-0000-000000000001'   # notebook Transform
$global:TWPl1 = 'bbbbbbbb-0000-0000-0000-000000000002'   # pipeline PL_Daily
$global:TWDf2 = 'bbbbbbbb-0000-0000-0000-000000000003'   # dataflow gen2 Import files

function New-TWList { param($Items) return @{ value = @($Items) } }

function New-TWWorld {
    $w1 = $global:TWW1; $w2 = $global:TWW2; $wp = $global:TWWP; $cap = $global:TWCap
    $pbi = [System.Collections.Generic.List[object]]::new()
    $fab = [System.Collections.Generic.List[object]]::new()

    # ---- Power BI ----
    $pbi.Add((New-TWRoute GET '^/groups\?' { New-TWList @(
                    @{ id = $global:TWW1; name = 'Sales'; type = 'Workspace'; isReadOnly = $false; isOnDedicatedCapacity = $true; capacityId = $global:TWCap; defaultDatasetStorageFormat = 'Small' },
                    @{ id = $global:TWW2; name = 'Finance DEV'; type = 'Workspace'; isReadOnly = $false; isOnDedicatedCapacity = $false }) }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/users$" { New-TWList @(
                    @{ displayName = 'Anna Admin'; emailAddress = 'anna@contoso.com'; identifier = 'anna@contoso.com'; groupUserAccessRight = 'Admin'; principalType = 'User'; graphId = 'g1' },
                    @{ displayName = 'Vic Viewer'; emailAddress = 'vic@contoso.com'; identifier = 'vic@contoso.com'; groupUserAccessRight = 'Viewer'; principalType = 'User'; graphId = 'g2' }) }))
    $pbi.Add((New-TWRoute GET "^/groups/$w2/users$" { New-TWError 401 'PowerBINotAuthorizedException' 'Unauthorized' }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/reports$" { New-TWList @(
                    @{ id = 'r1'; name = 'Sales overview'; reportType = 'PowerBIReport'; datasetId = 'd1'; datasetWorkspaceId = $global:TWW1; webUrl = 'https://app.powerbi.com/r1'; format = 'PBIX' },
                    @{ id = 'r2'; name = 'Invoice: print/layout'; reportType = 'PaginatedReport'; webUrl = 'https://app.powerbi.com/r2' }) }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/reports/r1/pages$" { New-TWList @(@{ name = 'p1'; displayName = 'Overview'; order = 0 }, @{ name = 'p2'; displayName = 'Detail'; order = 1 }) }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/reports/r2/datasources$" { New-TWList @(@{ datasourceType = 'SQL'; name = 'ds'; connectionDetails = @{ server = 'sql1.contoso.com'; database = 'Invoices' }; datasourceId = 'rds1' }) }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/reports/r1/Export$" { @{ StatusCode = 200; Content = ''; Headers = @{}; File = 'PBIX-BYTES' } }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/reports/r2/Export$" { New-TWError 400 'ExportNotSupported' 'Report cannot be exported' }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/datasets$" { New-TWList @(
                    @{ id = 'd1'; name = 'Sales model'; configuredBy = 'gone@contoso.com'; isRefreshable = $true; targetStorageMode = 'Abf'; createdDate = '2025-01-02T03:04:05Z' },
                    @{ id = 'd2'; name = 'Unused model'; configuredBy = 'anna@contoso.com'; isRefreshable = $false; targetStorageMode = 'Abf' }) }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/datasets/d1/datasources$" { New-TWList @(@{ datasourceType = 'Sql'; connectionDetails = @{ server = 'sql1.contoso.com'; database = 'Sales' }; datasourceId = 'ds1'; gatewayId = 'gw1' }) }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/datasets/d2/datasources$" { New-TWList @() }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/datasets/d1/users$" { New-TWError 403 'Forbidden' 'Not the owner' }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/datasets/d2/users$" { New-TWList @(@{ identifier = 'anna@contoso.com'; principalType = 'User'; datasetUserAccessRight = 'ReadWriteReshareExplore' }) }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/datasets/d\d/parameters$" { New-TWList @() }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/datasets/d1/refreshes\?" { New-TWList @(
                    @{ requestId = 'rq2'; refreshType = 'Scheduled'; status = 'Failed'; startTime = '2026-09-28T05:00:00Z'; endTime = '2026-09-28T05:01:30Z'; serviceExceptionJson = '{"errorCode":"ModelRefreshFailed_CredentialsNotSpecified"}' },
                    @{ requestId = 'rq1'; refreshType = 'Scheduled'; status = 'Completed'; startTime = '2026-09-27T05:00:00Z'; endTime = '2026-09-27T05:02:00Z' }) }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/datasets/d1/refreshSchedule$" { @{ enabled = $false; days = @('Monday', 'Friday'); times = @('05:00'); localTimeZoneId = 'UTC'; notifyOption = 'MailOnFailure' } }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/datasets/d2/directQueryRefreshSchedule$" { New-TWError 404 'NotFound' 'No schedule' }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/datasets/upstreamDataflows$" { New-TWList @(@{ datasetObjectId = 'd1'; dataflowObjectId = 'df1'; workspaceObjectId = $global:TWW1 }) }))
    $pbi.Add((New-TWRoute POST "^/groups/$w1/datasets/d1/executeQueries$" {
                param($Path, $Body)
                if ($Body -match 'MEASURES') { return ([char]0xFEFF + '{"results":[{"tables":[{"rows":[{"[Name]":"Total Sales","[Table]":"Sales","[Expression]":"SUM(Sales[Amount])","[IsHidden]":false}]}]}]}') }
                if ($Body -match 'TABLES') { return '{"results":[{"tables":[{"rows":[{"[Name]":"Sales","[StorageMode]":"Import"},{"[Name]":"Date","[StorageMode]":"Import"}]}]}]}' }
                if ($Body -match 'COLUMNS') { return '{"results":[{"tables":[{"rows":[{"[Name]":"Amount","[Table]":"Sales","[DataType]":"Double"}]}]}]}' }
                return '{"results":[{"tables":[{"rows":[{"[Relationship]":"Sales[DateKey] -> Date[DateKey]","[FromTable]":"Sales","[ToTable]":"Date","[IsActive]":true}]}]}]}' }))
    $pbi.Add((New-TWRoute POST "^/groups/$w1/datasets/d2/executeQueries$" { New-TWError 400 'DatasetExecuteQueriesError' 'Query execution is disabled' }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/dashboards$" { New-TWList @(@{ id = 'db1'; displayName = 'KPIs'; isReadOnly = $false; webUrl = 'https://app.powerbi.com/db1' }) }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/dashboards/db1/tiles$" { New-TWList @(@{ id = 't1'; title = 'Revenue'; reportId = 'r1'; datasetId = 'd1'; rowSpan = 2; colSpan = 3 }) }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/dataflows$" { New-TWList @(@{ objectId = 'df1'; name = 'Customers'; configuredBy = 'anna@contoso.com'; modifiedDateTime = '2026-09-01T00:00:00Z'; generation = 1 }) }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/dataflows/df1/datasources$" { New-TWList @(@{ datasourceType = 'Web'; connectionDetails = @{ url = 'https://api.contoso.com/customers' }; datasourceId = 'fds1' }) }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/dataflows/df1/transactions$" { New-TWList @(@{ id = 'tx1'; refreshType = 'Scheduled'; status = 'Success'; startTime = '2026-09-28T01:00:00Z'; endTime = '2026-09-28T01:05:00Z' }) }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/dataflows/df1/upstreamDataflows$" { New-TWList @() }))
    $pbi.Add((New-TWRoute GET "^/groups/$w1/dataflows/df1$" { '{"name":"Customers","entities":[{"name":"Customer"}],"pbi:mashup":{"document":"section Section1;\r\nshared Customers = let\r\n  Source = Web.Contents(\"https://api.contoso.com/customers\")\r\nin Source;"}}' }))
    $pbi.Add((New-TWRoute GET "^/groups/$w2/(reports|datasets|dashboards|dataflows)$" { New-TWList @() }))
    $pbi.Add((New-TWRoute GET "^/(reports|datasets|dashboards)$" { New-TWList @() }))   # My workspace
    $pbi.Add((New-TWRoute GET '^/apps$' { New-TWList @(@{ id = 'app1'; name = 'Sales app'; publishedBy = 'Anna Admin'; lastUpdate = '2026-08-01T00:00:00Z'; workspaceId = $global:TWW1 }) }))
    $pbi.Add((New-TWRoute GET '^/apps/app1/reports$' { New-TWList @(@{ id = 'ar1'; name = 'Sales overview'; datasetId = 'd1' }) }))
    $pbi.Add((New-TWRoute GET '^/apps/app1/dashboards$' { New-TWList @() }))
    $pbi.Add((New-TWRoute GET '^/capacities$' { New-TWList @(@{ id = $global:TWCap; displayName = 'cap-prod'; sku = 'F64'; region = 'West Europe'; state = 'Active'; capacityUserAccessRight = 'Assign' }) }))
    $pbi.Add((New-TWRoute GET "^/capacities/$cap/Workloads$" { New-TWError 403 'Forbidden' 'Capacity admin only' }))
    $pbi.Add((New-TWRoute GET '^/capacities/refreshables\?' { New-TWList @() }))
    $pbi.Add((New-TWRoute GET '^/gateways$' { New-TWList @() }))
    $pbi.Add((New-TWRoute GET '^/pipelines$' { New-TWList @(@{ id = 'pl1'; displayName = 'Sales pipeline'; description = 'dev to prod' }) }))
    $pbi.Add((New-TWRoute GET '^/pipelines/pl1/stages$' { New-TWList @(@{ order = 0; workspaceId = $global:TWW2; workspaceName = 'Finance DEV' }, @{ order = 1; workspaceId = $global:TWW1; workspaceName = 'Sales' }) }))
    $pbi.Add((New-TWRoute GET '^/pipelines/pl1/users$' { New-TWList @(@{ identifier = 'anna@contoso.com'; principalType = 'User'; accessRight = 'Admin' }) }))
    $pbi.Add((New-TWRoute GET '^/pipelines/pl1/operations$' { New-TWList @(@{ id = 'op1'; type = 'Deploy'; status = 'Succeeded'; sourceStageOrder = 0; targetStageOrder = 1; performedBy = @{ userPrincipalName = 'anna@contoso.com' } }) }))
    $pbi.Add((New-TWRoute GET '^/pipelines/pl1/stages/\d/artifacts$' { @{ reports = @(@{ artifactId = 'r1'; artifactDisplayName = 'Sales overview' }); datasets = @(@{ artifactId = 'd1'; artifactDisplayName = 'Sales model' }); dashboards = @() } }))
    $pbi.Add((New-TWRoute GET '^/availableFeatures$' { @{ features = @(@{ name = 'embedTrial'; state = 'Enabled'; extendedState = 'Enabled' }) } }))
    # admin (only reached when $global:TW.Admin)
    $pbi.Add((New-TWRoute GET '^/admin/capacities$' { New-TWList @(@{ id = $global:TWCap; displayName = 'cap-prod'; sku = 'F64'; region = 'West Europe'; state = 'Active'; admins = @('anna@contoso.com') }) }))
    $pbi.Add((New-TWRoute GET '^/admin/groups\?' { New-TWList @(
                    @{ id = $global:TWW1; name = 'Sales'; type = 'Workspace'; state = 'Active'; isOnDedicatedCapacity = $true; capacityId = $global:TWCap; users = @(@{ identifier = 'anna@contoso.com'; emailAddress = 'anna@contoso.com'; groupUserAccessRight = 'Admin'; principalType = 'User' }) },
                    @{ id = 'aaaaaaaa-0000-0000-0000-000000000009'; name = 'Orphaned'; type = 'Workspace'; state = 'Active'; isOnDedicatedCapacity = $false; users = @() },
                    @{ id = 'aaaaaaaa-0000-0000-0000-00000000000a'; name = 'PersonalWorkspace Bob'; type = 'PersonalGroup'; state = 'Active'; users = @() }) }))
    $pbi.Add((New-TWRoute GET '^/admin/workspaces/modified\?' { , @(@{ id = $global:TWW1 }, @{ id = 'aaaaaaaa-0000-0000-0000-000000000009' }) }))
    $pbi.Add((New-TWRoute POST '^/admin/workspaces/getInfo\?' { @{ StatusCode = 202; Content = '{"id":"scan1","status":"NotStarted"}'; Headers = @{} } }))
    $pbi.Add((New-TWRoute GET '^/admin/workspaces/scanStatus/scan1$' { @{ id = 'scan1'; status = 'Succeeded' } }))
    $pbi.Add((New-TWRoute GET '^/admin/workspaces/scanResult/scan1$' { @{
                    workspaces = @(@{ id = $global:TWW1; name = 'Sales'; type = 'Workspace'; state = 'Active'
                            reports = @(@{ id = 'r1'; name = 'Sales overview'; datasetId = 'd1'; modifiedBy = 'anna@contoso.com'; users = @(@{ identifier = 'vic@contoso.com'; displayName = 'Vic Viewer'; principalType = 'User'; reportUserAccessRight = 'Read' }) })
                            datasets = @(@{ id = 'd1'; name = 'Sales model'; configuredBy = 'gone@contoso.com'; targetStorageMode = 'Abf'; datasourceUsages = @(@{ datasourceInstanceId = 'dsi1' })
                                    expressions = @(@{ name = 'ServerName'; expression = '"sql1.contoso.com" meta [IsParameterQuery=true]' })
                                    tables = @(@{ name = 'Sales'; columns = @(@{ name = 'Amount'; dataType = 'Double' }); measures = @(@{ name = 'Total Sales'; expression = 'SUM(Sales[Amount])' }); source = @(@{ expression = 'let Source = Sql.Database(ServerName, "Sales") in Source' }) }) })
                            dashboards = @(); dataflows = @()
                            Lakehouse = @(@{ id = 'lh1'; name = 'Bronze'; modifiedBy = 'anna@contoso.com' }) })
                    datasourceInstances = @(@{ datasourceId = 'dsi1'; datasourceType = 'Sql'; connectionDetails = @{ server = 'sql1.contoso.com'; database = 'Sales' }; gatewayId = 'gw1' }) } }))
    $pbi.Add((New-TWRoute GET '^/admin/apps\?' { New-TWList @(@{ id = 'app1'; name = 'Sales app'; workspaceId = $global:TWW1 }) }))
    $pbi.Add((New-TWRoute GET '^/admin/widelySharedArtifacts/publishedToWeb$' { @{ ArtifactAccessEntities = @(@{ artifactId = 'r1'; displayName = 'Sales overview'; artifactType = 'Report'; shareType = 'PublishToWeb'; sharer = @{ emailAddress = 'anna@contoso.com'; principalType = 'User' } }) } }))
    $pbi.Add((New-TWRoute GET '^/admin/widelySharedArtifacts/linksSharedToWholeOrganization$' { @{ ArtifactAccessEntities = @() } }))
    $pbi.Add((New-TWRoute GET '^/admin/activityevents\?.*continuationToken=' { @{ activityEventEntities = @(@{ Id = 'e3'; CreationTime = '2026-09-28T12:00:00Z'; UserId = 'vic@contoso.com'; Activity = 'ExportReport'; Operation = 'ExportReport' }) } }))
    $pbi.Add((New-TWRoute GET '^/admin/activityevents\?startDateTime=.2026-09-28' { @{
                    activityEventEntities = @(@{ Id = 'e1'; CreationTime = '2026-09-28T10:00:00Z'; UserId = 'vic@contoso.com'; Activity = 'ViewReport'; Operation = 'ViewReport'; WorkSpaceName = 'Sales'; ReportName = 'Sales overview'; IsSuccess = $true },
                        @{ Id = 'e2'; CreationTime = '2026-09-28T11:00:00Z'; UserId = 'anna@contoso.com'; Activity = 'ShareReport'; Operation = 'ShareReport' })
                    continuationUri = 'https://api.powerbi.com/v1.0/myorg/admin/activityevents?continuationToken=abc' } }))
    $pbi.Add((New-TWRoute GET '^/admin/activityevents\?' { @{ activityEventEntities = @(); continuationUri = $null } }))

    # ---- Fabric ----
    $fab.Add((New-TWRoute GET '^/workspaces\?continuationToken=' { @{ value = @(
                        @{ id = $global:TWW2; displayName = 'Finance DEV'; type = 'Workspace'; description = 'finance sandbox' },
                        @{ id = $global:TWWP; displayName = 'My workspace'; type = 'Personal' }) } }))
    $fab.Add((New-TWRoute GET '^/workspaces$' { @{ value = @(@{ id = $global:TWW1; displayName = 'Sales'; type = 'Workspace'; capacityId = $global:TWCap; domainId = 'dom1' })
                    continuationUri = 'https://api.fabric.microsoft.com/v1/workspaces?continuationToken=p2' } }))
    $fab.Add((New-TWRoute GET "^/workspaces/$w1$" { @{ id = $global:TWW1; displayName = 'Sales'; capacityRegion = 'West Europe'
                    oneLakeEndpoints = @{ blobEndpoint = 'https://westeurope-onelake.blob.fabric.microsoft.com'; dfsEndpoint = 'https://westeurope-onelake.dfs.fabric.microsoft.com' } } }))
    $fab.Add((New-TWRoute GET "^/workspaces/($w2|$wp)$" { @{ displayName = 'x' } }))
    $fab.Add((New-TWRoute GET "^/workspaces/$w1/roleAssignments$" { New-TWList @(
                    @{ id = 'u1'; role = 'Admin'; principal = @{ id = 'u1'; displayName = 'Anna Admin'; type = 'User'; userDetails = @{ userPrincipalName = 'anna@contoso.com' } } },
                    @{ id = 'sp1'; role = 'Contributor'; principal = @{ id = 'sp1'; displayName = 'deploy-sp'; type = 'ServicePrincipal'; servicePrincipalDetails = @{ aadAppId = 'app-123' } } }) }))
    $fab.Add((New-TWRoute GET "^/workspaces/($w2|$wp)/roleAssignments$" { New-TWList @(@{ id = 'u1'; role = 'Admin'; principal = @{ id = 'u1'; displayName = 'Anna Admin'; type = 'User'; userDetails = @{ userPrincipalName = 'anna@contoso.com' } } }) }))
    $fab.Add((New-TWRoute GET "^/workspaces/$w1/items$" { New-TWList @(
                    @{ id = 'lh1'; displayName = 'Bronze'; type = 'Lakehouse' }, @{ id = 'lh2'; displayName = 'Gold'; type = 'Lakehouse' },
                    @{ id = 'dw1'; displayName = 'sales-dw'; type = 'Warehouse' }, @{ id = $global:TWNb2; displayName = 'Transform'; type = 'Notebook' },
                    @{ id = $global:TWPl1; displayName = 'PL_Daily'; type = 'DataPipeline' }, @{ id = $global:TWDf2; displayName = 'Import files'; type = 'Dataflow' }, @{ id = 'nb1'; displayName = 'Load bronze'; type = 'Notebook'; description = 'nightly load'; folderId = 'f2'; logicalId = 'l1'; tags = @(@{ id = 't1'; displayName = 'Certified source' }) },
                    @{ id = 'd1'; displayName = 'Sales model'; type = 'SemanticModel' }, @{ id = 'r1'; displayName = 'Sales overview'; type = 'Report' }) }))
    $fab.Add((New-TWRoute GET "^/workspaces/($w2|$wp)/items$" { New-TWList @() }))
    $fab.Add((New-TWRoute GET "^/workspaces/$w1/git/connection$" { @{ gitConnectionState = 'ConnectedAndInitialized'
                    gitProviderDetails = @{ gitProviderType = 'AzureDevOps'; organizationName = 'contoso'; projectName = 'BI'; repositoryName = 'sales'; branchName = 'main'; directoryName = '/' }
                    gitSyncDetails = @{ head = 'abc123'; lastSyncTime = '2026-09-20T00:00:00Z' } } }))
    $fab.Add((New-TWRoute GET "^/workspaces/$w2/git/connection$" { @{ gitConnectionState = 'NotConnected' } }))
    $fab.Add((New-TWRoute GET "^/workspaces/$w1/spark/settings$" { @{ automaticLog = @{ enabled = $true }; pool = @{ customizeComputeEnabled = $true; defaultPool = @{ name = 'Starter Pool'; type = 'Workspace' } }; environment = @{ runtimeVersion = '1.3' } } }))
    $fab.Add((New-TWRoute GET "^/workspaces/$w1/spark/pools$" { New-TWList @(@{ id = 'sp-1'; name = 'Starter Pool'; type = 'Workspace'; nodeFamily = 'MemoryOptimized'; nodeSize = 'Medium'; autoScale = @{ enabled = $true; minNodeCount = 1; maxNodeCount = 10 } }) }))
    $fab.Add((New-TWRoute GET "^/workspaces/$w1/managedPrivateEndpoints$" { New-TWList @() }))
    $fab.Add((New-TWRoute GET "^/workspaces/$w1/lakehouses$" { New-TWList @(@{ id = 'lh1'; displayName = 'Bronze'; type = 'Lakehouse'; properties = @{
                            oneLakeTablesPath = 'https://onelake.dfs.fabric.microsoft.com/w1/lh1/Tables'; oneLakeFilesPath = 'https://onelake.dfs.fabric.microsoft.com/w1/lh1/Files'
                            sqlEndpointProperties = @{ id = 'sql-ep1'; connectionString = 'abc.datawarehouse.fabric.microsoft.com'; provisioningStatus = 'Success' } } },
                    @{ id = 'lh2'; displayName = 'Gold'; type = 'Lakehouse'; properties = @{ defaultSchema = 'dbo'
                            sqlEndpointProperties = @{ id = 'sql-ep2'; connectionString = 'abc.datawarehouse.fabric.microsoft.com'; provisioningStatus = 'Success' } } }) }))
    $fab.Add((New-TWRoute GET "^/workspaces/$w1/lakehouses/lh2/tables$" { New-TWResponse -Obj @{ errorCode = 'UnsupportedOperationForSchemasEnabledLakehouse'; message = 'Not supported for Lakehouse with schemas enabled' } -Status 400 }))
    $fab.Add((New-TWRoute GET "^/workspaces/$w1/items/lh2/shortcuts$" { New-TWList @() }))
    $fab.Add((New-TWRoute GET "^/workspaces/$w1/warehouses$" { New-TWList @(@{ id = 'dw1'; displayName = 'sales-dw'; type = 'Warehouse'
                        properties = @{ connectionString = 'abc.datawarehouse.fabric.microsoft.com'; createdDate = '2025-05-01T00:00:00Z'; lastUpdatedTime = '2026-09-01T00:00:00Z' } }) }))
    $fab.Add((New-TWRoute GET "^/workspaces/$w1/lakehouses/lh1/tables$" { @{ data = @(@{ name = 'customers'; type = 'Managed'; format = 'delta'; location = 'abfss://x/Tables/customers' }, @{ name = 'orders'; type = 'Managed'; format = 'delta'; location = 'abfss://x/Tables/orders' }) } }))
    $fab.Add((New-TWRoute GET "^/workspaces/$w1/items/lh1/shortcuts$" { New-TWList @(@{ name = 'raw'; path = 'Files'; target = @{ type = 'AdlsGen2'; adlsGen2 = @{ location = 'https://stor.dfs.core.windows.net'; subpath = '/raw' } } }) }))
    $fab.Add((New-TWRoute GET "^/workspaces/$w1/items/(lh1|lh2|dw1|nb1)/connections$" { New-TWList @() }))
    $fab.Add((New-TWRoute GET "^/workspaces/$w1/items/d1/connections$" { New-TWList @(@{ id = 'c1'; displayName = 'sql1 Sales'; connectivityType = 'OnPremisesGateway'; gatewayId = 'gw1'; connectionDetails = @{ type = 'SQL'; path = 'sql1.contoso.com;Sales' } }) }))
    $fab.Add((New-TWRoute GET "^/workspaces/$w1/items/nb1/jobs/RunNotebook/schedules$" { New-TWList @(@{ id = 'sch1'; enabled = $true; createdDateTime = '2026-01-01T00:00:00Z'; configuration = @{ type = 'Daily'; times = @('02:00') }; owner = @{ displayName = 'Anna Admin' } }) }))
    $fab.Add((New-TWRoute POST "^/workspaces/$w1/items/nb1/getDefinition(\?format=ipynb)?$" { @{ StatusCode = 202; Content = ''; Headers = @{ 'x-ms-operation-id' = 'op-nb1'; 'Retry-After' = '1' } } }))
    $fab.Add((New-TWRoute GET '^/operations/op-nb1$' { @{ status = 'Succeeded' } }))
    $fab.Add((New-TWRoute GET '^/operations/op-nb1/result$' { @{ definition = @{ parts = @(
                            @{ path = 'notebook-content.py'; payload = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('print("bronze")')); payloadType = 'InlineBase64' },
                            @{ path = '../.platform'; payload = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes('{}')); payloadType = 'InlineBase64' }) } } }))
    $fab.Add((New-TWRoute POST "^/workspaces/$w1/items/\w+/getDefinition$" { New-TWResponse -Obj @{ errorCode = 'OperationNotSupportedForItem'; message = 'Not supported' } -Status 400 }))
    $fab.Add((New-TWRoute GET '^/connections$' { New-TWList @(@{ id = 'c1'; displayName = 'sql1 Sales'; connectivityType = 'OnPremisesGateway'; gatewayId = 'gw1'; privacyLevel = 'Organizational'
                        connectionDetails = @{ type = 'SQL'; path = 'sql1.contoso.com;Sales' }; credentialDetails = @{ credentialType = 'Basic'; singleSignOnType = 'None'; connectionEncryption = 'Encrypted'; skipTestConnection = $false } }) }))
    $fab.Add((New-TWRoute GET '^/gateways$' { New-TWList @(@{ id = 'gw1'; displayName = 'onprem-gw'; type = 'OnPremises'; version = '3000.1'; numberOfMemberGateways = 2 }) }))
    $fab.Add((New-TWRoute GET '^/capacities$' { New-TWList @(@{ id = $global:TWCap; displayName = 'cap-prod'; sku = 'F64'; region = 'West Europe'; state = 'Active' },
                    @{ id = 'cccccccc-0000-0000-0000-000000000002'; displayName = 'cap-test'; sku = 'F2'; region = 'North Europe'; state = 'Inactive' }) }))
    $fab.Add((New-TWRoute GET '^/admin/tenantsettings$' { @{ tenantSettings = @(
                        @{ settingName = 'PublishToWeb'; title = 'Publish to web'; enabled = $true; canSpecifySecurityGroups = $true; tenantSettingGroup = 'Export and sharing settings'; enabledSecurityGroups = @(@{ graphId = 'sg1'; name = 'Marketing' }) },
                        @{ settingName = 'AllowServicePrincipalsUseReadAdminAPIs'; title = 'Service principals can access read-only admin APIs'; enabled = $false; tenantSettingGroup = 'Admin API settings' }) } }))
    $fab.Add((New-TWRoute GET '^/admin/domains$' { @{ domains = @(@{ id = 'dom1'; displayName = 'Sales domain'; contributorsScope = 'AllTenant' }) } }))

    # ---- folders, network, Git status, Spark sessions, OneLake roles, deployment pipelines, domains ----
    $fab.Insert(0, (New-TWRoute GET "^/workspaces/$w1/folders$" { New-TWList @(@{ id = 'f1'; displayName = 'ETL' }, @{ id = 'f2'; displayName = 'Bronze'; parentFolderId = 'f1' }) }))
    $fab.Insert(0, (New-TWRoute GET "^/workspaces/($w2|$wp)/folders$" { New-TWList @() }))
    $fab.Insert(0, (New-TWRoute GET "^/workspaces/$w1/networking/communicationPolicy$" { @{ inbound = @{ publicAccessRules = @{ defaultAction = 'Allow' } }; outbound = @{ publicAccessRules = @{ defaultAction = 'Deny' } } } }))
    $fab.Insert(0, (New-TWRoute GET "^/workspaces/($w2|$wp)/networking/communicationPolicy$" { New-TWResponse -Obj @{ errorCode = 'FeatureNotAvailable'; message = 'not on this capacity' } -Status 400 }))
    $fab.Insert(0, (New-TWRoute GET "^/workspaces/$w1/git/status$" { @{ StatusCode = 202; Content = ''; Headers = @{ 'x-ms-operation-id' = 'op-git' } } }))
    $fab.Insert(0, (New-TWRoute GET '^/operations/op-git$' { @{ status = 'Succeeded' } }))
    $fab.Insert(0, (New-TWRoute GET '^/operations/op-git/result$' { @{ workspaceHead = 'abc123'; remoteCommitHash = 'abc999'; changes = @(
                    @{ itemMetadata = @{ itemIdentifier = @{ objectId = 'nb1'; logicalId = 'l1' }; itemType = 'Notebook'; displayName = 'Load bronze' }; workspaceChange = 'Modified'; conflictType = 'None' },
                    @{ itemMetadata = @{ itemIdentifier = @{ logicalId = 'l9' }; itemType = 'Report'; displayName = 'New report' }; remoteChange = 'Added'; conflictType = 'None' }) } }))
    $fab.Insert(0, (New-TWRoute GET "^/workspaces/$w1/spark/livySessions$" { New-TWList @(@{ livyId = 'livy1'; sparkApplicationId = 'application_1'; state = 'Succeeded'; origin = 'SubmittedJob'; jobType = 'SparkBatch'
                        submittedDateTime = '2026-09-28T02:00:00'; item = @{ workspaceId = $global:TWW1; itemId = 'nb1' }; totalDuration = @{ value = 2; timeUnit = 'Minutes' }; submitter = @{ id = 'u1'; type = 'User' } }) }))
    $fab.Insert(0, (New-TWRoute GET "^/workspaces/$w1/dataflows$" { New-TWList @(@{ id = $global:TWDf2; displayName = 'Import files'; type = 'Dataflow'; properties = @{ isParametric = $false } }) }))
    $fab.Insert(0, (New-TWRoute GET "^/workspaces/$w1/items/lh1/dataAccessRoles$" { New-TWList @(@{ name = 'SalesReaders'
                        decisionRules = @(@{ effect = 'Permit'; permission = @(@{ attributeName = 'Path'; attributeValueIncludedIn = @('/Tables/customers') }, @{ attributeName = 'Action'; attributeValueIncludedIn = @('Read') }) })
                        members = @{ microsoftEntraMembers = @(@{ tenantId = 't'; objectId = 'grp-1' }); fabricItemMembers = @(@{ sourcePath = "$($global:TWW1)/lh1"; itemAccess = @('ReadAll') }) } }) }))
    $fab.Insert(0, (New-TWRoute GET "^/workspaces/$w1/items/lh2/dataAccessRoles$" { New-TWList @() }))
    $fab.Insert(0, (New-TWRoute GET "^/workspaces/$w1/items/lh\d/jobs/TableMaintenance/schedules$" { New-TWList @() }))
    $fab.Insert(0, (New-TWRoute GET '^/domains$' { New-TWList @(@{ id = 'dom1'; displayName = 'Sales domain'; description = 'all sales data' }) }))
    $fab.Insert(0, (New-TWRoute GET '^/deploymentPipelines$' { New-TWList @(@{ id = 'dp1'; displayName = 'Sales pipeline' }) }))
    $fab.Insert(0, (New-TWRoute GET '^/deploymentPipelines/dp1/stages$' { New-TWList @(@{ id = 'st0'; order = 0; displayName = 'Development'; isPublic = $false },
                    @{ id = 'st1'; order = 1; displayName = 'Production'; workspaceId = $global:TWW1; workspaceName = 'Sales'; isPublic = $true }) }))
    $fab.Insert(0, (New-TWRoute GET '^/deploymentPipelines/dp1/stages/st1/items$' { New-TWList @(@{ itemId = 'nb1'; itemDisplayName = 'Load bronze'; itemType = 'Notebook'; sourceItemId = 'src-nb'; lastDeploymentTime = '2026-09-01T00:00:00Z' },
                    @{ itemId = 'lh1'; itemDisplayName = 'Bronze'; itemType = 'Lakehouse' }) }))
    $fab.Insert(0, (New-TWRoute GET '^/deploymentPipelines/dp1/roleAssignments$' { New-TWList @(@{ role = 'Admin'; principal = @{ id = 'u1'; type = 'User' } }) }))

    # ---- run history, definitions of the items that the Dependencies report parses ----
    $fab.Insert(0, (New-TWRoute GET "^/workspaces/$w1/items/nb1/jobs/instances$" { New-TWList @(
                    @{ id = 'run1'; jobType = 'RunNotebook'; invokeType = 'Scheduled'; status = 'Completed'; startTimeUtc = '2026-09-27T02:00:00'; endTimeUtc = '2026-09-27T02:05:00' },
                    @{ id = 'run2'; jobType = 'RunNotebook'; invokeType = 'Scheduled'; status = 'Failed'; startTimeUtc = '2026-09-28T02:00:00'; endTimeUtc = '2026-09-28T02:01:00'; failureReason = @{ message = 'Table not found' } }) }))
    $fab.Insert(0, (New-TWRoute GET "^/workspaces/$w1/items/bbbbbbbb-[^/]+/(connections|jobs/instances|jobs/\w+/schedules)$" { New-TWList @() }))
    $fab.Insert(0, (New-TWRoute POST "^/workspaces/$w1/items/$($global:TWPl1)/getDefinition$" { New-TWDefinition @{ 'pipeline-content.json' = (ConvertTo-Json -Depth 20 -Compress @{ properties = @{ activities = @(
                                @{ name = 'Run transform'; type = 'TridentNotebook'; dependsOn = @(); typeProperties = @{ notebookId = $global:TWNb2; workspaceId = $global:TWW1 } },
                                @{ name = 'Load model'; type = 'SqlServerStoredProcedure'; dependsOn = @(@{ activity = 'Run transform' })
                                    typeProperties = @{ storedProcedureName = '[dm].[load_data_model]'; storedProcedureParameters = @{ eo_start = @{ value = '1000'; type = 'Int32' } } } },
                                @{ name = 'Refresh model'; type = 'PBISemanticModelRefresh'; dependsOn = @(@{ activity = 'Load model' }); typeProperties = @{ datasetId = 'dddddddd-0000-0000-0000-00000000000d'; groupId = 'dddddddd-0000-0000-0000-00000000000e' } },
                                @{ name = 'On failure'; type = 'IfCondition'; dependsOn = @(@{ activity = 'Refresh model' })
                                    typeProperties = @{ expression = @{ value = "@equals(activity('Refresh model').Status,'Failed')"; type = 'Expression' }
                                        ifTrueActivities = @(@{ name = 'Mail'; type = 'Office365Email'; dependsOn = @(); typeProperties = @{ inputs = @{ method = 'post' } } }) } }) } }); '.platform' = '{}' } }))
    $fab.Insert(0, (New-TWRoute POST "^/workspaces/$w1/items/$($global:TWNb2)/getDefinition(\?format=ipynb)?$" { New-TWDefinition @{ 'notebook-content.py' = @'
# Fabric notebook source

# METADATA ********************

# META {
# META   "dependencies": {
# META     "lakehouse": {
# META       "default_lakehouse": "lh2",
# META       "default_lakehouse_name": "Gold",
# META       "default_lakehouse_workspace_id": "aaaaaaaa-0000-0000-0000-000000000001",
# META       "known_lakehouses": [ { "id": "lh1" }, { "id": "lh2" } ]
# META     }
# META   }
# META }

# CELL ********************

from pyspark.sql.functions import col
src = 'abfss://aaaaaaaa-0000-0000-0000-000000000001@onelake.dfs.fabric.microsoft.com/lh1/Tables/'
# old = 'abfss://x@oldaccount.dfs.core.windows.net/raw'
df = spark.read.table("silver.sales")
df2 = spark.sql("SELECT * FROM gold.dim_date d JOIN silver.sales s ON 1 = 1")
df.write.mode("overwrite").saveAsTable("gold.fact_sales")
notebookutils.notebook.run("Load bronze", 600)
requests.get("https://api.contoso.com/rates")

# METADATA ********************

# META {
# META   "language": "python"
# META }
'@ } }))
    $fab.Insert(0, (New-TWRoute POST "^/workspaces/$w1/items/$($global:TWDf2)/getDefinition$" { New-TWDefinition @{ 'mashup.pq' = @'
[StagingDefinition = [Kind = "FastCopy"]]
section Section1;
shared Budget = let
  Source = SharePoint.Files("https://contoso.sharepoint.com/sites/Finance", [ApiVersion = 15]),
  Navigation = Source{[Name = "Budget.xlsx", #"Folder Path" = "https://contoso.sharepoint.com/sites/Finance/Shared Documents/"]}[Content],
  Imported = Excel.Workbook(Navigation, null, true)
in Imported;
shared Orders = let
  Source = Sql.Database("abc.datawarehouse.fabric.microsoft.com", "sales-dw")
in Source;
shared Budget_DataDestination = let
  Pattern = Lakehouse.Contents([CreateNavigationProperties = false]),
  Navigation_1 = Pattern{[workspaceId = "aaaaaaaa-0000-0000-0000-000000000001"]}[Data],
  Navigation_2 = Navigation_1{[lakehouseId = "eeeeeeee-0000-0000-0000-00000000000e"]}[Data]
in Navigation_2;
'@; 'queryMetadata.json' = '{}' } }))
    $fab.Insert(0, (New-TWRoute POST "^/workspaces/$w1/items/d1/getDefinition$" { New-TWDefinition ([ordered]@{
                    'definition/cultures/en-US.tmdl'  = 'cultureInfo en-US'
                    'definition/expressions.tmdl'     = "expression Reference =`n`t`tlet`n`t`t    Source = Sql.Databases(`"abc.datawarehouse.fabric.microsoft.com`"),`n`t`t    db = Source{[Name=`"sales-dw`"]}[Data]`n`t`tin`n`t`t    db`n`tlineageTag: 1`n`n`tannotation PBI_ResultType = Table`n"
                    'definition/roles/RLS Sales.tmdl' = "role 'RLS Sales'`n`tmodelPermission: read`n`n`ttablePermission Sales = [Email] = USERPRINCIPALNAME()`n`n`tannotation PBI_Id = 1`n"
                    'definition/tables/Sales.tmdl'    = "table Sales`n`tlineageTag: 2`n`n`tmeasure 'Total Sales' = SUM(Sales[Amount])`n`t`tformatString: #,0`n`t`tdisplayFolder: KPI`n`n`tmeasure Margin =`n`t`t`tVAR s = [Total Sales]`n`t`t`tRETURN s * 0.2`n`t`tformatString: 0.0%`n`n`tcolumn Amount`n`t`tdataType: double`n`t`tsourceColumn: Amount`n`n`tcolumn Double = Sales[Amount] * 2`n`t`tdataType: double`n`n`tpartition Sales = m`n`t`tmode: import`n`t`tsource =`n`t`t`t`tlet`n`t`t`t`t    Source = Reference,`n`t`t`t`t    t = Source{[Schema=`"dbo`",Item=`"FactSales`"]}[Data]`n`t`t`t`tin`n`t`t`t`t    t`n`n`tannotation PBI_ResultType = Table`n"
                    'definition/tables/Date.tmdl'     = "table Date`n`n`tcolumn Day`n`t`tdataType: dateTime`n`n`tpartition Date = entity`n`t`tmode: directLake`n`t`tsource`n`t`t`tentityName: dim_date`n`t`t`tschemaName: gold`n`t`t`texpressionSource: Reference`n"
                }) }))

    # ---- OneLake (file listing) ----
    $one = [System.Collections.Generic.List[object]]::new()
    $one.Add((New-TWRoute GET "^/$w1\?.*directory=lh1%2FTables.*continuation=" { @{ paths = @(@{ name = 'lh1/Tables/orders/part-2.parquet'; contentLength = '500'; lastModified = 'Wed, 02 Sep 2026 10:00:00 GMT' }) } }))
    $one.Add((New-TWRoute GET "^/$w1\?.*directory=lh1%2FTables" { @{ StatusCode = 200; Headers = @{ 'x-ms-continuation' = 'page2' }; Content = (ConvertTo-Json -Depth 5 -Compress @{ paths = @(
                            @{ name = 'lh1/Tables/customers'; isDirectory = 'true'; contentLength = '0' }, @{ name = 'lh1/Tables/customers/_delta_log'; isDirectory = 'true'; contentLength = '0' },
                            @{ name = 'lh1/Tables/customers/_delta_log/0.json'; contentLength = '100'; lastModified = 'Tue, 01 Sep 2026 10:00:00 GMT' },
                            @{ name = 'lh1/Tables/customers/part-1.parquet'; contentLength = '1048576'; lastModified = 'Tue, 01 Sep 2026 10:00:00 GMT' },
                            @{ name = 'lh1/Tables/orders'; isDirectory = 'true'; contentLength = '0' }, @{ name = 'lh1/Tables/orders/part-1.parquet'; contentLength = '1500'; lastModified = 'Tue, 01 Sep 2026 10:00:00 GMT' }) }) } }))
    $one.Add((New-TWRoute GET "^/$w1\?.*directory=lh1%2FFiles" { @{ paths = @(@{ name = 'lh1/Files/raw'; isDirectory = 'true' }, @{ name = 'lh1/Files/raw/a.csv'; contentLength = '2000' }) } }))
    $one.Add((New-TWRoute GET "^/$w1\?.*directory=lh2%2FTables" { @{ paths = @(@{ name = 'lh2/Tables/gold'; isDirectory = 'true' }, @{ name = 'lh2/Tables/gold/fact_sales'; isDirectory = 'true' },
                    @{ name = 'lh2/Tables/gold/fact_sales/part-1.parquet'; contentLength = '3000'; lastModified = 'Tue, 01 Sep 2026 10:00:00 GMT' }) } }))
    $one.Add((New-TWRoute GET "^/$w1\?.*directory=lh2%2FFiles" { @{ paths = @() } }))
    $one.Add((New-TWRoute GET "^/$w1\?.*directory=dw1%2FTables" { New-TWError 403 'AuthorizationPermissionMismatch' 'This request is not authorized' }))

    # ---- Power Automate ----
    $flow = [System.Collections.Generic.List[object]]::new()
    $flow.Add((New-TWRoute GET '^/providers/Microsoft\.ProcessSimple/environments\?' { New-TWList @(@{ name = 'Default-1111'; location = 'europe'; properties = @{ displayName = 'Contoso (default)'; isDefault = $true; environmentSku = 'Default' } }) }))
    $flow.Add((New-TWRoute GET '/scopes/admin/environments/' { New-TWError 403 'Forbidden' 'Not an administrator' }))
    $flow.Add((New-TWRoute GET '/environments/Default-1111/flows\?.*filter=' { New-TWList @() }))
    $flow.Add((New-TWRoute GET '/environments/Default-1111/flows\?' { New-TWList @(@{ name = 'ffffffff-0000-0000-0000-000000000001'; properties = @{ displayName = 'Refresh sales model'; state = 'Started'; createdTime = '2026-01-01T00:00:00Z'
                            creator = @{ userId = 'u1' }; definitionSummary = @{ triggers = @(@{ type = 'Recurrence' }); actions = @(@{ type = 'OpenApiConnection'; api = @{ name = 'shared_powerbi' } }) }
                            connectionReferences = @{ shared_powerbi = @{ displayName = 'Power BI' } } } }) }))

    return @{
        OneLakeRoutes = $one
        FlowRoutes   = $flow
        Context      = [pscustomobject]@{ Account = [pscustomobject]@{ Id = 'tester@contoso.com' }; Tenant = [pscustomobject]@{ Id = '11111111-2222-3333-4444-555555555555' } }
        Admin        = $false
        Throttle     = [System.Collections.Generic.List[object]]::new()
        SqlSets      = @{
            'sales-dw' = @{
                columns    = @([ordered]@{ schema = 'dbo'; table = 'FactSales'; column = 'Amount'; ordinal = 1; type = 'decimal'; length = $null; precision = 18; scale = 2; nullable = 'YES' },
                    [ordered]@{ schema = 'dbo'; table = 'FactSales'; column = 'CustomerKey'; ordinal = 2; type = 'int'; length = $null; precision = 10; scale = 0; nullable = 'NO' })
                modules    = @([ordered]@{ schema = 'dbo'; name = 'vSales'; type = 'V'; definition = "CREATE VIEW dbo.vSales AS`nSELECT * FROM Bronze.dbo.orders" },
                    [ordered]@{ schema = 'dm'; name = 'load_data_model'; type = 'P'; definition = 'CREATE PROC dm.load_data_model AS SELECT 1' })
                references = @([ordered]@{ schema = 'dbo'; name = 'vSales'; ref_database = 'Bronze'; ref_schema = 'dbo'; ref_entity = 'orders' })
            }
            'Gold'     = @{ modules = 'The SELECT permission was denied on the object sql_modules' }
        }
        # SQL endpoints: database -> rows of the catalog query, or a text = the connection error
        Sql          = @{
            'Bronze'   = @((New-TWSqlObject 'dbo' 'customers' 'U'), (New-TWSqlObject 'dbo' 'orders' 'U'), (New-TWSqlObject 'dbo' 'v_orders' 'V'))
            'Gold'     = @((New-TWSqlObject 'gold' 'fact_sales' 'U'), (New-TWSqlObject 'gold' 'dim_date' 'U'), (New-TWSqlObject 'silver' 'sales' 'U'))
            'sales-dw' = @((New-TWSqlObject 'dbo' 'FactSales' 'U'), (New-TWSqlObject 'dbo' 'vSales' 'V'), (New-TWSqlObject 'rpt' 'vMargin' 'V'),
                (New-TWSqlObject 'dm' 'load_data_model' 'P'), (New-TWSqlObject 'dbo' 'fn_fx' 'FN'))
        }
        PbiRoutes    = $pbi
        FabricRoutes = $fab
    }
}
