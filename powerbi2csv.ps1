<#PSScriptInfo
.VERSION 1.0.0
.GUID 6f0b1c0e-6a55-4f6b-9a0b-0d0c6f3d7e21
.AUTHOR Aleksandr Reznik
.COMPANYNAME it4it.solutions
.COPYRIGHT (c) 2026 Aleksandr Reznik. MIT License.
.TAGS PowerBI MicrosoftFabric Fabric Inventory CSV Migration TenantToTenant Documentation OneLake Lakehouse SemanticModel Notebook
.LICENSEURI https://github.com/AleksandrReznik/powerbi2csv/blob/main/LICENSE
.PROJECTURI https://github.com/AleksandrReznik/powerbi2csv
.EXTERNALMODULEDEPENDENCIES Az.Accounts
.RELEASENOTES
1.0.0 First public release: 11 reports, up to 90 CSV files, 14 findings. Same structure and helpers as azure2csv.ps1
so that the Power BI blocks can be merged into it. See CHANGELOG.md.
#>

#Requires -Version 5.1

<#
.SYNOPSIS
Inventory of a Power BI / Microsoft Fabric tenant to CSV files (read-only).

.DESCRIPTION
powerbi2csv reads everything the signed-in account can see in Power BI and Microsoft Fabric and writes it to CSV
files (UTF-8 with BOM, ready for Excel), plus an optional Excel workbook, a log and a machine-readable run summary.

It uses the Az sign-in (Az.Accounts) to get tokens for the Power BI REST API and the Fabric REST API. No app
registration, no secrets, no other module. Only read operations are used: GET requests, three POST requests that
do not change anything (DAX INFO queries, item definition export, admin metadata scan) and one SELECT on the system
catalog of each lakehouse SQL endpoint and warehouse (report FabricSql, needs outbound TCP 1433).

What is collected depends on the account:
  - Any user: every workspace the account is a member of, with reports, semantic models, data sources, refresh
    schedules and history, dashboards, dataflows, apps, capacities, gateways, deployment pipelines and all Fabric
    items (lakehouses, warehouses, notebooks, pipelines, ...), Git connections, Spark settings, shortcuts, and
    the tables, views and stored procedures of lakehouses and warehouses.
  - Fabric administrator: also the whole tenant (all workspaces including personal and orphaned ones, who has access
    to what, data source lineage, model expressions, tenant settings, activity log).

The script has the same structure as azure2csv.ps1 (report catalog, datasets, report builders, findings) and uses
the same helper functions, so its Power BI blocks can be pasted into azure2csv.ps1. See README.md.

.PARAMETER TenantId
Tenant to sign in to. Default: the tenant of the current Az sign-in.

.PARAMETER Workspace
Only these workspaces: names (wildcards allowed) or IDs. Default: all workspaces the account can see.

.PARAMETER ExcludeWorkspace
Workspaces to leave out: names (wildcards allowed) or IDs.

.PARAMETER PbiIncludePersonalWorkspaces
Also inventory your own "My workspace". (Other users' personal workspaces are only visible to the PowerBIAdmin report.)

.PARAMETER ClientScope
Share-safe scope run for a Fabric administrator of a customer tenant: only the admin APIs and an Entra ID lookup, and
only what a migration scope needs. Writes adminWorkspaces.csv (type, state, capacity, number of users, workspace
admins, item counts, last change seen, active users, data source types, gateway use), adminCapacities.csv,
adminDatasourceInstances.csv without addresses, repointingMap.csv with host names only, capacityFootprint.csv,
principals.csv (groups, apps and admins named, other people as counts), workspaceScopeSummary.csv, workspacePlan.csv
(with -WorkspaceListFile), findings and coverage. The activity log is counted, not written. Does not collect DAX, M,
model schemas, code, per-item users, item names, database names or paths, tenant settings or sharing links, and
writes no raw Scanner API answer. -Reports, -AddReports, -SkipReports and -ReportProfile are ignored; -Zip is on.

.PARAMETER OutputPath
Folder in which the run folder powerbi2csv_<yyyyMMdd-HHmmss>_<tenant> is created. Default: the current folder.

.PARAMETER ReportProfile
Quick, Standard (default), Full, Migration, Security or Governance. Run -ListReports to see what each one contains.

.PARAMETER Reports
Run exactly these reports (or groups) instead of a profile, for example -Reports PowerBI,Fabric

.PARAMETER AddReports
Reports to add to the profile, for example -AddReports PowerBIModels,PowerBIExport

.PARAMETER SkipReports
Reports to leave out of the profile, for example -SkipReports PowerBIAdmin

.PARAMETER ListReports
Show the report catalog and exit. Together with a selection it shows what that selection would run.

.PARAMETER PbiRefreshHistory
Number of refresh history entries to read per semantic model (0 = none). Default 20.

.PARAMETER PbiRunHistory
Number of runs (job instances) to read per pipeline and notebook (0 = none; then no Spark sessions either). Default 20.

.PARAMETER WorkspaceListFile
A .csv or .xlsx file with the workspaces you expect (first sheet, or the sheet named Workspaces; the workspace name
in the column Workspace or in the first column). workspacePlan.csv then says for each one whether this account sees
it, also under a slightly different spelling, lists the visible workspaces that are not on the list, and carries the
other columns of the file over (owner, purpose, to migrate, ...).

.PARAMETER PbiNotebookFormat
Format of downloaded notebooks: Source (default; the Fabric Git format, one notebook-content.py / .sql per notebook)
or Ipynb (Jupyter).

.PARAMETER PbiActivityDays
Days of the activity log to read with the PowerBIActivity report (1-28, the service keeps 28 days). Default 7.

.PARAMETER PbiExportMaxSizeMB
PowerBIExport: stop downloading .pbix/.rdl files once this many MB were downloaded (0 = no limit). Default 1024.
Item definitions (TMDL, PBIR, notebooks, pipelines) are small and always exported.

.PARAMETER Zip
Also create <run folder>.zip next to the run folder.

.PARAMETER NoExcel
Do not create the Excel workbook (it is only created when the ImportExcel module is installed).

.PARAMETER UseDeviceAuthentication
Sign in with a device code (for sessions without a browser).

.PARAMETER PropertiesMaxChars
Longest text kept in one CSV cell (Excel limit is 32767). Default 32000.

.PARAMETER Delimiter
CSV delimiter: comma (default), semicolon or tab.

.EXAMPLE
.\powerbi2csv.ps1
Standard inventory of every workspace the account can see.

.EXAMPLE
.\powerbi2csv.ps1 -ListReports
Shows the reports, the profiles they belong to and the files they write.

.EXAMPLE
.\powerbi2csv.ps1 -ReportProfile Full -Workspace 'Finance*' -Zip
Everything, including model tables/measures and the activity log, for the workspaces whose name starts with Finance.

.EXAMPLE
.\powerbi2csv.ps1 -ClientScope -WorkspaceListFile .\scope.xlsx
As a Fabric administrator of a customer tenant: the migration scope only, safe to send to a consultant.

.EXAMPLE
.\powerbi2csv.ps1 -AddReports PowerBIExport -Workspace 'Sales'
Standard inventory plus a download of the .pbix files and item definitions of the workspace Sales.

.LINK
https://github.com/AleksandrReznik/powerbi2csv

.LINK
https://it4it.solutions/2026/10/23/powerbi2csv-power-bi-fabric-tenant-inventory/

.LINK
https://learn.microsoft.com/rest/api/power-bi/

.LINK
https://learn.microsoft.com/rest/api/fabric/
#>

# Report names are not a ValidateSet: they are checked in New-A2COptions, so a comma list also works when the script
# is started with pwsh -File or from cmd.
[CmdletBinding(PositionalBinding = $false)]
param(
    [string] $TenantId,

    [string[]] $Workspace,

    [string[]] $ExcludeWorkspace,

    [switch] $PbiIncludePersonalWorkspaces,

    [switch] $ClientScope,

    [string] $OutputPath,

    [Alias('Profile')]
    [ValidateSet('Quick', 'Standard', 'Full', 'Migration', 'Security', 'Governance')]
    [string] $ReportProfile = 'Standard',

    [ArgumentCompleter({ $a2cWord = [string]$args[2]; @('All', 'Analytics', 'PowerBI', 'Fabric', 'FabricSql', 'Dependencies', 'FabricCode', 'FabricStorage', 'PowerAutomate', 'PowerBIModels', 'PowerBIAdmin', 'PowerBIActivity', 'Principals', 'PowerBIExport') | Where-Object { $_ -like ($a2cWord + '*') } })]
    [string[]] $Reports,

    [ArgumentCompleter({ $a2cWord = [string]$args[2]; @('All', 'Analytics', 'PowerBI', 'Fabric', 'FabricSql', 'Dependencies', 'FabricCode', 'FabricStorage', 'PowerAutomate', 'PowerBIModels', 'PowerBIAdmin', 'PowerBIActivity', 'Principals', 'PowerBIExport') | Where-Object { $_ -like ($a2cWord + '*') } })]
    [string[]] $AddReports,

    [ArgumentCompleter({ $a2cWord = [string]$args[2]; @('All', 'Analytics', 'PowerBI', 'Fabric', 'FabricSql', 'Dependencies', 'FabricCode', 'FabricStorage', 'PowerAutomate', 'PowerBIModels', 'PowerBIAdmin', 'PowerBIActivity', 'Principals', 'PowerBIExport') | Where-Object { $_ -like ($a2cWord + '*') } })]
    [string[]] $SkipReports,

    [switch] $ListReports,

    [ValidateRange(0, 1000)]
    [int] $PbiRefreshHistory = 20,

    [ValidateRange(0, 100)]
    [int] $PbiRunHistory = 20,

    [string] $WorkspaceListFile,

    [ValidateSet('Source', 'Ipynb')]
    [string] $PbiNotebookFormat = 'Source',

    [ValidateRange(1, 28)]
    [int] $PbiActivityDays = 7,

    [ValidateRange(0, 1048576)]
    [int] $PbiExportMaxSizeMB = 1024,

    [switch] $Zip,
    [switch] $NoExcel,
    [switch] $UseDeviceAuthentication,

    [ValidateRange(1000, 32000)]
    [int] $PropertiesMaxChars = 32000,

    [ValidateSet(',', ';', "`t")]
    [string] $Delimiter = ',',

    # words of a list that cmd or pwsh -File passed separately (see Resolve-A2CListArguments)
    [Parameter(ValueFromRemainingArguments = $true, DontShow = $true)]
    [string[]] $RemainingArguments
)

#region ============================== TABLE OF CONTENTS ===============================================
#  1. Constants and catalogs .... report catalog, profiles, datasets, branding
#  2. Utilities ................. logging, progress, JSON, dates, CSV writer, coverage      (as in azure2csv.ps1)
#  3. Transport ................. Invoke-A2CHttp (single I/O seam), retry, Power BI and Fabric REST, SQL endpoint
#  4. Context and scope ......... sign-in                                                    (as in azure2csv.ps1)
#  5. Datasets (collectors) ..... one Get-A2CDataset* function per dataset
#  7. Report builders ........... one New-A2CReport* function per report, schemas next to them
#  8. Findings .................. rules evaluated on collected data
#  9. Output .................... Excel, zip, <yyyyMMdd-HHmmss>__run.json
# 10. Main ...................... orchestration; skipped when the file is dot-sourced (tests)
# Region numbers match azure2csv.ps1 (there is no region 6 here). Blocks between "# ---- Power BI" and
# "# ---- end Power BI" are the ones to paste into the same region of azure2csv.ps1.
#endregion

Set-StrictMode -Version 1.0

#region ============================== 1. CONSTANTS AND CATALOGS ======================================

$script:A2CVersion = '1.0.0'
$script:A2CBrand = [ordered]@{
    Tool    = 'powerbi2csv'
    Author  = 'Aleksandr Reznik'
    Site    = 'https://it4it.solutions'
    Email   = 'aleksandr@reznik.lt'
    Repo    = 'https://github.com/AleksandrReznik/powerbi2csv'
    Tagline = 'Need help with a Power BI / Fabric or Azure tenant-to-tenant migration? Get in touch.'
}

# ---- Report catalog: the single source of truth for -ListReports, profiles, selection and README --------
# Needs = datasets (see $A2CDatasetCatalog). Speed: Fast = a few calls per tenant, Medium = a few calls per
# workspace, Slow = one or more calls per item.
$script:A2CReportCatalog = @(
    # ---- Power BI ----
    @{ Name = 'PowerBI';         Group = 'Analytics'; Speed = 'Slow';   OptIn = $false; Perm = 'Workspace member (more with Contributor/Admin)'
       Files = 'PowerBI\workspaces.csv, workspaceUsers.csv, reports.csv, reportPages.csv, reportDatasources.csv, semanticModels.csv, semanticModelDatasources.csv, semanticModelUsers.csv, semanticModelParameters.csv, refreshSchedules.csv, refreshHistory.csv, dashboards.csv, dashboardTiles.csv, dataflows.csv, dataflowDatasources.csv, dataflowTransactions.csv, lineage.csv, apps.csv, appContent.csv, capacities.csv, capacityWorkloads.csv, refreshables.csv, gateways.csv, gatewayDatasources.csv, gatewayDatasourceUsers.csv, deploymentPipelines.csv, pipelineStages.csv, pipelineStageItems.csv, pipelineUsers.csv, pipelineOperations.csv, availableFeatures.csv'
       Needs = @('PbiWorkspaces', 'PowerBI')
       Desc  = 'Power BI content of every workspace the account can see: reports, semantic models, data sources, refreshes, dashboards, dataflows, apps, capacities, gateways, deployment pipelines' }
    @{ Name = 'Fabric';          Group = 'Analytics'; Speed = 'Medium'; OptIn = $false; Perm = 'Workspace member'
       Files = 'PowerBI\workspaces.csv, fabricItems.csv, workspaceRoleAssignments.csv, workspaceGit.csv, workspaceSparkSettings.csv, sparkPools.csv, managedPrivateEndpoints.csv, lakehouses.csv, lakehouseTables.csv, warehouses.csv, shortcuts.csv, itemSchedules.csv, itemRuns.csv, sparkSessions.csv, itemConnections.csv, environmentMatrix.csv, folders.csv, workspaceNetwork.csv, gitItemStatus.csv, onelakeRoles.csv, deploymentStages.csv, deploymentStageItems.csv, deploymentRoleAssignments.csv, domains.csv, connections.csv, capacities.csv, gateways.csv'
       Needs = @('PbiWorkspaces', 'Fabric')
       Desc  = 'Fabric items of every workspace (lakehouses, warehouses, notebooks, pipelines, ...), folders, role assignments, OneLake security roles, Git connection and item status, network policy, Spark settings and sessions, shortcuts, connections, run history, deployment pipelines, DEV/TEST/PROD comparison' }
    @{ Name = 'FabricSql';       Group = 'Analytics'; Speed = 'Medium'; OptIn = $false; Perm = 'Read on the lakehouse or warehouse; outbound TCP 1433'
       Files = 'PowerBI\sqlObjects.csv, sqlColumns.csv, sqlModules.csv, sqlReferences.csv (+ in the Fabric files: warehouses.csv counts, lakehouses.csv Views/StoredProcedures/Functions, tables of lakehouses with schemas in lakehouseTables.csv)'
       Needs = @('PbiWorkspaces', 'Fabric', 'FabricSql')
       Desc  = 'Tables, views, stored procedures and functions of every lakehouse SQL endpoint and warehouse, their columns, the code of views and procedures and references between databases (catalog queries over the SQL endpoint)' }
    @{ Name = 'Dependencies';    Group = 'Analytics'; Speed = 'Slow';   OptIn = $false; Perm = 'Contributor on the workspace (item definitions)'
       Files = 'PowerBI\pipelineActivities.csv, notebookDependencies.csv, dataflowSources.csv, modelPartitions.csv, modelExpressions.csv, modelRoles.csv, modelSources.csv, itemDependencies.csv, externalDependencies.csv (+ StorageModes/Measures/RlsRoles in semanticModels.csv)'
       Needs = @('PbiWorkspaces', 'PowerBI', 'Fabric', 'ItemDefinitions')
       Desc  = 'What runs and reads what, from the item definitions: pipeline activities, notebook lakehouses/tables/calls, dataflow sources and destinations, model partitions, DAX and M, roles; all item-to-item links and every external target (SharePoint, SQL, storage, ...)' }
    @{ Name = 'FabricCode';      Group = 'Analytics'; Speed = 'Slow';   OptIn = $true;  Perm = 'Contributor on the workspace (item definitions)'
       Files = 'PowerBI\exportIndex.csv, PowerBI\<workspace>\Notebook\<notebook>\notebook-content.py, PowerBI\<workspace>\<item type>\<item>\* (pipelines, dataflows, semantic models as TMDL, ...), PowerBI\<workspace>\dataflows\*.json, PowerBI\<workspace>\sql\<database>\*.sql'
       Needs = @('PbiWorkspaces', 'PowerBI', 'Fabric', 'FabricCode')
       Desc  = 'Download of the code: notebooks, pipelines, dataflows, semantic model definitions and every other item type with a definition, plus the SQL of views and procedures when FabricSql runs. No report files and no data (that is PowerBIExport)' }
    @{ Name = 'FabricStorage';   Group = 'Analytics'; Speed = 'Slow';   OptIn = $true;  Perm = 'Read on the lakehouse or warehouse (OneLake)'
       Files = 'PowerBI\onelakeStorage.csv, onelakeTableSizes.csv'
       Needs = @('PbiWorkspaces', 'Fabric', 'FabricStorage')
       Desc  = 'OneLake size per lakehouse and warehouse table and of the Files area (file listing only; includes old Delta versions)' }
    @{ Name = 'PowerAutomate';   Group = 'Analytics'; Speed = 'Medium'; OptIn = $true;  Perm = 'Flow owner or co-owner; Power Platform administrator for all flows'
       Files = 'PowerBI\flowEnvironments.csv, flows.csv'
       Needs = @('PowerAutomate')
       Desc  = 'Power Automate environments and cloud flows with their triggers and connectors (flows often start refreshes and pipelines)' }
    @{ Name = 'PowerBIModels';   Group = 'Analytics'; Speed = 'Slow';   OptIn = $true;  Perm = 'Build permission on the model; tenant setting "Semantic model Execute Queries REST API"'
       Files = 'PowerBI\modelTables.csv, modelColumns.csv, modelMeasures.csv, modelRelationships.csv'
       Needs = @('PbiWorkspaces', 'PowerBI', 'PowerBIModels')
       Desc  = 'Tables, columns, measures and relationships of every semantic model (DAX INFO.VIEW queries, 4 per model). The service does not return the DAX expressions this way: PowerBIExport (TMDL files) and PowerBIAdmin (adminModelExpressions.csv) have them' }
    @{ Name = 'PowerBIAdmin';    Group = 'Analytics'; Speed = 'Medium'; OptIn = $false; Perm = 'Fabric administrator'
       Files = 'PowerBI\adminWorkspaces.csv, workspaceScopeSummary.csv, adminWorkspaceUsers.csv, adminItems.csv, adminItemUsers.csv, adminDatasourceInstances.csv, adminModelTables.csv, adminModelExpressions.csv, tenantSettings.csv, adminCapacities.csv, adminApps.csv, widelyShared.csv, adminDomains.csv, scan\scanResult_*.json'
       Needs = @('PowerBIAdmin')
       Desc  = 'Whole tenant with the admin APIs: all workspaces (also personal and orphaned), access per item, data source lineage, model expressions (Scanner API), tenant settings. Skipped without the role' }
    @{ Name = 'PowerBIActivity'; Group = 'Analytics'; Speed = 'Slow';   OptIn = $true;  Perm = 'Fabric administrator'
       Files = 'PowerBI\activityEvents.csv'
       Needs = @('PowerBIActivity')
       Desc  = 'Activity log (who viewed, exported, changed or shared what) of the last -PbiActivityDays days' }
    @{ Name = 'Principals';      Group = 'Analytics'; Speed = 'Medium'; OptIn = $false; Perm = 'any; the Entra ID lookup needs directory read, which members have by default'
       Files = 'PowerBI\principals.csv (+ principal counts in workspaceScopeSummary.csv, departed admins in adminWorkspaces.csv)'
       Needs = @('Principals')
       Desc  = 'Every user, guest, group, service principal and managed identity seen in workspaces, items, capacities, gateways and deployment pipelines, de-duplicated, with roles, workspace count and whether the account still exists in Entra ID' }
    @{ Name = 'PowerBIExport';   Group = 'Analytics'; Speed = 'Slow';   OptIn = $true;  Explicit = $true; Perm = 'Contributor on the workspace'
       Files = 'PowerBI\exportIndex.csv, PowerBI\<workspace>\reports\*.pbix, PowerBI\<workspace>\<item type>\<item>\* (definitions), PowerBI\<workspace>\dataflows\*.json, PowerBI\<workspace>\sql\<database>\*.sql'
       Needs = @('PbiWorkspaces', 'PowerBI', 'Fabric', 'PowerBIExport')
       Desc  = 'Download of report files (.pbix/.rdl), item definitions (TMDL, PBIR, notebooks, pipelines, ...) and dataflow definitions' }
    # ---- end Power BI ----
)

$script:A2CReportGroups = @('Analytics')

$script:A2CProfiles = [ordered]@{
    Quick      = @('Fabric')
    Standard   = @('PowerBI', 'Fabric', 'FabricSql', 'Dependencies', 'PowerBIAdmin', 'Principals')
    Full       = @('All')
    Migration  = @('PowerBI', 'Fabric', 'FabricSql', 'Dependencies', 'FabricCode', 'FabricStorage', 'PowerAutomate', 'PowerBIAdmin', 'PowerBIModels', 'Principals')
    Security   = @('PowerBI', 'Fabric', 'PowerBIAdmin', 'PowerBIActivity', 'Principals')
    Governance = @('PowerBI', 'Fabric', 'FabricSql', 'Dependencies', 'PowerBIAdmin', 'Principals')
}

# ---- Datasets: what is collected. Order = execution order (dependencies first). -----------------------
$script:A2CDatasetCatalog = [ordered]@{
    # ---- Power BI ----
    PbiWorkspaces   = @{ Needs = @();                          Desc = 'Workspaces in scope (Power BI and Fabric REST with the Az sign-in)' }
    PowerBI         = @{ Needs = @('PbiWorkspaces');           Desc = 'Power BI content per workspace, apps, capacities, gateways, deployment pipelines (Power BI REST)' }
    Fabric          = @{ Needs = @('PbiWorkspaces');           Desc = 'Fabric items and workspace settings (Fabric REST)' }
    FabricSql       = @{ Needs = @('PbiWorkspaces', 'Fabric');  Desc = 'Objects of lakehouse SQL endpoints and warehouses (catalog query over the SQL endpoint)' }
    ItemDefinitions = @{ Needs = @('PbiWorkspaces', 'PowerBI', 'Fabric'); Desc = 'Definitions of pipelines, notebooks, dataflows and semantic models (getDefinition, 20 at a time)' }
    FabricCode      = @{ Needs = @('PbiWorkspaces', 'PowerBI', 'Fabric'); Desc = 'Item definitions and SQL code written to files (download)' }
    FabricStorage   = @{ Needs = @('PbiWorkspaces', 'Fabric');  Desc = 'OneLake file listing of lakehouse and warehouse tables' }
    PowerAutomate   = @{ Needs = @();                          Desc = 'Power Automate environments and flows (Power Automate REST)' }
    PowerBIModels   = @{ Needs = @('PbiWorkspaces', 'PowerBI'); Desc = 'Semantic model schema (DAX INFO.VIEW queries, 4 calls per model)' }
    PowerBIAdmin    = @{ Needs = @();                          Desc = 'Tenant-wide metadata (admin APIs, Scanner API)' }
    PowerBIActivity = @{ Needs = @();                          Desc = 'Activity log (admin API, one or more calls per day)' }
    Principals      = @{ Needs = @();                          Desc = 'Principals of the collected data, looked up in Entra ID (Microsoft Graph); runs after the other datasets' }
    PowerBIExport   = @{ Needs = @('PbiWorkspaces', 'PowerBI', 'Fabric'); Desc = 'Report files and item definitions (download)' }
    # ---- end Power BI ----
}

#endregion

#region ============================== 2. UTILITIES ====================================================

$script:A2C = $null   # run context (created by New-A2CRunContext)

function New-A2CRunContext {
    param([Parameter(Mandatory)] $Options)
    $script:A2C = [pscustomobject]@{
        Options        = $Options
        Now            = $Options.Now
        StartedUtc     = [DateTime]::UtcNow
        OutDir         = ''
        RunName        = ''
        RunStamp       = ''                        # yyyyMMdd-HHmmss of the run folder (names the general log)
        TenantId       = ''
        Account        = ''
        Log            = [System.Collections.Generic.List[string]]::new()
        LogPath        = ''                        # general log file, appended line by line once the run folder exists
        Coverage       = [System.Collections.Generic.List[object]]::new()
        Files          = [System.Collections.Generic.List[object]]::new()
        Out            = @{}                       # file name -> List of rows (kept for HTML/Excel/findings)
        Data           = @{}                       # dataset name -> collected data
        DatasetStatus  = @{}                       # dataset name -> Collected/Partial/Skipped/Failed
        Selected       = @()                       # selected report names
        Datasets       = @()                       # datasets to collect
        Stages         = [System.Collections.Generic.List[object]]::new()
        Stats          = @{ pbi = 0; fabric = 0; sql = 0; onelake = 0; flow = 0; graph = 0; Retries = 0; Throttled = 0; SleepSec = 0.0 }
        Tokens         = @{}                       # resource -> access token cache (Get-A2CBearerToken: Power BI, Fabric)
        SignInFailed   = @{}                       # API kind -> sign-in error (later calls of that kind fail at once)
        SignInRetried  = @{}                       # API kind -> $true once the user was asked to sign in again
        Warnings       = [System.Collections.Generic.List[string]]::new()
        Errors         = [System.Collections.Generic.List[string]]::new()
        Findings       = [System.Collections.Generic.List[object]]::new()
        Env            = @{}
        SchemaWarnings = @{}
    }
    $script:A2CPropCache = @{}
    return $script:A2C
}

$script:A2CUtf8Bom = New-Object System.Text.UTF8Encoding $true

function Add-A2CLogLine {
    # Adds a line to the run log (not to the screen). Once the run folder exists every line is also appended to the
    # general log file at once, so the file is complete even when the window closes or the run stops on an error.
    param([string] $Level, [string] $Message)
    if ($null -eq $script:A2C) { return }
    $line = '[{0}] {1,-7} {2}' -f ([DateTime]::Now.ToString('HH:mm:ss')), $Level.ToUpperInvariant(), $Message
    $script:A2C.Log.Add($line)
    if ($script:A2C.LogPath) {
        try { [IO.File]::AppendAllText($script:A2C.LogPath, $line + "`r`n", $script:A2CUtf8Bom) }
        catch {
            $script:A2C.LogPath = ''
            Write-Host ('   WARNING: the general log can no longer be written: ' + $_.Exception.Message) -ForegroundColor Yellow
        }
    }
}

function Get-A2CStackText {
    # Where an error happened, for the log: "Func, azure2csv.ps1: line 12 <- Caller, azure2csv.ps1: line 34" (no folders).
    param($ErrorRecord)
    $s = [string]$ErrorRecord.ScriptStackTrace
    if (-not $s) { return '' }
    $frames = @($s -split "`r?`n") | Where-Object { $_ } | Select-Object -First 6 |
        ForEach-Object { ($_ -replace '^at ', '') -replace ', [^,]*[\\/]([^\\/,]+): line', ', $1: line' }
    return (@($frames) -join ' <- ')
}

function Write-A2CLog {
    param(
        [Parameter(Position = 0)] [string] $Message,
        [ValidateSet('Info', 'Step', 'Ok', 'Notice', 'Warn', 'Error', 'Verbose', 'Detail')] [string] $Level = 'Info'
    )
    # Notice = shown in yellow but not counted as a warning (e.g. a retry that may still succeed)
    Add-A2CLogLine -Level $Level -Message $Message
    switch ($Level) {
        'Verbose' { Write-Verbose $Message; return }
        'Step'    { Write-Host ''; Write-Host ('== ' + $Message) -ForegroundColor Cyan }
        'Ok'      { Write-Host ('   ' + $Message) -ForegroundColor Green }
        'Notice'  { Write-Host ('   ' + $Message) -ForegroundColor Yellow }
        'Warn'    { Write-Host ('   WARNING: ' + $Message) -ForegroundColor Yellow; if ($null -ne $script:A2C) { $script:A2C.Warnings.Add($Message) } }
        'Error'   { Write-Host ('   ERROR: ' + $Message) -ForegroundColor Red; if ($null -ne $script:A2C) { $script:A2C.Errors.Add($Message) } }
        'Detail'  { Write-Host ('     ' + $Message) -ForegroundColor DarkGray }
        default   { Write-Host ('   ' + $Message) }
    }
}

function Write-A2CProgress {
    param([string] $Activity, [int] $Current, [int] $Total, [datetime] $Started, [string] $Item = '', [switch] $Completed)
    if ($Completed) { Write-Progress -Activity $Activity -Completed; return }
    if ($Total -le 0) { return }
    $pct = [math]::Min(100, [int](100 * $Current / $Total))
    $status = '{0} of {1}' -f $Current, $Total
    $secs = -1
    if ($Current -gt 0) {
        $elapsed = ([DateTime]::Now - $Started).TotalSeconds
        $secs = [int](($elapsed / $Current) * ($Total - $Current))
    }
    if ($Item) { $status = $status + ' - ' + $Item }
    Write-Progress -Activity $Activity -Status $status -PercentComplete $pct -SecondsRemaining $secs
}

function Start-A2CSleep {
    param([double] $Seconds)
    if ($Seconds -le 0) { return }
    if ($null -ne $script:A2C) { $script:A2C.Stats.SleepSec += $Seconds }
    Start-Sleep -Milliseconds ([int]($Seconds * 1000))
}

function Add-A2CCoverage {
    param(
        [string] $Area,
        [ValidateSet('Collected', 'Partial', 'Skipped', 'Failed', 'NotSelected', 'Info')] [string] $Status,
        [string] $Detail = '',
        [int] $Items = -1
    )
    $script:A2C.Coverage.Add([ordered]@{ Area = $Area; Status = $Status; Items = $(if ($Items -ge 0) { $Items } else { $null }); Detail = $Detail })
}

# ---- JSON --------------------------------------------------------------------------------------------
$script:A2CJss = $null

function ConvertTo-A2CCaseInsensitive {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in $Value.Keys) { if ([string]$k -ne '') { $h[[string]$k] = ConvertTo-A2CCaseInsensitive $Value[$k] } }
        return [pscustomobject]$h
    }
    if ($Value -is [System.Collections.IList] -and -not ($Value -is [string])) {
        $l = [System.Collections.Generic.List[object]]::new()
        foreach ($x in $Value) { $l.Add((ConvertTo-A2CCaseInsensitive $x)) }
        return , $l.ToArray()
    }
    return $Value
}

function ConvertFrom-A2CJson {
    param([string] $Json)
    if ([string]::IsNullOrWhiteSpace($Json)) { return $null }
    try {
        return (ConvertFrom-Json -InputObject $Json -ErrorAction Stop)
    }
    catch {
        # Windows PowerShell 5.1 refuses keys that differ only by case; fall back to JavaScriptSerializer.
        if ($PSVersionTable.PSEdition -eq 'Core') { throw }
        if ($null -eq $script:A2CJss) {
            Add-Type -AssemblyName System.Web.Extensions
            $script:A2CJss = New-Object System.Web.Script.Serialization.JavaScriptSerializer
            $script:A2CJss.MaxJsonLength = [int]::MaxValue
            $script:A2CJss.RecursionLimit = 1000
        }
        return (ConvertTo-A2CCaseInsensitive ($script:A2CJss.DeserializeObject($Json)))
    }
}

function ConvertTo-A2CJsonString {
    # Compact JSON for CSV cells; never throws.
    param($Value, [int] $Depth = 20)
    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value }
    try { return (ConvertTo-Json -InputObject $Value -Depth $Depth -Compress) } catch { return [string]$Value }
}

function ConvertTo-A2CJsonLiteral {
    # JSON string literal (with quotes) for building request bodies.
    param([string] $Value)
    if ($null -eq $Value) { return 'null' }
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('"')
    foreach ($ch in $Value.ToCharArray()) {
        switch ($ch) {
            '"'  { [void]$sb.Append('\"') }
            '\'  { [void]$sb.Append('\\') }
            "`n" { [void]$sb.Append('\n') }
            "`r" { [void]$sb.Append('\r') }
            "`t" { [void]$sb.Append('\t') }
            default {
                if ([int]$ch -lt 32) { [void]$sb.Append(('\u{0:x4}' -f [int]$ch)) } else { [void]$sb.Append($ch) }
            }
        }
    }
    [void]$sb.Append('"')
    return $sb.ToString()
}

# ---- Dates -------------------------------------------------------------------------------------------
function ConvertTo-A2CUtc {
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [DateTime]) {
        if ($Value.Kind -eq [DateTimeKind]::Unspecified) { return [DateTime]::SpecifyKind($Value, [DateTimeKind]::Utc) }
        return $Value.ToUniversalTime()
    }
    if ($Value -is [DateTimeOffset]) { return $Value.UtcDateTime }
    $s = [string]$Value
    if ($s -eq '') { return $null }
    $dto = [DateTimeOffset]::MinValue
    if ([DateTimeOffset]::TryParse($s, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::AssumeUniversal, [ref]$dto)) {
        return $dto.UtcDateTime
    }
    return $null
}

function Format-A2CDate {
    param($Value, [switch] $DateOnly)
    if ($null -eq $Value) { return '' }
    $dt = ConvertTo-A2CUtc $Value
    if ($null -eq $dt) { return [string]$Value }
    if ($DateOnly) { return $dt.ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture) }
    return $dt.ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
}

function Get-A2CDaysUntil {
    param($Value)
    $dt = ConvertTo-A2CUtc $Value
    if ($null -eq $dt) { return $null }
    return [int][math]::Floor(($dt - $script:A2C.Now).TotalDays)
}

function ConvertTo-A2CSeconds {
    param($Value)
    if ($null -eq $Value) { return 0 }
    $s = ([string]$Value).Trim()
    $d = 0.0
    if ([double]::TryParse($s, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return $d }
    $ts = [TimeSpan]::Zero
    if ([TimeSpan]::TryParse($s, [Globalization.CultureInfo]::InvariantCulture, [ref]$ts)) { return $ts.TotalSeconds }
    $dt = ConvertTo-A2CUtc $s
    if ($null -ne $dt) { return [math]::Max(0, ($dt - [DateTime]::UtcNow).TotalSeconds) }
    return 0
}

# ---- Strings, lists and ARM IDs ----------------------------------------------------------------------
function Join-A2C {
    param($Items, [string] $Separator = '; ')
    if ($null -eq $Items) { return '' }
    if ($Items -is [string]) { return $Items }
    $l = [System.Collections.Generic.List[string]]::new()
    foreach ($i in $Items) {
        if ($null -eq $i) { continue }
        $s = [string]$i
        if ($s -ne '') { $l.Add($s) }
    }
    return [string]::Join($Separator, $l)
}

function Get-A2CUnique {
    # Distinct non-empty strings, first occurrence order, case-insensitive.
    param($Items)
    $seen = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $l = [System.Collections.Generic.List[string]]::new()
    foreach ($i in $Items) {
        if ($null -eq $i) { continue }
        $s = [string]$i
        if ($s -ne '' -and $seen.Add($s)) { $l.Add($s) }
    }
    return , $l.ToArray()
}

function Get-A2CCount {
    param($Items)
    if ($null -eq $Items) { return 0 }
    if ($Items -is [string]) { return 1 }
    return @($Items).Count
}

# ---- Report rows and CSV writer ----------------------------------------------------------------------
$script:A2CSchemas = @{}

function Register-A2CSchema {
    # Spec: space separated Column:kind tokens; kind drives -Mask (default p = keep).
    param([string] $File, [string] $Spec)
    $cols = [System.Collections.Generic.List[object]]::new()
    foreach ($tok in ($Spec -split '\s+')) {
        if (-not $tok) { continue }
        $i = $tok.LastIndexOf(':')
        if ($i -gt 0) { $n = $tok.Substring(0, $i); $k = $tok.Substring($i + 1) } else { $n = $tok; $k = 'p' }
        $cols.Add([pscustomobject]@{ Name = $n; Kind = $k })
    }
    $script:A2CSchemas[$File] = $cols
}

function Format-A2CCell {
    param($Value)
    if ($null -eq $Value) { return '' }
    if ($Value -is [string]) { return $Value }
    if ($Value -is [bool]) { if ($Value) { return 'True' } else { return 'False' } }
    if ($Value -is [DateTime] -or $Value -is [DateTimeOffset]) { return (Format-A2CDate $Value) }
    if ($Value -is [System.Collections.IDictionary] -or $Value -is [System.Management.Automation.PSCustomObject]) {
        return (ConvertTo-A2CJsonString $Value)
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        $l = [System.Collections.Generic.List[string]]::new()
        foreach ($x in $Value) { $l.Add((Format-A2CCell $x)) }
        return [string]::Join('; ', $l)
    }
    if ($Value -is [IFormattable]) { return $Value.ToString($null, [Globalization.CultureInfo]::InvariantCulture) }
    return [string]$Value
}

function ConvertTo-A2CCsvField {
    param([string] $Text, [int] $Max)
    $s = $Text
    if ($null -eq $s) { $s = '' }
    if ($s.Length -gt $Max) { $s = $s.Substring(0, $Max - 20) + '...[truncated]' }
    if ($s.IndexOf([char]10) -ge 0 -or $s.IndexOf([char]13) -ge 0) { $s = ($s -replace "`r`n|`r|`n", ' ') }
    if ($s.Length -gt 0) {
        $c = $s[0]
        if ($c -eq '=' -or $c -eq '+' -or $c -eq '-' -or $c -eq '@' -or $c -eq [char]9) {
            $d = 0.0
            if (-not [double]::TryParse($s, [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { $s = "'" + $s }
        }
    }
    return '"' + $s.Replace('"', '""') + '"'
}

function Add-A2CRows {
    # Store report rows (ordered dictionaries) for writing and later use (HTML/Excel/findings).
    param([string] $File, $Rows)
    if (-not $script:A2C.Out.ContainsKey($File)) { $script:A2C.Out[$File] = [System.Collections.Generic.List[object]]::new() }
    foreach ($r in $Rows) { if ($null -ne $r) { $script:A2C.Out[$File].Add($r) } }
}

$script:A2CSplitBy = @{}   # file -> column: one file per value, in a subfolder named after it (ADO/<organization>/projects.csv)

function Get-A2CSafePathSegment {
    # A file or folder name that cannot escape its folder ("..") and is valid on Windows.
    param([string] $Name)
    $s = (Get-A2CSafeFileName $Name).TrimEnd('.', ' ')
    if (-not $s -or $s -match '^\.+$') { return '_' }
    return $s
}

function Get-A2CSplitFolders {
    # Folders that get every split file (header only when empty): each collected Azure DevOps organization.
    param([string] $File)
    if ($File -notlike 'ADO/*') { return @() }
    return @((Get-A2COutputRows 'ADO/organizations.csv') | Where-Object { $_['Status'] -eq 'Collected' } | ForEach-Object { [string]$_['Organization'] })
}

function Save-A2CCsv {
    # Writes a registered report file (UTF-8 with BOM, all fields quoted, formula-injection guard). A file in
    # $A2CSplitBy is written once per value of its column, into a subfolder named after the (masked) value.
    param([string] $File)
    $schema = $script:A2CSchemas[$File]
    if ($null -eq $schema) { throw "No schema registered for $File" }
    $rows = Get-A2COutputRows $File
    $by = $script:A2CSplitBy[$File]
    if (-not $by) { Write-A2CCsvFile -File $File -Schema $schema -Rows $rows -SchemaKey $File; return }
    $groups = [ordered]@{}
    foreach ($n in (Get-A2CSplitFolders $File)) { $groups[$n] = [System.Collections.Generic.List[object]]::new() }
    foreach ($row in $rows) {
        $k = [string]$row[$by]
        if (-not $groups.Contains($k)) { $groups[$k] = [System.Collections.Generic.List[object]]::new() }
        $groups[$k].Add($row)
    }
    $dir = [IO.Path]::GetDirectoryName($File)
    $leaf = [IO.Path]::GetFileName($File)
    foreach ($k in $groups.Keys) {
        Write-A2CCsvFile -File ($dir + '/' + (Get-A2CSafePathSegment $k) + '/' + $leaf) -Schema $schema -Rows $groups[$k] -SchemaKey $File
    }
}

function Write-A2CCsvFile {
    param([string] $File, $Schema, $Rows, [string] $SchemaKey)
    $schema = $Schema
    $rows = $Rows
    $max = [int]$script:A2C.Options.PropertiesMaxChars
    $sep = [string]$script:A2C.Options.Delimiter
    $path = [IO.Path]::Combine($script:A2C.OutDir, $File)
    $dir = [IO.Path]::GetDirectoryName($path)
    if (-not [IO.Directory]::Exists($dir)) { [void][IO.Directory]::CreateDirectory($dir) }   # e.g. ADO/
    $colNames = @($schema | ForEach-Object { $_.Name })
    $known = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($n in $colNames) { [void]$known.Add($n) }
    $sw = New-Object System.IO.StreamWriter($path, $false, (New-Object System.Text.UTF8Encoding $true))
    try {
        $sw.Write([string]::Join($sep, @($colNames | ForEach-Object { ConvertTo-A2CCsvField -Text $_ -Max $max })))
        $sw.Write("`r`n")
        foreach ($row in $rows) {
            $fields = [System.Collections.Generic.List[string]]::new()
            foreach ($n in $colNames) {
                $v = $null
                if ($row.Contains($n)) { $v = $row[$n] }
                $fields.Add((ConvertTo-A2CCsvField -Text (Format-A2CCell $v) -Max $max))
            }
            $sw.Write([string]::Join($sep, $fields))
            $sw.Write("`r`n")
            if (-not $script:A2C.SchemaWarnings.ContainsKey($SchemaKey)) {
                foreach ($k in $row.Keys) {
                    if (-not $known.Contains([string]$k)) {
                        $script:A2C.SchemaWarnings[$SchemaKey] = $true
                        Write-A2CLog -Level Verbose ('Schema of {0} has no column "{1}"' -f $SchemaKey, $k)
                        break
                    }
                }
            }
        }
    }
    finally { $sw.Dispose() }
    $script:A2C.Files.Add([pscustomobject]@{ File = $File; Rows = $rows.Count; Path = $path })
}

function Save-A2CText {
    param([string] $File, [string] $Text, [switch] $NoBom)
    $path = [IO.Path]::Combine($script:A2C.OutDir, $File)
    $enc = New-Object System.Text.UTF8Encoding (-not $NoBom)
    [IO.File]::WriteAllText($path, $Text, $enc)
    return $path
}

function Get-A2CRawRows {
    param([string] $File)
    $r = $script:A2C.Out[$File]
    if ($null -eq $r) { return , ([System.Collections.Generic.List[object]]::new()) }
    return , $r
}

function Get-A2CRows {
    # Report rows; while findings are evaluated under -Mask the masked rows are returned so details never leak names.
    param([string] $File)
    if ($script:A2C.Data['FindingsMasked']) { return , (Get-A2COutputRows $File) }
    return , (Get-A2CRawRows $File)
}

function Test-A2CSelected {
    param([string] $Report)
    return ($script:A2C.Selected -contains $Report)
}

function Test-A2CDataset {
    # True when a dataset was collected (fully or partially).
    param([string] $Name)
    $s = $script:A2C.DatasetStatus[$Name]
    return ($s -eq 'Collected' -or $s -eq 'Partial')
}

function Get-A2CSafeFileName {
    param([string] $Name)
    return (($Name -replace '[\\/:*?"<>|]', '_').Trim())
}

function Get-A2CFirst {
    # First non-empty value of several dotted property paths, e.g. 'dataSourceInfo.resourceID'.
    param($Object, [string[]] $Paths)
    if ($null -eq $Object) { return $null }
    foreach ($path in $Paths) {
        $cur = $Object
        foreach ($seg in $path.Split('.')) {
            if ($null -eq $cur) { break }
            $cur = $cur.$seg
        }
        if ($null -ne $cur -and [string]$cur -ne '') { return $cur }
    }
    return $null
}

function Save-A2CExportFile {
    # An exported file below the run folder: UTF-8 without BOM, as YAML and JSON tools expect; -Bom for PowerShell
    # scripts (Windows PowerShell 5.1 reads BOM-less non-ASCII scripts as ANSI).
    param([string] $File, [string] $Text, [switch] $Bom)
    $path = [IO.Path]::Combine($script:A2C.OutDir, $File)
    $dir = [IO.Path]::GetDirectoryName($path)
    if (-not [IO.Directory]::Exists($dir)) { [void][IO.Directory]::CreateDirectory($dir) }
    [IO.File]::WriteAllText($path, $Text, (New-Object System.Text.UTF8Encoding ([bool]$Bom)))
}

#endregion

#region ============================== 3. TRANSPORT =====================================================

function Invoke-A2CHttp {
    # The single network seam (tests replace this function). Never throws.
    # Returns @{ StatusCode = int (0 = transport error); Content = string; Headers = case-insensitive hashtable }
    # -OutFile (report downloads): the response body is written to that file instead of returned as Content
    param([string] $Method = 'GET', [Parameter(Mandatory)] [string] $Uri, [string] $Body, [string] $Kind = 'pbi', [string] $OutFile)
    # ---- Power BI ----
    if ($Kind -eq 'pbi') { return (Invoke-A2CTokenHttp -Resource $script:A2CPbiResource -Method $Method -Uri $Uri -Body $Body -OutFile $OutFile) }
    if ($Kind -eq 'fabric') { return (Invoke-A2CTokenHttp -Resource $script:A2CFabricResource -Method $Method -Uri $Uri -Body $Body -OutFile $OutFile) }
    if ($Kind -eq 'onelake') { return (Invoke-A2CTokenHttp -Resource $script:A2COneLakeResource -Method $Method -Uri $Uri -Body $Body -OutFile $OutFile) }
    if ($Kind -eq 'flow') { return (Invoke-A2CTokenHttp -Resource $script:A2CFlowResource -Method $Method -Uri $Uri -Body $Body -OutFile $OutFile) }
    if ($Kind -eq 'graph') { return (Invoke-A2CTokenHttp -Resource $script:A2CGraphResource -Method $Method -Uri $Uri -Body $Body -OutFile $OutFile) }
    # ---- end Power BI ----
    return @{ StatusCode = 0; Content = ('Unknown API kind: ' + $Kind); Headers = @{} }
}

function Get-A2CBearerToken {
    # Entra access token for a resource (Power BI, Fabric) from the Az sign-in - no app registration or secret -
    # cached until 5 minutes before it expires.
    param([string] $Resource)
    $c = $script:A2C.Tokens[$Resource]
    if ($null -ne $c -and $c.ExpiresOn -gt [DateTimeOffset]::UtcNow.AddMinutes(5)) { return $c.Token }
    $p = @{ ResourceUrl = $Resource; ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' }
    if ($script:A2C.TenantId) { $p['TenantId'] = $script:A2C.TenantId }
    if ((Get-Command Get-AzAccessToken).Parameters.ContainsKey('AsSecureString')) { $p['AsSecureString'] = $true }
    $t = Get-AzAccessToken @p
    $token = $t.Token
    if ($token -is [securestring]) { $token = [System.Net.NetworkCredential]::new('', $token).Password }
    $expires = [DateTimeOffset]::UtcNow.AddMinutes(30)
    if ($null -ne $t.ExpiresOn) { $expires = [DateTimeOffset]$t.ExpiresOn }
    $script:A2C.Tokens[$Resource] = [pscustomobject]@{ Token = [string]$token; ExpiresOn = $expires }
    return [string]$token
}

function Test-A2CHtmlFile {
    # A downloaded file that is an HTML page (a sign-in page instead of the zip or package).
    param([string] $Path)
    try {
        $fs = [IO.File]::OpenRead($Path)
        try { $buf = New-Object byte[] 256; $n = $fs.Read($buf, 0, 256) } finally { $fs.Dispose() }
        return ([Text.Encoding]::ASCII.GetString($buf, 0, $n) -match '^(?i)\s*<(!DOCTYPE html|html)')
    }
    catch { return $false }
}

function Invoke-A2CTokenHttp {
    # A request with an Entra bearer token (Azure DevOps, Kudu); part of the network seam (only Invoke-A2CHttp calls
    # it). -OutFile streams the body to a file (zip downloads). Never throws.
    param([string] $Resource, [switch] $AzureDevOps, [string] $Method, [string] $Uri, [string] $Body, [string] $OutFile)
    try { $token = Get-A2CBearerToken $Resource }
    catch { return @{ StatusCode = 0; Content = [string]$_.Exception.Message; Headers = @{} } }
    $p = @{ Uri = $Uri; Method = $Method; UseBasicParsing = $true; TimeoutSec = 120; ErrorAction = 'Stop'
        Headers = @{ Authorization = 'Bearer ' + $token; Accept = 'application/json' } }
    # Azure DevOps: without X-TFS-FedAuthRedirect a refused token is redirected to a sign-in page instead of HTTP 401
    if ($AzureDevOps) { $p['Headers']['X-TFS-FedAuthRedirect'] = 'Suppress' }
    if ($Body) { $p['Body'] = $Body; $p['ContentType'] = 'application/json' }
    if ($OutFile) {
        # a download (zip, package file): streamed to disk, so it never passes through a string. Package downloads can
        # redirect to blob storage; Invoke-WebRequest drops the Authorization header on such redirects.
        $p['OutFile'] = $OutFile
        $p['TimeoutSec'] = 3600
        $p['Headers']['Accept'] = 'application/zip, */*'
    }
    try {
        $r = Invoke-WebRequest @p
        if ($OutFile) {
            if ($AzureDevOps -and (Test-A2CHtmlFile $OutFile)) {
                Remove-Item -LiteralPath $OutFile -Force -ErrorAction SilentlyContinue
                return @{ StatusCode = 401; Content = 'Azure DevOps returned its sign-in page (no access with this account)'; Headers = @{} }
            }
            return @{ StatusCode = 200; Content = ''; Headers = @{} }
        }
        $h = @{}
        foreach ($k in $r.Headers.Keys) { $h[[string]$k] = [string]::Join(',', @($r.Headers[$k])) }
        $code = [int]$r.StatusCode
        # without access Azure DevOps can answer with its sign-in page instead of JSON
        if ($code -eq 203 -or [string]$h['Content-Type'] -like 'text/html*') {
            return @{ StatusCode = 401; Content = 'Azure DevOps returned its sign-in page (no access with this account)'; Headers = $h }
        }
        $content = $r.Content
        if ($content -is [byte[]]) { $content = [Text.Encoding]::UTF8.GetString($content) }   # e.g. CSV on Windows PowerShell
        return @{ StatusCode = $code; Content = [string]$content; Headers = $h }
    }
    catch {
        $code = 0
        $content = [string]$_.Exception.Message
        $h = @{}
        $resp = $_.Exception.Response
        if ($null -ne $resp) {
            try { $code = [int]$resp.StatusCode } catch { $code = 0 }
            # Retry-After for throttling; X-VSS-ResourceTenant names the tenant of an organization that refused the token
            foreach ($hn in @('Retry-After', 'X-VSS-ResourceTenant')) {
                try {
                    $hv = $null
                    if ($resp -is [System.Net.HttpWebResponse]) { $hv = $resp.Headers[$hn] }
                    else { $vals = $null; if ($resp.Headers.TryGetValues($hn, [ref]$vals)) { $hv = @($vals)[0] } }
                    if ($hv) { $h[$hn] = [string]$hv }
                }
                catch { Write-Verbose ('Header not readable: ' + $hn) }
            }
        }
        if ($null -ne $_.ErrorDetails -and $_.ErrorDetails.Message) { $content = [string]$_.ErrorDetails.Message }
        return @{ StatusCode = $code; Content = $content; Headers = $h }
    }
}

function Get-A2CRetryDelay {
    param($Response, [int] $Attempt)
    $h = $Response.Headers
    $wait = 0.0
    foreach ($n in @('Retry-After', 'x-ms-ratelimit-microsoft.costmanagement-qpu-retry-after',
            'x-ms-ratelimit-microsoft.costmanagement-entity-retry-after', 'x-ms-ratelimit-microsoft.costmanagement-tenant-retry-after',
            'x-ms-ratelimit-microsoft.costmanagement-client-retry-after', 'x-ms-user-quota-resets-after')) {
        if ($null -ne $h -and $h.ContainsKey($n)) { $s = ConvertTo-A2CSeconds $h[$n]; if ($s -gt $wait) { $wait = $s } }
    }
    if ($wait -le 0) { $wait = [math]::Min(60, [math]::Pow(2, $Attempt)) }
    $jitter = (Get-Random -Minimum 0 -Maximum 1000) / 1000.0
    return [math]::Min(120, $wait + $jitter)
}

function Get-A2CShortText {
    param([string] $Text, [int] $Max = 300)
    $t = ($Text -replace '\s+', ' ').Trim()
    if ($t.Length -gt $Max) { $t = $t.Substring(0, $Max) + '...' }
    return $t
}

function Test-A2CSignInError {
    # A transport error (no HTTP status) caused by an expired or blocked Azure sign-in: retrying cannot fix it.
    param([string] $Message)
    return ($Message -match 'Connect-AzAccount|AADSTS\d+|Authentication failed|interaction is required|interaction_required|credentials have not been set up|refresh token has expired')
}

function Invoke-A2CRequest {
    # One request with retries on 429/5xx/transport errors; retries are shown so a slow call never looks like a hang.
    # A sign-in error opens the sign-in once per API kind; if that does not help, it is reported and later calls of
    # that kind fail at once.
    param([string] $Method = 'GET', [string] $Uri, [string] $Body, [string] $Kind = 'arm', [int] $MaxRetries = 6, [string] $OutFile)
    $signInError = $script:A2C.SignInFailed[$Kind]
    if ($signInError) { return @{ StatusCode = 0; Content = $signInError; Headers = @{} } }
    $attempt = 0
    while ($true) {
        $script:A2C.Stats[$Kind] = [int]$script:A2C.Stats[$Kind] + 1
        $r = Invoke-A2CHttp -Method $Method -Uri $Uri -Body $Body -Kind $Kind -OutFile $OutFile
        $code = [int]$r.StatusCode
        if ($code -eq 0 -and (Test-A2CSignInError ([string]$r.Content))) {
            # once per API kind: let the user sign in again (expired session, MFA) and repeat the call
            if (-not $script:A2C.SignInRetried[$Kind] -and (Test-A2CCanPrompt)) {
                $script:A2C.SignInRetried[$Kind] = $true
                if (Request-A2CSignIn -ErrorText ([string]$r.Content)) { continue }
            }
            $script:A2C.SignInFailed[$Kind] = [string]$r.Content
            $fix = 'Connect-AzAccount'
            if ($script:A2C.TenantId) { $fix += ' -TenantId ' + $script:A2C.TenantId }
            Write-A2CLog -Level Error ('Azure sign-in failed ({0} calls): {1}  ->  sign in again with  {2}  and start the script again.' -f $Kind, (Get-A2CShortText $r.Content), $fix)
            return $r
        }
        $retryable = ($code -eq 0 -or $code -eq 429 -or $code -eq 500 -or $code -eq 502 -or $code -eq 503 -or $code -eq 504)
        if (-not $retryable -or $attempt -ge $MaxRetries) { return $r }
        $attempt++
        $delay = Get-A2CRetryDelay -Response $r -Attempt $attempt
        $script:A2C.Stats.Retries++
        if ($code -eq 429) { $script:A2C.Stats.Throttled++ }
        $why = 'HTTP ' + $code
        if ($code -eq 0) { $why = 'No response (' + (Get-A2CShortText $r.Content) + ')' }
        elseif ($code -eq 429) { $why = 'Throttled (HTTP 429)' }
        $msg = '{0}: {1} {2} - retry {3} of {4} in {5:N0}s' -f $why, $Kind, (Get-A2CShortUri $Uri), $attempt, $MaxRetries, $delay
        # short throttling pauses are routine on big tenants and stay in the log only
        if ($code -eq 429 -and $delay -lt 10) { Write-A2CLog -Level Verbose $msg } else { Write-A2CLog -Level Notice $msg }
        Start-A2CSleep -Seconds $delay
    }
}

function Get-A2CShortUri {
    param([string] $Uri)
    if (-not $Uri) { return '' }
    $u = ($Uri -replace '^https?://[^/]+', '') -replace '([?&]sig=)[^&]*', '$1***'   # storage links carry a signature
    if ($u.Length -gt 140) { $u = $u.Substring(0, 140) + '...' }
    return $u
}

function Get-A2CError {
    param($Response)
    $code = ''; $msg = ''
    try {
        $j = ConvertFrom-A2CJson $Response.Content
        if ($null -ne $j) {
            if ($null -ne $j.error) { $code = [string]$j.error.code; $msg = [string]$j.error.message }
            elseif ($null -ne $j.errorCode) { $code = [string]$j.errorCode; $msg = [string]$j.message }   # Fabric
            elseif ($null -ne $j.code) { $code = [string]$j.code; $msg = [string]$j.message }
            elseif ($null -ne $j.message) { $code = [string]$j.typeKey; $msg = [string]$j.message }   # Azure DevOps
        }
    }
    catch { Write-Verbose 'Error response body is not JSON' }
    if (-not $msg) { $msg = [string]$Response.Content }
    if ($msg.Length -gt 400) { $msg = $msg.Substring(0, 400) + '...' }
    return [pscustomobject]@{ Status = [int]$Response.StatusCode; Code = $code; Message = $msg }
}

function Get-A2CErrorText {
    param($ErrorInfo)
    if ($null -eq $ErrorInfo) { return '' }
    $t = 'HTTP ' + $ErrorInfo.Status
    if ($ErrorInfo.Code) { $t += ' ' + $ErrorInfo.Code }
    if ($ErrorInfo.Message) { $t += ': ' + $ErrorInfo.Message }
    return $t
}

function Set-A2CQueryParam {
    param([string] $Uri, [string] $Name, [string] $Value)
    $enc = [uri]::EscapeDataString($Value)
    $rx = '([?&])' + [regex]::Escape($Name) + '=[^&]*'
    if ($Uri -match $rx) { return [regex]::Replace($Uri, $rx, ('$1' + $Name + '=' + $enc.Replace('$', '$$'))) }
    if ($Uri.Contains('?')) { return $Uri + '&' + $Name + '=' + $enc }
    return $Uri + '?' + $Name + '=' + $enc
}

# ---- Power BI ----
# Power BI REST and Fabric REST with Entra tokens from the Az sign-in; read-only.
$script:A2CPbiResource = 'https://analysis.windows.net/powerbi/api'
$script:A2CPbiBase = 'https://api.powerbi.com/v1.0/myorg'
$script:A2CFabricResource = 'https://api.fabric.microsoft.com'
$script:A2CFabricBase = 'https://api.fabric.microsoft.com/v1'

function Invoke-A2CPbi {
    # Single Power BI (-Kind pbi) or Fabric (-Kind fabric) call. $Path is relative to the API base or a full URL.
    # Returns @{ Ok; Status; Json; Content; Error; Headers }
    param([string] $Path, [string] $Method = 'GET', [string] $Body, [string] $Kind = 'pbi', [int] $MaxRetries = 6)
    $uri = $Path
    if ($uri -notmatch '^https?://') { $uri = $(if ($Kind -eq 'fabric') { $script:A2CFabricBase } else { $script:A2CPbiBase }) + $Path }
    $r = Invoke-A2CRequest -Method $Method -Uri $uri -Body $Body -Kind $Kind -MaxRetries $MaxRetries
    $ok = ($r.StatusCode -ge 200 -and $r.StatusCode -lt 300)
    $content = ([string]$r.Content).TrimStart([char]0xFEFF)   # executeQueries answers with a byte order mark
    $json = $null
    if ($ok -and $content) { try { $json = ConvertFrom-A2CJson $content } catch { $json = $null } }
    $err = $null
    if (-not $ok) {
        $err = Get-A2CError $r
        # Power BI errors often have a code and no message: the code says it all, the raw JSON body adds nothing
        if ($err.Code -and $err.Message.TrimStart().StartsWith('{')) { $err.Message = '' }
    }
    return [pscustomobject]@{ Ok = $ok; Status = [int]$r.StatusCode; Json = $json; Content = $content; Error = $err; Headers = $r.Headers }
}

$script:A2CPbiNotApplicable = @('UnsupportedCapacitySKU', 'UnsupportedSku', 'OperationNotSupportedForItem', 'TenantSwitchDisabled', 'FeatureNotAvailable')
$script:A2CPbiListProperties =@('value', 'data', 'tenantSettings', 'domains', 'activityEventEntities', 'ArtifactAccessEntities', 'features')

function Invoke-A2CPbiList {
    # Power BI / Fabric GET with paging. Items are in "value" (or the property named by -ItemsProperty, or the first
    # known list property); the next page comes from continuationUri or @odata.nextLink. Returns @{ Ok; Items; Error }
    param([string] $Path, [string] $Kind = 'pbi', [string] $ItemsProperty = '', [int] $MaxRetries = 6, [int] $MaxItems = 0)
    $items = [System.Collections.Generic.List[object]]::new()
    $seen = [System.Collections.Generic.HashSet[string]]::new()
    $next = $Path
    while ($next -and $seen.Count -lt 10000 -and $seen.Add($next)) {
        $r = Invoke-A2CPbi -Path $next -Kind $Kind -MaxRetries $MaxRetries
        if (-not $r.Ok) { return [pscustomobject]@{ Ok = $false; Items = $items; Error = $r.Error } }
        $j = $r.Json
        $next = ''
        if ($null -eq $j) { break }
        if ($j -is [array]) { foreach ($v in $j) { if ($null -ne $v) { $items.Add($v) } }; break }   # a bare JSON array
        $names = $script:A2CPbiListProperties
        if ($ItemsProperty) { $names = @($ItemsProperty) }
        foreach ($n in $names) {
            if ($null -ne $j.PSObject.Properties[$n]) { foreach ($v in @($j.$n)) { if ($null -ne $v) { $items.Add($v) } }; break }
        }
        if ($MaxItems -gt 0 -and $items.Count -ge $MaxItems) { break }
        if ($null -ne $j.PSObject.Properties['continuationUri'] -and $j.continuationUri) { $next = [string]$j.continuationUri }
        elseif ($null -ne $j.PSObject.Properties['@odata.nextLink'] -and $j.'@odata.nextLink') { $next = [string]$j.'@odata.nextLink' }
        elseif ($null -ne $j.PSObject.Properties['nextLink'] -and $j.nextLink) { $next = [string]$j.nextLink }   # Power Automate
    }
    return [pscustomobject]@{ Ok = $true; Items = $items; Error = $null }
}

function New-A2CLroError {
    # Result in the shape of Invoke-A2CPbi for a long running operation that did not deliver.
    param($Response, [string] $Code, [string] $Message)
    return [pscustomobject]@{ Ok = $false; Status = [int]$Response.Status; Json = $Response.Json; Content = [string]$Response.Content; Headers = $Response.Headers
        Error = [pscustomobject]@{ Status = [int]$Response.Status; Code = $Code; Message = $Message } }
}

function Invoke-A2CFabricLroBatch {
    # Fabric POSTs that can answer 202 Accepted (long running operations), started together and then polled together:
    # a notebook or model definition takes about 20 seconds on its own, a batch takes about as long as one. Used for
    # getDefinition, which reads an item and changes nothing. Returns one result per path, each like Invoke-A2CPbi.
    param([string[]] $Paths, [int] $TimeoutSec = 600, [string] $Method = 'POST')
    $res = New-Object object[] $Paths.Count
    $pending = [System.Collections.Generic.List[object]]::new()
    for ($i = 0; $i -lt $Paths.Count; $i++) {
        $r = Invoke-A2CPbi -Path $Paths[$i] -Method $Method -Kind 'fabric' -MaxRetries 3
        if ($r.Status -ne 202) { $res[$i] = $r; continue }
        $loc = [string]$r.Headers['Location']
        $opId = [string]$r.Headers['x-ms-operation-id']
        if ($opId) { $loc = $script:A2CFabricBase + '/operations/' + $opId }
        if (-not $loc) { $res[$i] = New-A2CLroError $r 'NoOperation' 'Accepted without an operation to wait for'; continue }
        $pending.Add([pscustomobject]@{ Index = $i; Loc = $loc })
    }
    $started = [DateTime]::UtcNow
    $wait = 2.0
    while ($pending.Count -gt 0) {
        Start-A2CSleep -Seconds $wait
        $wait = [math]::Min(10, $wait * 1.5)
        $late = (([DateTime]::UtcNow - $started).TotalSeconds -gt $TimeoutSec)
        foreach ($p in @($pending)) {
            $s = Invoke-A2CPbi -Path $p.Loc -Kind 'fabric' -MaxRetries 3
            $state = [string]$s.Json.status
            if (-not $s.Ok) { $res[$p.Index] = $s }
            elseif ($state -eq 'Succeeded') { $res[$p.Index] = Invoke-A2CPbi -Path ($p.Loc.TrimEnd('/') + '/result') -Kind 'fabric' -MaxRetries 3 }
            elseif ($state -eq 'Failed' -or $state -eq 'Undefined') {
                $msg = [string](Get-A2CFirst $s.Json @('error.message', 'error.errorCode'))
                if (-not $msg) { $msg = 'operation ' + $state }
                $res[$p.Index] = New-A2CLroError $s ([string](Get-A2CFirst $s.Json @('error.errorCode'))) $msg
            }
            elseif ($late) { $res[$p.Index] = New-A2CLroError $s 'Timeout' ('operation still ' + $state + ' after ' + $TimeoutSec + ' s') }
            else { continue }
            [void]$pending.Remove($p)
        }
    }
    return , $res
}

function Invoke-A2CFabricLro {
    # One long running Fabric POST (see Invoke-A2CFabricLroBatch).
    param([string] $Path)
    return (Invoke-A2CFabricLroBatch -Paths @($Path))[0]
}

$script:A2COneLakeBase = 'https://onelake.dfs.fabric.microsoft.com'
$script:A2COneLakeResource = 'https://storage.azure.com/'
$script:A2CFlowBase = 'https://api.flow.microsoft.com'
$script:A2CFlowResource = 'https://service.flow.microsoft.com/'
$script:A2CGraphBase = 'https://graph.microsoft.com'
$script:A2CGraphResource = 'https://graph.microsoft.com/'

function Get-A2COneLakeUsage {
    # Files and bytes below a OneLake folder, summed per subfolder of the given depth (ADLS Gen2 list API, recursive;
    # only names and sizes are read, no file content). $Directory is "<item id>/Tables"; -Depth 2 groups by
    # Tables/<table>, 3 by Tables/<schema>/<table>. Returns @{ Ok; Groups = key -> @{ Files; Bytes; LastModified }; Truncated; Error }
    param([string] $WorkspaceId, [string] $Directory, [int] $Depth, [int] $MaxPages = 200)
    $groups = @{}
    $cont = ''
    for ($page = 0; $page -lt $MaxPages; $page++) {
        $uri = '{0}/{1}?resource=filesystem&recursive=true&maxResults=5000&directory={2}' -f $script:A2COneLakeBase, $WorkspaceId, [uri]::EscapeDataString($Directory)
        if ($cont) { $uri += '&continuation=' + [uri]::EscapeDataString($cont) }
        $r = Invoke-A2CPbi -Path $uri -Kind 'onelake' -MaxRetries 3
        if (-not $r.Ok) { return [pscustomobject]@{ Ok = $false; Groups = $groups; Truncated = $false; Error = $r.Error } }
        foreach ($p in @($r.Json.paths)) {
            if ($null -eq $p) { continue }
            $seg = ([string]$p.name).Split('/')
            if ($seg.Count -le $Depth) { continue }                       # above the grouping level
            $isDir = ([string]$p.isDirectory -eq 'true')
            if ($isDir -and $seg.Count -ne ($Depth + 1)) { continue }       # deeper folders add nothing
            $key = [string]::Join('/', $seg[1..$Depth])
            $g = $groups[$key]
            if ($null -eq $g) { $g = [pscustomobject]@{ Files = 0; Bytes = [long]0; LastModified = $null }; $groups[$key] = $g }
            if ($isDir) { continue }
            $g.Files++
            $len = [long]0
            if ([long]::TryParse([string]$p.contentLength, [ref]$len)) { $g.Bytes += $len }
            $lm = ConvertTo-A2CUtc $p.lastModified
            if ($null -ne $lm -and ($null -eq $g.LastModified -or $lm -gt $g.LastModified)) { $g.LastModified = $lm }
        }
        $cont = [string]$r.Headers['x-ms-continuation']
        if (-not $cont) { return [pscustomobject]@{ Ok = $true; Groups = $groups; Truncated = $false; Error = $null } }
    }
    return [pscustomobject]@{ Ok = $true; Groups = $groups; Truncated = $true; Error = $null }
}

function Add-A2CPbiProblem {
    # Counts a refused or failed call per area (e.g. "semantic model datasources"); the first error text is kept.
    # A workspace with problems gets the coverage status Partial. An answer that says the item or the capacity SKU
    # does not have the feature is not a problem: nothing is missing.
    param([System.Collections.IDictionary] $Problems, [string] $Area, $ErrorInfo)
    if ($null -ne $ErrorInfo -and $script:A2CPbiNotApplicable -contains [string]$ErrorInfo.Code) { return }
    $p = $Problems[$Area]
    if ($null -eq $p) {
        $p = [pscustomobject]@{ Count = 0; First = (Get-A2CShortText (Get-A2CErrorText $ErrorInfo) 160) }
        $Problems[$Area] = $p
    }
    $p.Count++
}

function Get-A2CPbiProblemText {
    # "workspace users: 1 call failed (HTTP 401 ...) | ..."
    param([System.Collections.IDictionary] $Problems)
    $l = [System.Collections.Generic.List[string]]::new()
    foreach ($k in $Problems.Keys) { $l.Add(('{0}: {1} call(s) failed ({2})' -f $k, $Problems[$k].Count, $Problems[$k].First)) }
    return [string]::Join(' | ', $l)
}

function Get-A2CPbiItems {
    # Items of a list call. A refused or failed call is counted in $Problems and gives no items, so one item that
    # the account may not read never stops the workspace.
    param([System.Collections.IDictionary] $Problems, [string] $Area, [string] $Path, [string] $Kind = 'pbi', [string] $ItemsProperty = '')
    $r = Invoke-A2CPbiList -Path $Path -Kind $Kind -ItemsProperty $ItemsProperty
    if (-not $r.Ok) { Add-A2CPbiProblem $Problems $Area $r.Error }
    return , $r.Items
}

$script:A2CSqlResource = 'https://database.windows.net/'

function Invoke-A2CSqlQuery {
    # Read-only queries against a Fabric SQL endpoint (lakehouse or warehouse) over one connection, with an Entra
    # token from the Az sign-in; System.Data.SqlClient is part of Windows PowerShell and PowerShell 7. Like
    # Invoke-A2CHttp this is a network seam (tests replace it) and never throws. $Queries: name -> SELECT.
    # Returns @{ Ok (connected); Error; Sets = name -> rows (ordered dictionaries); SetErrors = name -> text }
    param([string] $Server, [string] $Database, [System.Collections.IDictionary] $Queries, [int] $TimeoutSec = 300)
    $script:A2C.Stats['sql'] = [int]$script:A2C.Stats['sql'] + 1
    $sets = @{}
    $setErrors = @{}
    $conn = $null
    try {
        $b = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
        $b['Data Source'] = $Server; $b['Initial Catalog'] = $Database; $b['Encrypt'] = $true; $b['Connect Timeout'] = 30
        $b['Application Name'] = [string]$script:A2CBrand.Tool
        $conn = New-Object System.Data.SqlClient.SqlConnection $b.ConnectionString
        $conn.AccessToken = Get-A2CBearerToken $script:A2CSqlResource
        $conn.Open()
    }
    catch {
        if ($null -ne $conn) { $conn.Dispose() }
        return [pscustomobject]@{ Ok = $false; Error = (Get-A2CShortText $_.Exception.GetBaseException().Message 200); Sets = $sets; SetErrors = $setErrors }
    }
    try {
        foreach ($name in $Queries.Keys) {
            $rows = [System.Collections.Generic.List[object]]::new()
            try {
                $cmd = $conn.CreateCommand()
                $cmd.CommandText = [string]$Queries[$name]
                $cmd.CommandTimeout = $TimeoutSec
                $rd = $cmd.ExecuteReader()
                try {
                    while ($rd.Read()) {
                        $row = [ordered]@{}
                        for ($i = 0; $i -lt $rd.FieldCount; $i++) { $row[$rd.GetName($i)] = $(if ($rd.IsDBNull($i)) { $null } else { $rd.GetValue($i) }) }
                        $rows.Add($row)
                    }
                }
                finally { $rd.Dispose() }
            }
            catch { $setErrors[$name] = (Get-A2CShortText $_.Exception.GetBaseException().Message 200) }
            $sets[$name] = $rows
        }
    }
    finally { $conn.Dispose() }
    return [pscustomobject]@{ Ok = $true; Error = ''; Sets = $sets; SetErrors = $setErrors }
}

function Get-A2CPbiObject {
    # The JSON object of a single GET; $null (and a counted problem, unless -Quiet) when the call fails.
    param([System.Collections.IDictionary] $Problems, [string] $Area, [string] $Path, [string] $Kind = 'pbi', [switch] $Quiet)
    $r = Invoke-A2CPbi -Path $Path -Kind $Kind
    if ($r.Ok) { return $r.Json }
    if (-not $Quiet) { Add-A2CPbiProblem $Problems $Area $r.Error }
    return $null
}
# ---- end Power BI ----

#endregion

#region ============================== 4. CONTEXT AND SCOPE =============================================

function Get-A2CAzContext {
    try { return (Get-AzContext -ErrorAction Stop) } catch { return $null }
}

function Assert-A2CPrerequisites {
    $m = Get-Module -ListAvailable -Name Az.Accounts | Sort-Object Version -Descending | Select-Object -First 1
    if ($null -eq $m) {
        throw 'The Az.Accounts module is required. Install it with:  Install-Module Az.Accounts -Scope CurrentUser'
    }
    if ($m.Version -lt [version]'2.7.0') {
        throw ('Az.Accounts {0} is too old; 2.7.0 or newer is required:  Update-Module Az.Accounts' -f $m.Version)
    }
    Import-Module Az.Accounts -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
    $loaded = Get-Module Az.Accounts | Sort-Object Version -Descending | Select-Object -First 1
    return [string]$loaded.Version
}

function Connect-A2CAzure {
    param($Options)
    $ctx = Get-A2CAzContext
    $need = ($null -eq $ctx -or $null -eq $ctx.Account)
    if (-not $need -and $Options.TenantId -and $ctx.Tenant -and ($ctx.Tenant.Id -ne $Options.TenantId)) { $need = $true }
    if ($need) {
        Write-A2CLog 'Signing in to Azure (Connect-AzAccount)...'
        $p = @{ ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' }
        if ($Options.TenantId) { $p['TenantId'] = $Options.TenantId }
        if ($Options.UseDeviceAuthentication) { $p['UseDeviceAuthentication'] = $true }
        Invoke-A2CConnectAzAccount -Params $p
        $ctx = Get-A2CAzContext
    }
    else {
        Write-A2CLog ('Reusing the current Az sign-in ({0})' -f $ctx.Account.Id)
    }
    if ($null -eq $ctx) { throw 'No Azure context after sign-in.' }
    return $ctx
}

function Invoke-A2CConnectAzAccount {
    # The only place that calls Connect-AzAccount (tests replace it).
    param([hashtable] $Params)
    Connect-AzAccount @Params | Out-Null
}

function Test-A2CCanPrompt {
    # An interactive sign-in is possible: a user session not started with -NonInteractive (scheduled task, CI).
    return ([Environment]::UserInteractive -and [Environment]::CommandLine -notmatch '\s-noni')
}

function Request-A2CSignIn {
    # The saved sign-in expired or needs MFA again: open the Microsoft sign-in (a device code in Cloud Shell or with
    # -UseDeviceAuthentication) so the run can continue, with the -AuthScope that Az recommends in its error.
    # Returns $true when the sign-in succeeded.
    param([string] $ErrorText)
    $p = @{ ErrorAction = 'Stop'; WarningAction = 'SilentlyContinue' }
    if ($script:A2C.TenantId) { $p['TenantId'] = $script:A2C.TenantId }
    if ($script:A2C.Options.UseDeviceAuthentication -or $script:A2C.Env['CloudShell']) { $p['UseDeviceAuthentication'] = $true }
    if ($ErrorText -match '-AuthScope\s+([^''"\s]+)') { $p['AuthScope'] = $Matches[1].TrimEnd('.') }
    Write-A2CLog -Level Notice 'The Azure sign-in has expired or needs MFA again: opening the Microsoft sign-in (the window may open behind this one).'
    Write-A2CLog -Level Notice 'If Azure asks you to pick a subscription, any choice is fine - the script reads every subscription you can access.'
    try {
        Invoke-A2CConnectAzAccount -Params $p
        $ctx = Get-A2CAzContext
        $who = ''
        if ($null -ne $ctx -and $null -ne $ctx.Account) { $who = ' as ' + $ctx.Account.Id }
        Write-A2CLog -Level Ok ('Signed in again{0} - continuing.' -f $who)
        return $true
    }
    catch {
        Write-A2CLog -Level Error ('The sign-in did not complete: ' + (Get-A2CShortText $_.Exception.Message))
        return $false
    }
}

function Get-A2CHostInfo {
    $cloudShell = [bool]($env:AZUREPS_HOST_ENVIRONMENT -like 'cloud-shell*' -or $env:ACC_CLOUD)
    $os = 'Windows'
    if ($PSVersionTable.PSEdition -eq 'Core') {
        if ($IsLinux) { $os = 'Linux' } elseif ($IsMacOS) { $os = 'macOS' }
    }
    return @{
        PSEdition  = [string]$PSVersionTable.PSEdition
        PSVersion  = [string]$PSVersionTable.PSVersion
        OS         = $os
        CloudShell = $cloudShell
        ClouddriveExists = ($cloudShell -and (Test-Path (Join-Path $HOME 'clouddrive')))
    }
}

#endregion

#region ============================== 5. DATASETS (COLLECTORS) =========================================

# ---- Power BI ----
# Fabric item types whose definition can be exported (getDefinition) and whose connections / schedules are read.
$script:A2CPbiDefinitionTypes = @('SemanticModel', 'Report', 'Notebook', 'DataPipeline', 'SparkJobDefinition', 'Eventstream', 'Eventhouse',
    'KQLDatabase', 'KQLQueryset', 'KQLDashboard', 'Dataflow', 'CopyJob', 'MirroredDatabase', 'Reflex', 'GraphQLApi', 'VariableLibrary',
    'MountedDataFactory', 'DataAgent', 'UserDataFunction', 'ApacheAirflowJob', 'Lakehouse')
$script:A2CPbiConnectionTypes = @('SemanticModel', 'DataPipeline', 'Dataflow', 'Notebook', 'Lakehouse', 'Warehouse', 'Eventstream', 'CopyJob',
    'MirroredDatabase', 'KQLQueryset')
$script:A2CPbiJobTypes = @{ DataPipeline = 'Pipeline'; Notebook = 'RunNotebook'; SparkJobDefinition = 'sparkjob'; Lakehouse = 'TableMaintenance' }
# item types with a list of their own that adds type specific properties to the generic item list
$script:A2CPbiTypedLists = @{ Dataflow = 'dataflows'; Environment = 'environments'; MirroredDatabase = 'mirroredDatabases'; Eventhouse = 'eventhouses'; KQLDatabase = 'kqlDatabases'
    SQLDatabase = 'sqlDatabases'; SparkJobDefinition = 'sparkJobDefinitions' }
$script:A2CPbiRunTypes = @('DataPipeline', 'Notebook', 'SparkJobDefinition', 'CopyJob')   # items with a run history (job instances)

function New-A2CPbiWorkspace {
    # One workspace with everything collected for it. Problems / FabricProblems: area -> failed calls.
    param([string] $Id, [string] $Name)
    return [pscustomobject]@{
        Id = $Id; Name = $Name; Type = ''; Description = ''; CapacityId = ''; DomainId = ''; IsReadOnly = $null; IsOnDedicatedCapacity = $null
        DefaultDatasetStorageFormat = ''; InPbi = $false; InFabric = $false
        PbiPrefix = '/groups/' + $Id                   # '' for "My workspace"
        Status = 'Not collected'; Problems = [ordered]@{}
        FabricStatus = 'Not collected'; FabricProblems = [ordered]@{}
        Users = @(); Reports = @(); Datasets = @(); Dashboards = @(); Dataflows = @(); UpstreamDataflows = @()
        Detail = $null; RoleAssignments = @(); Items = @(); Git = $null; SparkSettings = $null; SparkPools = @()
        ManagedPrivateEndpoints = @(); Lakehouses = @(); Warehouses = @()
        Sub = @{}                                      # kind -> item id -> data read per item (pages, datasources, ...)
    }
}

function Set-A2CPbiSub {
    param($Ws, [string] $Kind, [string] $Id, $Value)
    if (-not $Ws.Sub.ContainsKey($Kind)) { $Ws.Sub[$Kind] = @{} }
    $Ws.Sub[$Kind][$Id] = $Value
}

function Get-A2CPbiSub {
    # Data read for one item: the items of a list (use @(...)), or an object for schedules; nothing when not read.
    param($Ws, [string] $Kind, [string] $Id)
    $k = $Ws.Sub[$Kind]
    if ($null -eq $k -or -not $k.ContainsKey($Id)) { return }
    return $k[$Id]
}

function Test-A2CPbiWorkspaceInScope {
    # -Workspace / -ExcludeWorkspace: names with wildcards, or IDs.
    param($Ws)
    $o = $script:A2C.Options
    foreach ($x in @($o.ExcludeWorkspace)) { if ($x -and ($Ws.Name -like $x -or $Ws.Id -ieq $x)) { return $false } }
    if (@($o.Workspace).Count -eq 0) { return $true }
    foreach ($x in @($o.Workspace)) { if ($x -and ($Ws.Name -like $x -or $Ws.Id -ieq $x)) { return $true } }
    return $false
}

function Get-A2CDatasetPbiWorkspaces {
    $byId = [ordered]@{}
    $errors = [System.Collections.Generic.List[string]]::new()
    $skip = 0
    while ($true) {
        $r = Invoke-A2CPbi ('/groups?$top=5000&$skip=' + $skip)
        if (-not $r.Ok) { $errors.Add('Power BI: ' + (Get-A2CErrorText $r.Error)); break }
        $page = @($r.Json.value)
        foreach ($g in $page) {
            if ($null -eq $g) { continue }
            $w = New-A2CPbiWorkspace ([string]$g.id) ([string]$g.name)
            $w.InPbi = $true; $w.Type = [string]$g.type; $w.IsReadOnly = $g.isReadOnly; $w.IsOnDedicatedCapacity = $g.isOnDedicatedCapacity
            $w.CapacityId = [string]$g.capacityId; $w.DefaultDatasetStorageFormat = [string]$g.defaultDatasetStorageFormat
            $byId[$w.Id.ToLowerInvariant()] = $w
        }
        if ($page.Count -lt 5000) { break }
        $skip += 5000
    }
    $f = Invoke-A2CPbiList -Path '/workspaces' -Kind 'fabric'
    if (-not $f.Ok) { $errors.Add('Fabric: ' + (Get-A2CErrorText $f.Error)) }
    foreach ($g in $f.Items) {
        $key = ([string]$g.id).ToLowerInvariant()
        $w = $byId[$key]
        if ($null -eq $w) {
            $w = New-A2CPbiWorkspace ([string]$g.id) ([string]$g.displayName)
            if ([string]$g.type -eq 'Personal') { $w.PbiPrefix = '' }
            $byId[$key] = $w
        }
        $w.InFabric = $true; $w.Type = [string]$g.type; $w.Description = [string]$g.description; $w.DomainId = [string]$g.domainId
        if ($g.capacityId) { $w.CapacityId = [string]$g.capacityId }
    }
    if ($errors.Count -ge 2) { throw ('The workspace list cannot be read. ' + [string]::Join(' | ', $errors)) }

    $all = @($byId.Values)
    $personal = @($all | Where-Object { $_.Type -eq 'Personal' })
    $scope = [System.Collections.Generic.List[object]]::new()
    foreach ($w in @($all | Sort-Object Name)) {
        if ($w.Type -eq 'Personal' -and -not $script:A2C.Options.PbiIncludePersonalWorkspaces) { continue }
        if (Test-A2CPbiWorkspaceInScope $w) { $scope.Add($w) }
    }
    $script:A2C.Data['PbiWorkspaces'] = $scope
    $script:A2C.Data['PbiWorkspacesVisible'] = $all.Count
    $script:A2C.Data['PbiWorkspacesAll'] = $all
    $note = '{0} of {1} visible workspace(s) in scope' -f $scope.Count, $all.Count
    if ($personal.Count -gt 0 -and -not $script:A2C.Options.PbiIncludePersonalWorkspaces) { $note += '; "My workspace" left out (-PbiIncludePersonalWorkspaces adds it)' }
    if ($errors.Count -gt 0) { $note += '; ' + [string]::Join(' | ', $errors) }
    Write-A2CLog -Level Ok $note
    $status = $(if ($errors.Count -gt 0) { 'Partial' } else { 'Collected' })
    Add-A2CCoverage -Area 'Workspaces' -Status $status -Items $scope.Count -Detail $note
    if ($scope.Count -eq 0) { Write-A2CLog -Level Warn 'No workspace in scope: check -Workspace / -ExcludeWorkspace, or the account is not a member of any workspace.' }
    return New-A2CDatasetResult -Status $status -Note $note
}

function Get-A2CPbiWorkspaceData {
    # Power BI content of one workspace. Per-item calls that the account may not make are counted, not fatal.
    param($Ws)
    $p = $Ws.PbiPrefix
    $pr = $Ws.Problems
    if ($p) { $Ws.Users = Get-A2CPbiItems $pr 'workspace users' ($p + '/users') }

    $Ws.Reports = Get-A2CPbiItems $pr 'reports' ($p + '/reports')
    foreach ($x in $Ws.Reports) {
        $id = [string]$x.id
        if ([string]$x.reportType -eq 'PaginatedReport') {
            Set-A2CPbiSub $Ws 'reportDatasources' $id (Get-A2CPbiItems $pr 'paginated report datasources' ($p + '/reports/' + $id + '/datasources'))
        }
        else { Set-A2CPbiSub $Ws 'reportPages' $id (Get-A2CPbiItems $pr 'report pages' ($p + '/reports/' + $id + '/pages')) }
    }

    $history = [int]$script:A2C.Options.PbiRefreshHistory
    $Ws.Datasets = Get-A2CPbiItems $pr 'semantic models' ($p + '/datasets')
    foreach ($x in $Ws.Datasets) {
        $id = [string]$x.id
        $dp = $p + '/datasets/' + $id
        Set-A2CPbiSub $Ws 'datasetDatasources' $id (Get-A2CPbiItems $pr 'semantic model datasources' ($dp + '/datasources'))
        Set-A2CPbiSub $Ws 'datasetUsers' $id (Get-A2CPbiItems $pr 'semantic model users' ($dp + '/users'))
        Set-A2CPbiSub $Ws 'datasetParameters' $id (Get-A2CPbiItems $pr 'semantic model parameters' ($dp + '/parameters'))
        if ($x.isRefreshable) {
            if ($history -gt 0) { Set-A2CPbiSub $Ws 'datasetRefreshes' $id (Get-A2CPbiItems $pr 'refresh history' ($dp + '/refreshes?$top=' + $history)) }
            Set-A2CPbiSub $Ws 'datasetSchedule' $id (Get-A2CPbiObject $pr 'refresh schedule' ($dp + '/refreshSchedule'))
        }
        else {
            # DirectQuery / live connection: a cache refresh schedule may exist; most models have none (not a problem)
            Set-A2CPbiSub $Ws 'datasetDqSchedule' $id (Get-A2CPbiObject $pr 'refresh schedule' ($dp + '/directQueryRefreshSchedule') -Quiet)
        }
    }
    if ($p -and @($Ws.Datasets).Count -gt 0) { $Ws.UpstreamDataflows = Get-A2CPbiItems $pr 'model to dataflow links' ($p + '/datasets/upstreamDataflows') }

    $Ws.Dashboards = Get-A2CPbiItems $pr 'dashboards' ($p + '/dashboards')
    foreach ($x in $Ws.Dashboards) {
        $id = [string]$x.id
        Set-A2CPbiSub $Ws 'dashboardTiles' $id (Get-A2CPbiItems $pr 'dashboard tiles' ($p + '/dashboards/' + $id + '/tiles'))
    }

    if ($p) {
        $Ws.Dataflows = Get-A2CPbiItems $pr 'dataflows' ($p + '/dataflows')
        foreach ($x in $Ws.Dataflows) {
            $id = [string]$x.objectId
            $fp = $p + '/dataflows/' + $id
            Set-A2CPbiSub $Ws 'dataflowDatasources' $id (Get-A2CPbiItems $pr 'dataflow datasources' ($fp + '/datasources'))
            Set-A2CPbiSub $Ws 'dataflowTransactions' $id (Get-A2CPbiItems $pr 'dataflow transactions' ($fp + '/transactions'))
            Set-A2CPbiSub $Ws 'dataflowUpstream' $id (Get-A2CPbiItems $pr 'dataflow to dataflow links' ($fp + '/upstreamDataflows'))
        }
    }
    $Ws.Status = 'Collected'
}

function Get-A2CPbiTenantData {
    # What is not tied to one workspace: apps, capacities, gateways, deployment pipelines, tenant features.
    $pr = [ordered]@{}
    $t = [pscustomobject]@{ Problems = $pr; Apps = @(); Capacities = @(); Refreshables = @(); Gateways = @(); Pipelines = @(); Features = @(); Sub = @{} }
    $t.Apps = Get-A2CPbiItems $pr 'apps' '/apps'
    foreach ($x in $t.Apps) {
        $id = [string]$x.id
        Set-A2CPbiSub $t 'appReports' $id (Get-A2CPbiItems $pr 'app reports' ('/apps/' + $id + '/reports'))
        Set-A2CPbiSub $t 'appDashboards' $id (Get-A2CPbiItems $pr 'app dashboards' ('/apps/' + $id + '/dashboards'))
    }
    $t.Capacities = Get-A2CPbiItems $pr 'capacities' '/capacities'
    foreach ($x in $t.Capacities) {
        $id = [string]$x.id
        Set-A2CPbiSub $t 'workloads' $id (Get-A2CPbiItems $pr 'capacity workloads (capacity admin)' ('/capacities/' + $id + '/Workloads'))
    }
    if (@($t.Capacities).Count -gt 0) { $t.Refreshables = Get-A2CPbiItems $pr 'refreshables (capacity admin)' '/capacities/refreshables?$expand=capacity,group&$top=1000' }
    $t.Gateways = Get-A2CPbiItems $pr 'gateways' '/gateways'
    foreach ($x in $t.Gateways) {
        $id = [string]$x.id
        $ds = Get-A2CPbiItems $pr 'gateway datasources' ('/gateways/' + $id + '/datasources')
        Set-A2CPbiSub $t 'gatewayDatasources' $id $ds
        foreach ($d in $ds) {
            Set-A2CPbiSub $t 'gatewayDatasourceUsers' ([string]$d.id) (Get-A2CPbiItems $pr 'gateway datasource users' ('/gateways/' + $id + '/datasources/' + [string]$d.id + '/users'))
        }
    }
    $t.Pipelines = Get-A2CPbiItems $pr 'deployment pipelines' '/pipelines'
    foreach ($x in $t.Pipelines) {
        $id = [string]$x.id
        $stages = Get-A2CPbiItems $pr 'pipeline stages' ('/pipelines/' + $id + '/stages')
        Set-A2CPbiSub $t 'pipelineStages' $id $stages
        Set-A2CPbiSub $t 'pipelineUsers' $id (Get-A2CPbiItems $pr 'pipeline users' ('/pipelines/' + $id + '/users'))
        Set-A2CPbiSub $t 'pipelineOperations' $id (Get-A2CPbiItems $pr 'pipeline operations' ('/pipelines/' + $id + '/operations'))
        foreach ($s in $stages) {
            if (-not $s.workspaceId) { continue }
            $a = Get-A2CPbiObject $pr 'pipeline stage items' ('/pipelines/' + $id + '/stages/' + [string]$s.order + '/artifacts')
            Set-A2CPbiSub $t 'pipelineArtifacts' ($id + '|' + [string]$s.order) $a
        }
    }
    $t.Features = Get-A2CPbiItems $pr 'available features' '/availableFeatures' 'pbi' 'features'
    return $t
}

function Get-A2CDatasetPowerBI {
    $wss = @($script:A2C.Data['PbiWorkspaces'])
    $t0 = [DateTime]::Now
    $i = 0
    $partial = 0
    foreach ($w in $wss) {
        $i++
        Write-A2CProgress -Activity 'Power BI workspaces' -Current $i -Total $wss.Count -Started $t0 -Item $w.Name
        Get-A2CPbiWorkspaceData $w
        $n = @($w.Reports).Count + @($w.Datasets).Count + @($w.Dashboards).Count + @($w.Dataflows).Count
        $st = 'Collected'
        if ($w.Problems.Count -gt 0) { $st = 'Partial'; $partial++ }
        Add-A2CCoverage -Area ('Power BI workspace ' + $w.Name) -Status $st -Items $n -Detail (Get-A2CPbiProblemText $w.Problems)
        Write-A2CLog -Level Detail ('{0}: {1} reports, {2} models, {3} dashboards, {4} dataflows{5}' -f $w.Name, @($w.Reports).Count, @($w.Datasets).Count,
            @($w.Dashboards).Count, @($w.Dataflows).Count, $(if ($st -eq 'Partial') { ' (partial - see coverage.csv)' } else { '' }))
    }
    Write-A2CProgress -Activity 'Power BI workspaces' -Completed
    $t = Get-A2CPbiTenantData
    $script:A2C.Data['PbiTenant'] = $t
    $tst = $(if ($t.Problems.Count -gt 0) { 'Partial' } else { 'Collected' })
    Add-A2CCoverage -Area 'Power BI apps, capacities, gateways, pipelines' -Status $tst -Detail (Get-A2CPbiProblemText $t.Problems) `
        -Items (@($t.Apps).Count + @($t.Capacities).Count + @($t.Gateways).Count + @($t.Pipelines).Count)
    Write-A2CLog -Level Ok ('{0} workspace(s), {1} app(s), {2} capacities, {3} gateway(s), {4} deployment pipeline(s)' -f $wss.Count, @($t.Apps).Count,
        @($t.Capacities).Count, @($t.Gateways).Count, @($t.Pipelines).Count)
    if ($partial -gt 0 -or $tst -eq 'Partial') {
        return New-A2CDatasetResult -Status 'Partial' -Note ('{0} workspace(s) partly readable with this account' -f $partial)
    }
    return New-A2CDatasetResult
}

function Get-A2CFabricWorkspaceData {
    # Fabric side of one workspace: items of every type, role assignments, Git, Spark, lakehouse details.
    param($Ws)
    $b = '/workspaces/' + $Ws.Id
    $pr = $Ws.FabricProblems
    $Ws.Detail = Get-A2CPbiObject $pr 'workspace details' $b 'fabric'
    $Ws.RoleAssignments = Get-A2CPbiItems $pr 'role assignments' ($b + '/roleAssignments') 'fabric'
    $Ws.Items = Get-A2CPbiItems $pr 'items' ($b + '/items') 'fabric'
    if ($Ws.Type -ne 'Personal') {
        $Ws.Git = Get-A2CPbiObject $pr 'git connection' ($b + '/git/connection') 'fabric'
        Set-A2CPbiSub $Ws 'folders' 'ws' (Get-A2CPbiItems $pr 'folders' ($b + '/folders') 'fabric')
        Set-A2CPbiSub $Ws 'network' 'ws' (Get-A2CPbiObject $pr 'network policy' ($b + '/networking/communicationPolicy') 'fabric' -Quiet)
        if ([string]$Ws.Git.gitConnectionState -eq 'ConnectedAndInitialized') {
            # which items differ from the branch; can be refused when the account has no Git credentials for the repository
            $gs = (Invoke-A2CFabricLroBatch -Paths @($b + '/git/status') -Method 'GET')[0]
            Set-A2CPbiSub $Ws 'gitStatus' 'ws' ([pscustomobject]@{ Ok = $gs.Ok; Json = $gs.Json; Error = (Get-A2CShortText (Get-A2CErrorText $gs.Error) 200) })
        }
    }
    if ($Ws.CapacityId) {
        # Spark and managed private endpoints exist only on a capacity
        $Ws.SparkSettings = Get-A2CPbiObject $pr 'spark settings' ($b + '/spark/settings') 'fabric'
        $Ws.SparkPools = Get-A2CPbiItems $pr 'spark pools' ($b + '/spark/pools') 'fabric'
        $Ws.ManagedPrivateEndpoints = Get-A2CPbiItems $pr 'managed private endpoints' ($b + '/managedPrivateEndpoints') 'fabric'
        if ([int]$script:A2C.Options.PbiRunHistory -gt 0) {
            # the most recent Spark sessions of the workspace (notebook and job runs, also interactive ones)
            $sr = Invoke-A2CPbiList -Path ($b + '/spark/livySessions') -Kind 'fabric' -MaxItems 200
            if ($sr.Ok) { Set-A2CPbiSub $Ws 'sparkSessions' 'ws' @($sr.Items | Select-Object -First 200) } else { Add-A2CPbiProblem $pr 'spark sessions' $sr.Error }
        }
    }
    $types = @{}
    foreach ($x in $Ws.Items) { $types[[string]$x.type] = $true }
    foreach ($type in @($types.Keys)) {
        $list = $script:A2CPbiTypedLists[$type]
        if (-not $list) { continue }
        $tl = Invoke-A2CPbiList -Path ($b + '/' + $list) -Kind 'fabric'   # extra detail only: no problem when it is refused
        foreach ($x in $tl.Items) {
            $tp = $x.properties
            if ($type -eq 'Environment') {
                $eb = $b + '/environments/' + [string]$x.id
                $tp = [ordered]@{ properties = $x.properties; libraries = (Get-A2CPbiObject $pr 'environment' ($eb + '/libraries') 'fabric' -Quiet)
                    sparkCompute = (Get-A2CPbiObject $pr 'environment' ($eb + '/sparkcompute') 'fabric' -Quiet) }
            }
            if ($null -ne $tp) { Set-A2CPbiSub $Ws 'typedProps' ([string]$x.id) $tp }
        }
    }
    if ($types.ContainsKey('Lakehouse')) {
        $Ws.Lakehouses = Get-A2CPbiItems $pr 'lakehouses' ($b + '/lakehouses') 'fabric'
        foreach ($x in $Ws.Lakehouses) {
            $id = [string]$x.id
            # the table list API does not work for lakehouses with schemas: those keep an empty Tables column
            $tr = Invoke-A2CPbiList -Path ($b + '/lakehouses/' + $id + '/tables') -Kind 'fabric' -ItemsProperty 'data'
            if ($tr.Ok) { Set-A2CPbiSub $Ws 'lakehouseTables' $id $tr.Items }
            elseif ([string]$tr.Error.Code -ne 'UnsupportedOperationForSchemasEnabledLakehouse') { Add-A2CPbiProblem $pr 'lakehouse tables' $tr.Error }
            Set-A2CPbiSub $Ws 'shortcuts' $id (Get-A2CPbiItems $pr 'shortcuts' ($b + '/items/' + $id + '/shortcuts') 'fabric')
            Set-A2CPbiSub $Ws 'dataAccessRoles' $id (Get-A2CPbiItems $pr 'OneLake security roles' ($b + '/items/' + $id + '/dataAccessRoles') 'fabric')
        }
    }
    if ($types.ContainsKey('Warehouse')) { $Ws.Warehouses = Get-A2CPbiItems $pr 'warehouses' ($b + '/warehouses') 'fabric' }
    foreach ($x in $Ws.Items) {
        $id = [string]$x.id
        $type = [string]$x.type
        if ($script:A2CPbiConnectionTypes -contains $type) {
            Set-A2CPbiSub $Ws 'itemConnections' $id (Get-A2CPbiItems $pr 'item connections' ($b + '/items/' + $id + '/connections') 'fabric')
        }
        $runs = [int]$script:A2C.Options.PbiRunHistory
        if ($runs -gt 0 -and $script:A2CPbiRunTypes -contains $type) {
            $rr = Invoke-A2CPbiList -Path ($b + '/items/' + $id + '/jobs/instances') -Kind 'fabric' -MaxItems $runs
            if ($rr.Ok) { Set-A2CPbiSub $Ws 'itemRuns' $id @($rr.Items | Select-Object -First $runs) } else { Add-A2CPbiProblem $pr 'item runs' $rr.Error }
        }
        $job = $script:A2CPbiJobTypes[$type]
        if ($job) {
            Set-A2CPbiSub $Ws 'itemSchedules' $id (Get-A2CPbiItems $pr 'item schedules' ($b + '/items/' + $id + '/jobs/' + $job + '/schedules') 'fabric')
        }
    }
    $Ws.FabricStatus = 'Collected'
}

function Get-A2CDatasetFabric {
    $wss = @($script:A2C.Data['PbiWorkspaces'])
    $t0 = [DateTime]::Now
    $i = 0
    $partial = 0
    $items = 0
    foreach ($w in $wss) {
        $i++
        Write-A2CProgress -Activity 'Fabric workspaces' -Current $i -Total $wss.Count -Started $t0 -Item $w.Name
        Get-A2CFabricWorkspaceData $w
        $items += @($w.Items).Count
        $st = 'Collected'
        if ($w.FabricProblems.Count -gt 0) { $st = 'Partial'; $partial++ }
        Add-A2CCoverage -Area ('Fabric workspace ' + $w.Name) -Status $st -Items @($w.Items).Count -Detail (Get-A2CPbiProblemText $w.FabricProblems)
    }
    Write-A2CProgress -Activity 'Fabric workspaces' -Completed
    $pr = [ordered]@{}
    $t = [pscustomobject]@{ Problems = $pr; Connections = @(); Gateways = @(); Capacities = @(); Domains = @(); DeploymentPipelines = @(); Sub = @{} }
    $t.Connections = Get-A2CPbiItems $pr 'connections' '/connections' 'fabric'
    $t.Gateways = Get-A2CPbiItems $pr 'gateways' '/gateways' 'fabric'
    $t.Capacities = Get-A2CPbiItems $pr 'capacities' '/capacities' 'fabric'
    $t.Domains = Get-A2CPbiItems $pr 'domains' '/domains' 'fabric'
    # deployment pipelines through the Fabric API: stages with every item type (the Power BI API knows only its own)
    $t.DeploymentPipelines = Get-A2CPbiItems $pr 'deployment pipelines' '/deploymentPipelines' 'fabric'
    foreach ($dp in $t.DeploymentPipelines) {
        $dpid = [string]$dp.id
        $stages = Get-A2CPbiItems $pr 'deployment pipeline stages' ('/deploymentPipelines/' + $dpid + '/stages') 'fabric'
        Set-A2CPbiSub $t 'dpStages' $dpid $stages
        Set-A2CPbiSub $t 'dpRoles' $dpid (Get-A2CPbiItems $pr 'deployment pipeline roles' ('/deploymentPipelines/' + $dpid + '/roleAssignments') 'fabric')
        foreach ($st in $stages) {
            if (-not $st.workspaceId) { continue }
            Set-A2CPbiSub $t 'dpStageItems' ([string]$st.id) (Get-A2CPbiItems $pr 'deployment pipeline stage items' ('/deploymentPipelines/' + $dpid + '/stages/' + [string]$st.id + '/items') 'fabric')
        }
    }
    $script:A2C.Data['FabricTenant'] = $t
    $tst = $(if ($pr.Count -gt 0) { 'Partial' } else { 'Collected' })
    Add-A2CCoverage -Area 'Fabric connections, gateways, capacities' -Status $tst -Detail (Get-A2CPbiProblemText $pr) `
        -Items (@($t.Connections).Count + @($t.Gateways).Count + @($t.Capacities).Count)
    Write-A2CLog -Level Ok ('{0} Fabric item(s) in {1} workspace(s), {2} connection(s)' -f $items, $wss.Count, @($t.Connections).Count)
    if ($partial -gt 0 -or $tst -eq 'Partial') {
        return New-A2CDatasetResult -Status 'Partial' -Note ('{0} workspace(s) partly readable with this account' -f $partial)
    }
    return New-A2CDatasetResult
}

$script:A2CSqlObjectTypes = @{ U = 'Table'; V = 'View'; P = 'StoredProcedure'; FN = 'Function'; IF = 'Function'; TF = 'Function' }
# One connection per database, four catalog queries. No user table is read.
$script:A2CSqlQueries = [ordered]@{
    objects    = "SELECT s.name AS [schema], o.name AS [name], RTRIM(o.type) AS [type], o.create_date, o.modify_date " +
        "FROM sys.objects o JOIN sys.schemas s ON s.schema_id = o.schema_id " +
        "WHERE o.type IN ('U', 'V', 'P', 'FN', 'IF', 'TF') AND o.is_ms_shipped = 0 ORDER BY s.name, o.name"
    columns    = "SELECT TABLE_SCHEMA AS [schema], TABLE_NAME AS [table], COLUMN_NAME AS [column], ORDINAL_POSITION AS [ordinal], DATA_TYPE AS [type], " +
        "CHARACTER_MAXIMUM_LENGTH AS [length], NUMERIC_PRECISION AS [precision], NUMERIC_SCALE AS [scale], IS_NULLABLE AS [nullable] " +
        "FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA NOT IN ('sys', 'queryinsights', 'INFORMATION_SCHEMA') ORDER BY 1, 2, 4"
    modules    = "SELECT OBJECT_SCHEMA_NAME(m.object_id) AS [schema], OBJECT_NAME(m.object_id) AS [name], RTRIM(o.type) AS [type], m.definition " +
        "FROM sys.sql_modules m JOIN sys.objects o ON o.object_id = m.object_id WHERE o.is_ms_shipped = 0 ORDER BY 1, 2"
    references = "SELECT DISTINCT OBJECT_SCHEMA_NAME(d.referencing_id) AS [schema], OBJECT_NAME(d.referencing_id) AS [name], d.referenced_database_name AS [ref_database], " +
        "d.referenced_schema_name AS [ref_schema], d.referenced_entity_name AS [ref_entity] FROM sys.sql_expression_dependencies d WHERE d.referenced_database_name IS NOT NULL"
}

function Get-A2CDatasetFabricSql {
    # Catalog of every lakehouse SQL endpoint and warehouse: user tables, views, procedures and functions, their
    # columns, the code of views and procedures, and references to other databases. This is the only way to list
    # the tables of a lakehouse with schemas and the views and procedures of a warehouse.
    $map = @{}   # item id -> @{ Ok; Error; Objects; Columns; Modules; References }
    $script:A2C.Data['FabricSql'] = $map
    $targets = [System.Collections.Generic.List[object]]::new()
    foreach ($w in @($script:A2C.Data['PbiWorkspaces'])) {
        foreach ($x in $w.Lakehouses) {
            $server = [string]$x.properties.sqlEndpointProperties.connectionString
            if ($server) { $targets.Add([pscustomobject]@{ Ws = $w; Type = 'Lakehouse'; Id = [string]$x.id; Name = [string]$x.displayName; Server = $server }) }
        }
        foreach ($x in $w.Warehouses) {
            $server = [string]$x.properties.connectionString
            if ($server) { $targets.Add([pscustomobject]@{ Ws = $w; Type = 'Warehouse'; Id = [string]$x.id; Name = [string]$x.displayName; Server = $server }) }
        }
    }
    if ($targets.Count -eq 0) {
        Add-A2CCoverage -Area 'Fabric SQL endpoints' -Status Skipped -Detail 'No lakehouse or warehouse with a SQL endpoint in scope'
        return New-A2CDatasetResult -Status 'Skipped' -Note 'No lakehouse or warehouse with a SQL endpoint in scope'
    }
    $failed = [System.Collections.Generic.List[string]]::new()
    $partly = [System.Collections.Generic.List[string]]::new()
    $inRow = 0
    $objects = 0
    $t0 = [DateTime]::Now
    $i = 0
    foreach ($t in $targets) {
        $i++
        Write-A2CProgress -Activity 'Fabric SQL endpoints' -Current $i -Total $targets.Count -Started $t0 -Item $t.Name
        if ($inRow -ge 3) {
            # e.g. TCP 1433 is blocked: every further database would wait for the connect timeout
            $map[$t.Id] = [pscustomobject]@{ Ok = $false; Error = 'not tried: the 3 databases before it could not be read'; Objects = @(); Columns = @(); Modules = @(); References = @() }
            continue
        }
        $r = Invoke-A2CSqlQuery -Server $t.Server -Database $t.Name -Queries $script:A2CSqlQueries
        $ok = ($r.Ok -and -not $r.SetErrors.ContainsKey('objects'))
        $err = [string]$r.Error
        if ($r.Ok -and -not $ok) { $err = [string]$r.SetErrors['objects'] }
        $map[$t.Id] = [pscustomobject]@{ Ok = $ok; Error = $err; Objects = @($r.Sets['objects']); Columns = @($r.Sets['columns']); Modules = @($r.Sets['modules']); References = @($r.Sets['references']) }
        if ($ok) {
            $inRow = 0
            $objects += @($r.Sets['objects']).Count
            foreach ($k in @($r.SetErrors.Keys)) { $partly.Add(('{0} / {1} {2}: {3}' -f $t.Ws.Name, $t.Name, $k, $r.SetErrors[$k])) }
        }
        else { $inRow++; $failed.Add(('{0} / {1}: {2}' -f $t.Ws.Name, $t.Name, $err)) }
    }
    Write-A2CProgress -Activity 'Fabric SQL endpoints' -Completed
    $ok = @($map.Values | Where-Object { $_.Ok }).Count
    $note = ''
    if ($ok -lt $targets.Count) { $note = ('{0} of {1} database(s) not read: ' -f ($targets.Count - $ok), $targets.Count) + [string]::Join(' | ', @($failed | Select-Object -First 5)) }
    if ($partly.Count -gt 0) { $note = ($note + ' ' + ('{0} catalog query(ies) refused: ' -f $partly.Count) + [string]::Join(' | ', @($partly | Select-Object -First 5))).Trim() }
    $status = 'Collected'
    if ($ok -lt $targets.Count -or $partly.Count -gt 0) { $status = $(if ($ok -gt 0) { 'Partial' } else { 'Skipped' }) }
    Add-A2CCoverage -Area 'Fabric SQL endpoints' -Status $status -Items $ok -Detail $note
    if ($ok -eq 0) { Write-A2CLog -Level Warn ('No SQL endpoint could be read (outbound TCP 1433 to *.datawarehouse.fabric.microsoft.com is needed): ' + $failed[0]) }
    else { Write-A2CLog -Level Ok ('{0} object(s) in {1} of {2} lakehouse SQL endpoint(s) and warehouse(s)' -f $objects, $ok, $targets.Count) }
    return New-A2CDatasetResult -Status $status -Note $note
}

function Get-A2CSqlSummary {
    # Counts per object type of one lakehouse or warehouse from the FabricSql dataset; Status says why there are none.
    param([string] $ItemId)
    $res = [ordered]@{ Status = ''; Tables = $null; Views = $null; StoredProcedures = $null; Functions = $null; Objects = @() }
    $map = $script:A2C.Data['FabricSql']
    if ($null -eq $map) { $res.Status = 'Not selected (report FabricSql)'; return $res }
    $x = $map[$ItemId]
    if ($null -eq $x) { $res.Status = 'No SQL endpoint'; return $res }
    if (-not $x.Ok) { $res.Status = 'Not read: ' + $x.Error; return $res }
    $res.Status = 'Read'; $res.Objects = @($x.Objects)
    $n = @{ Table = 0; View = 0; StoredProcedure = 0; Function = 0 }
    foreach ($o in $res.Objects) { $k = $script:A2CSqlObjectTypes[[string]$o['type']]; if ($k) { $n[$k]++ } }
    $res.Tables = $n.Table; $res.Views = $n.View; $res.StoredProcedures = $n.StoredProcedure; $res.Functions = $n.Function
    return $res
}

# ---- Item definitions: fetched once, parsed by the Dependencies report, written to disk by PowerBIExport ----
$script:A2CPbiParsedTypes = @('DataPipeline', 'Notebook', 'Dataflow', 'SemanticModel')

function Get-A2CPbiDefinitionText {
    # The parts of an item definition that the dependency parsers read, as text: path -> content. Big parts that
    # nobody parses (translations, diagram layout, report visuals) are left out.
    param([string] $Type, $Definition)
    $parts = @{}
    foreach ($p in @($Definition.parts)) {
        if ($null -eq $p -or -not $p.path) { continue }
        $path = [string]$p.path
        $keep = $false
        switch ($Type) {
            'DataPipeline'  { $keep = ($path -eq 'pipeline-content.json') }
            'Notebook'      { $keep = ($path -like 'notebook-content.*' -or $path -like '*.ipynb') }
            'Dataflow'      { $keep = ($path -eq 'mashup.pq') }
            'SemanticModel' { $keep = ($path -eq 'model.bim' -or $path -like 'definition/tables/*' -or $path -like 'definition/roles/*' -or $path -eq 'definition/expressions.tmdl') }
        }
        if (-not $keep) { continue }
        try { $parts[$path] = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String([string]$p.payload)).TrimStart([char]0xFEFF) }
        catch { Write-Verbose ('Definition part not readable: ' + $path) }
    }
    return $parts
}

function Initialize-A2CPbiExport {
    # exportIndex.csv rows and the item -> exported path map; shared by the ItemDefinitions and PowerBIExport datasets.
    if ($null -ne $script:A2C.Data['PbiExportIndex']) { return }
    $script:A2C.Data['PbiExportIndex'] = [System.Collections.Generic.List[object]]::new()
    $script:A2C.Data['PbiExportMap'] = @{}
    $script:A2C.Data['PbiExportDone'] = @{}
}

function Save-A2CPbiDefinitionParts {
    # Writes every part of a fetched item definition below PowerBI/<workspace>/<type>/<item> and adds the index line.
    param($Ws, $Item, $Result)
    $id = [string]$Item.id
    $type = [string]$Item.type
    $name = [string]$Item.displayName
    $script:A2C.Data['PbiExportDone'][$id] = $true
    if (-not $Result.Ok) {
        Add-A2CPbiExport $Ws $type $name $id 'definition' 'Failed' -Detail (Get-A2CShortText (Get-A2CErrorText $Result.Error) 200)
        return
    }
    $dir = Get-A2CPbiExportDir $Ws $type $name
    $files = 0
    $bytes = [long]0
    $bad = ''
    foreach ($part in @($Result.Json.definition.parts)) {
        if ($null -eq $part -or -not $part.path) { continue }
        try {
            $segs = @(([string]$part.path -split '[\\/]') | Where-Object { $_ } | ForEach-Object { Get-A2CSafePathSegment $_ })
            $path = [IO.Path]::Combine($script:A2C.OutDir, ($dir + '/' + ($segs -join '/')))
            $pd = [IO.Path]::GetDirectoryName($path)
            if (-not [IO.Directory]::Exists($pd)) { [void][IO.Directory]::CreateDirectory($pd) }
            $data = [Convert]::FromBase64String([string]$part.payload)
            [IO.File]::WriteAllBytes($path, $data)
            $files++
            $bytes += $data.Length
        }
        catch { $bad = 'part ' + [string]$part.path + ': ' + $_.Exception.Message }
    }
    if ($files -eq 0) { Add-A2CPbiExport $Ws $type $name $id 'definition' 'Failed' -Detail $(if ($bad) { $bad } else { 'the definition has no parts' }); return }
    Add-A2CPbiExport $Ws $type $name $id 'definition' 'Exported' -Files $files -Bytes $bytes -Path $dir -Detail $bad
}

function Get-A2CPbiDefinitionPath {
    # getDefinition call of an item; notebooks as .ipynb with -PbiNotebookFormat Ipynb (default: Fabric source, .py / .sql)
    param($Ws, $Item)
    $path = '/workspaces/' + $Ws.Id + '/items/' + [string]$Item.id + '/getDefinition'
    if ([string]$Item.type -eq 'Notebook' -and [string]$script:A2C.Options.PbiNotebookFormat -eq 'Ipynb') { $path += '?format=ipynb' }
    return $path
}

function Get-A2CDatasetItemDefinitions {
    $defs = @{}        # item id -> @{ Ws; Item; Parts = path -> text }
    $gen1 = @{}        # dataflow (gen1) id -> model.json text
    $script:A2C.Data['PbiDefs'] = $defs
    $script:A2C.Data['PbiDataflowDefs'] = $gen1
    $export = ((Test-A2CSelected 'PowerBIExport') -or (Test-A2CSelected 'FabricCode'))   # then the definitions are written to disk as they arrive
    if ($export) { Initialize-A2CPbiExport }
    $todo = [System.Collections.Generic.List[object]]::new()
    foreach ($w in @($script:A2C.Data['PbiWorkspaces'])) {
        foreach ($x in $w.Items) { if ($script:A2CPbiParsedTypes -contains [string]$x.type) { $todo.Add([pscustomobject]@{ Ws = $w; Item = $x }) } }
    }
    $stat = @{}
    foreach ($t in $script:A2CPbiParsedTypes) { $stat[$t] = [pscustomobject]@{ Total = 0; Ok = 0; First = '' } }
    $size = 20
    $t0 = [DateTime]::Now
    for ($i = 0; $i -lt $todo.Count; $i += $size) {
        Write-A2CProgress -Activity 'Item definitions' -Current $i -Total $todo.Count -Started $t0 -Item ([string]$todo[$i].Ws.Name)
        $chunk = @($todo[$i..([math]::Min($i + $size - 1, $todo.Count - 1))])
        $res = Invoke-A2CFabricLroBatch -Paths @($chunk | ForEach-Object { Get-A2CPbiDefinitionPath $_.Ws $_.Item })
        for ($k = 0; $k -lt $chunk.Count; $k++) {
            $c = $chunk[$k]
            $r = $res[$k]
            $type = [string]$c.Item.type
            $stat[$type].Total++
            if ($r.Ok) {
                $stat[$type].Ok++
                $defs[[string]$c.Item.id] = [pscustomobject]@{ Ws = $c.Ws; Item = $c.Item; Parts = (Get-A2CPbiDefinitionText $type $r.Json.definition) }
            }
            elseif (-not $stat[$type].First) { $stat[$type].First = ('{0} / {1}: {2}' -f $c.Ws.Name, $c.Item.displayName, (Get-A2CShortText (Get-A2CErrorText $r.Error) 140)) }
            if ($export) { Save-A2CPbiDefinitionParts $c.Ws $c.Item $r }
        }
    }
    Write-A2CProgress -Activity 'Item definitions' -Completed
    # dataflows that are not Fabric items (gen1): their definition is model.json
    $g1 = [pscustomobject]@{ Total = 0; Ok = 0; First = '' }
    foreach ($w in @($script:A2C.Data['PbiWorkspaces'])) {
        foreach ($x in $w.Dataflows) {
            $id = [string]$x.objectId
            if ($defs.ContainsKey($id)) { continue }
            $g1.Total++
            $r = Invoke-A2CPbi ($w.PbiPrefix + '/dataflows/' + $id) -MaxRetries 2
            if ($r.Ok) { $g1.Ok++; $gen1[$id] = [pscustomobject]@{ Ws = $w; Dataflow = $x; Text = $r.Content } }
            elseif (-not $g1.First) { $g1.First = ('{0} / {1}: {2}' -f $w.Name, $x.name, (Get-A2CShortText (Get-A2CErrorText $r.Error) 140)) }
        }
    }
    $stat['Dataflow (gen1)'] = $g1
    $total = 0; $ok = 0
    $notes = [System.Collections.Generic.List[string]]::new()
    foreach ($t in @($stat.Keys | Sort-Object)) {
        $s = $stat[$t]
        $total += $s.Total; $ok += $s.Ok
        if ($s.Total -eq 0) { continue }
        $detail = ''
        if ($s.Ok -lt $s.Total) { $detail = ('{0} of {1} not readable, e.g. {2}' -f ($s.Total - $s.Ok), $s.Total, $s.First); $notes.Add($t + ': ' + $detail) }
        Add-A2CCoverage -Area ('Item definitions: ' + $t) -Status $(if ($s.Ok -eq $s.Total) { 'Collected' } elseif ($s.Ok -gt 0) { 'Partial' } else { 'Skipped' }) -Items $s.Ok -Detail $detail
    }
    Write-A2CLog -Level Ok ('{0} of {1} item definition(s) read (pipelines, notebooks, dataflows, semantic models)' -f $ok, $total)
    $status = 'Collected'
    if ($ok -lt $total) { $status = $(if ($ok -gt 0) { 'Partial' } else { 'Skipped' }) }
    return New-A2CDatasetResult -Status $status -Note ([string]::Join(' | ', $notes))
}

function Get-A2CDatasetFabricStorage {
    # OneLake size of every lakehouse and warehouse table (and of the Files area of lakehouses): file names and
    # sizes only. The bytes include old Delta versions that were not vacuumed yet.
    $list = [System.Collections.Generic.List[object]]::new()
    $script:A2C.Data['FabricStorage'] = $list
    $targets = [System.Collections.Generic.List[object]]::new()
    foreach ($w in @($script:A2C.Data['PbiWorkspaces'])) {
        foreach ($x in $w.Lakehouses) { $targets.Add([pscustomobject]@{ Ws = $w; Type = 'Lakehouse'; Item = $x; Schemas = [bool]$x.properties.defaultSchema }) }
        foreach ($x in $w.Warehouses) { $targets.Add([pscustomobject]@{ Ws = $w; Type = 'Warehouse'; Item = $x; Schemas = $true }) }
    }
    if ($targets.Count -eq 0) {
        Add-A2CCoverage -Area 'OneLake storage' -Status Skipped -Detail 'No lakehouse or warehouse in scope'
        return New-A2CDatasetResult -Status 'Skipped' -Note 'No lakehouse or warehouse in scope'
    }
    $failed = [System.Collections.Generic.List[string]]::new()
    $t0 = [DateTime]::Now
    $i = 0
    foreach ($t in $targets) {
        $i++
        $id = [string]$t.Item.id
        Write-A2CProgress -Activity 'OneLake storage' -Current $i -Total $targets.Count -Started $t0 -Item ([string]$t.Item.displayName)
        $entry = [pscustomobject]@{ Ws = $t.Ws; Type = $t.Type; Item = $t.Item; Schemas = $t.Schemas; Tables = $null; Files = $null; Error = ''; Truncated = $false }
        $r = Get-A2COneLakeUsage -WorkspaceId $t.Ws.Id -Directory ($id + '/Tables') -Depth $(if ($t.Schemas) { 3 } else { 2 })
        if ($r.Ok) { $entry.Tables = $r.Groups; $entry.Truncated = $r.Truncated }
        else { $entry.Error = Get-A2CShortText (Get-A2CErrorText $r.Error) 160; $failed.Add(('{0} / {1}: {2}' -f $t.Ws.Name, $t.Item.displayName, $entry.Error)) }
        if ($r.Ok -and $t.Type -eq 'Lakehouse') {
            $f = Get-A2COneLakeUsage -WorkspaceId $t.Ws.Id -Directory ($id + '/Files') -Depth 1
            if ($f.Ok) { $entry.Files = $f.Groups; if ($f.Truncated) { $entry.Truncated = $true } }
        }
        $list.Add($entry)
    }
    Write-A2CProgress -Activity 'OneLake storage' -Completed
    $ok = @($list | Where-Object { -not $_.Error }).Count
    $note = ''
    if ($failed.Count -gt 0) { $note = ('{0} of {1} not readable: ' -f $failed.Count, $targets.Count) + [string]::Join(' | ', @($failed | Select-Object -First 5)) }
    $status = 'Collected'
    if ($failed.Count -gt 0) { $status = $(if ($ok -gt 0) { 'Partial' } else { 'Skipped' }) }
    Add-A2CCoverage -Area 'OneLake storage' -Status $status -Items $ok -Detail $note
    Write-A2CLog -Level Ok ('OneLake size of {0} of {1} lakehouse(s) and warehouse(s) read' -f $ok, $targets.Count)
    return New-A2CDatasetResult -Status $status -Note $note
}

function Get-A2CDatasetPowerAutomate {
    # Power Automate environments and cloud flows: the flows of the signed-in user and those shared with them, plus
    # all flows of an environment when the account is a Power Platform administrator.
    $pa = [pscustomobject]@{ Environments = @(); Flows = [System.Collections.Generic.List[object]]::new() }
    $script:A2C.Data['PowerAutomate'] = $pa
    $api = '?api-version=2016-11-01'
    $envs = Invoke-A2CPbiList -Path ($script:A2CFlowBase + '/providers/Microsoft.ProcessSimple/environments' + $api) -Kind 'flow' -MaxRetries 2
    if (-not $envs.Ok) {
        $note = 'Power Automate cannot be read with this account: ' + (Get-A2CShortText (Get-A2CErrorText $envs.Error) 160)
        Add-A2CCoverage -Area 'Power Automate' -Status Skipped -Detail $note
        return New-A2CDatasetResult -Status 'Skipped' -Note $note
    }
    $pa.Environments = $envs.Items
    $seen = @{}
    $adminEnvs = 0
    foreach ($e in $envs.Items) {
        $base = $script:A2CFlowBase + '/providers/Microsoft.ProcessSimple'
        $lists = [ordered]@{
            Admin  = $base + '/scopes/admin/environments/' + [string]$e.name + '/v2/flows' + $api
            Owned  = $base + '/environments/' + [string]$e.name + '/flows' + $api + '&$top=50'
            Shared = $base + '/environments/' + [string]$e.name + '/flows' + $api + '&$top=50&$filter=' + [uri]::EscapeDataString("search('team')")
        }
        foreach ($scope in $lists.Keys) {
            $r = Invoke-A2CPbiList -Path $lists[$scope] -Kind 'flow' -MaxRetries 2
            if (-not $r.Ok) { continue }   # Admin: refused without the Power Platform administrator role
            if ($scope -eq 'Admin') { $adminEnvs++ }
            foreach ($f in $r.Items) {
                $key = [string]$e.name + '|' + [string]$f.name
                if ($seen.ContainsKey($key)) { continue }
                $seen[$key] = $true
                $pa.Flows.Add([pscustomobject]@{ Environment = $e; Flow = $f; Scope = $scope })
            }
        }
    }
    $note = '{0} environment(s), {1} flow(s)' -f @($envs.Items).Count, $pa.Flows.Count
    if ($adminEnvs -eq 0) { $note += '; only the flows this account owns or co-owns (all flows need the Power Platform administrator role)' }
    Add-A2CCoverage -Area 'Power Automate' -Status $(if ($adminEnvs -gt 0) { 'Collected' } else { 'Partial' }) -Items $pa.Flows.Count -Detail $note
    Write-A2CLog -Level Ok ('Power Automate: ' + $note)
    return New-A2CDatasetResult -Status $(if ($adminEnvs -gt 0) { 'Collected' } else { 'Partial' }) -Note $note
}

function Invoke-A2CPbiDax {
    # One DAX query against a semantic model (executeQueries; reads, never changes). Returns @{ Ok; Rows; Error }
    param([string] $Prefix, [string] $DatasetId, [string] $Query)
    $body = '{"queries":[{"query":' + (ConvertTo-A2CJsonLiteral $Query) + '}],"serializerSettings":{"includeNulls":true}}'
    $r = Invoke-A2CPbi -Path ($Prefix + '/datasets/' + $DatasetId + '/executeQueries') -Method 'POST' -Body $body -MaxRetries 3
    if (-not $r.Ok) { return [pscustomobject]@{ Ok = $false; Rows = @(); Error = $r.Error } }
    $res = @($r.Json.results)[0]
    if ($null -ne $res -and $null -ne $res.error) {
        return [pscustomobject]@{ Ok = $false; Rows = @(); Error = [pscustomobject]@{ Status = 200; Code = [string]$res.error.code; Message = [string]$res.error.message } }
    }
    $rows = @()
    if ($null -ne $res) { $t = @($res.tables)[0]; if ($null -ne $t) { $rows = @($t.rows | Where-Object { $null -ne $_ }) } }
    return [pscustomobject]@{ Ok = $true; Rows = $rows; Error = $null }
}

function Get-A2CDaxValue {
    # Value of a column of an executeQueries row; the keys are written as "[Name]".
    param($Row, [string[]] $Names)
    foreach ($n in $Names) {
        $p = $Row.PSObject.Properties['[' + $n + ']']
        if ($null -ne $p -and $null -ne $p.Value -and [string]$p.Value -ne '') { return $p.Value }
    }
    return $null
}

function Get-A2CDatasetPowerBIModels {
    $models = [System.Collections.Generic.List[object]]::new()
    $script:A2C.Data['PbiModels'] = $models
    $queries = [ordered]@{ Tables = 'EVALUATE INFO.VIEW.TABLES()'; Columns = 'EVALUATE INFO.VIEW.COLUMNS()'
        Measures = 'EVALUATE INFO.VIEW.MEASURES()'; Relationships = 'EVALUATE INFO.VIEW.RELATIONSHIPS()' }
    $todo = [System.Collections.Generic.List[object]]::new()
    foreach ($w in @($script:A2C.Data['PbiWorkspaces'])) { foreach ($d in $w.Datasets) { $todo.Add(@($w, $d)) } }
    $failed = [System.Collections.Generic.List[string]]::new()
    $t0 = [DateTime]::Now
    $i = 0
    foreach ($pair in $todo) {
        $w = $pair[0]; $d = $pair[1]
        $i++
        Write-A2CProgress -Activity 'Semantic model schema' -Current $i -Total $todo.Count -Started $t0 -Item ([string]$d.name)
        $m = [pscustomobject]@{ Ws = $w; Dataset = $d; Tables = @(); Columns = @(); Measures = @(); Relationships = @() }
        $ok = $true
        foreach ($k in $queries.Keys) {
            $q = Invoke-A2CPbiDax -Prefix $w.PbiPrefix -DatasetId ([string]$d.id) -Query $queries[$k]
            if (-not $q.Ok) { $failed.Add(('{0} / {1}: {2}' -f $w.Name, $d.name, (Get-A2CShortText (Get-A2CErrorText $q.Error) 140))); $ok = $false; break }
            $m.$k = $q.Rows
        }
        if ($ok) { $models.Add($m) }
    }
    Write-A2CProgress -Activity 'Semantic model schema' -Completed
    $note = ''
    if ($failed.Count -gt 0) { $note = ('{0} model(s) not readable: ' -f $failed.Count) + [string]::Join(' | ', @($failed | Select-Object -First 5)) }
    $status = 'Collected'
    if ($failed.Count -gt 0) { $status = $(if ($models.Count -gt 0) { 'Partial' } else { 'Skipped' }) }
    Add-A2CCoverage -Area 'Semantic model schema (DAX INFO.VIEW)' -Status $status -Items $models.Count -Detail $note
    Write-A2CLog -Level Ok ('Schema of {0} of {1} semantic model(s) read' -f $models.Count, $todo.Count)
    return New-A2CDatasetResult -Status $status -Note $note
}

function Get-A2CPbiAdminSkip {
    # Coverage row and dataset result for an admin API that this account may not use.
    param([string] $Area, $ErrorInfo)
    $note = 'Needs the Fabric administrator role (or a service principal allowed to use the read-only admin APIs): ' + (Get-A2CShortText (Get-A2CErrorText $ErrorInfo) 120)
    Add-A2CCoverage -Area $Area -Status Skipped -Detail $note
    Write-A2CLog ($Area + ' skipped: the account is not a Fabric administrator.')
    return New-A2CDatasetResult -Status 'Skipped' -Note $note
}

function Get-A2CDatasetPowerBIAdmin {
    $pr = [ordered]@{}
    $a = [pscustomobject]@{ Problems = $pr; Capacities = @(); Workspaces = [System.Collections.Generic.List[object]]::new()
        Scan = [System.Collections.Generic.List[object]]::new(); DatasourceInstances = [System.Collections.Generic.List[object]]::new()
        TenantSettings = @(); Apps = @(); PublishedToWeb = @(); OrgLinks = @(); Domains = @(); ScanFiles = 0
        Refreshables = [System.Collections.Generic.List[object]]::new() }
    $script:A2C.Data['PbiAdmin'] = $a
    $cs = [bool]$script:A2C.Options.ClientScope
    $probe = Invoke-A2CPbiList -Path '/admin/capacities' -MaxRetries 2
    if (-not $probe.Ok) {
        if ($probe.Error.Status -eq 401 -or $probe.Error.Status -eq 403) { return (Get-A2CPbiAdminSkip 'Power BI admin (tenant-wide)' $probe.Error) }
        throw (Get-A2CErrorText $probe.Error)
    }
    $a.Capacities = $probe.Items

    # all workspaces of the tenant with their users
    $skip = 0
    while ($true) {
        $r = Invoke-A2CPbi ('/admin/groups?$top=5000&$skip=' + $skip + '&$expand=users') -MaxRetries 2
        if (-not $r.Ok) { Add-A2CPbiProblem $pr 'admin workspaces' $r.Error; break }
        $page = @($r.Json.value | Where-Object { $null -ne $_ })
        foreach ($g in $page) { $a.Workspaces.Add($g) }
        if ($page.Count -lt 5000) { break }
        $skip += 5000
    }

    # Scanner API: metadata of every workspace in batches of 100 (a POST that starts a read-only scan)
    $m = Invoke-A2CPbi '/admin/workspaces/modified?excludePersonalWorkspaces=False&excludeInActiveWorkspaces=False' -MaxRetries 2
    if (-not $m.Ok) { Add-A2CPbiProblem $pr 'scanner: workspace list' $m.Error }
    else {
        $ids = @($m.Json | ForEach-Object { [string]$_.id } | Where-Object { $_ })
        # -ClientScope: no model schema, no expressions and no users per item; they never leave the service
        $scanPath = '/admin/workspaces/getInfo?lineage=True&datasourceDetails=True'
        if (-not $cs) { $scanPath += '&datasetSchema=True&datasetExpressions=True&getArtifactUsers=True' }
        $t0 = [DateTime]::Now
        for ($i = 0; $i -lt $ids.Count; $i += 100) {
            Write-A2CProgress -Activity 'Scanner API' -Current $i -Total $ids.Count -Started $t0
            $batch = @($ids[$i..([math]::Min($i + 99, $ids.Count - 1))])
            $body = '{"workspaces":[' + [string]::Join(',', @($batch | ForEach-Object { '"' + $_ + '"' })) + ']}'
            $s = Invoke-A2CPbi $scanPath -Method 'POST' -Body $body -MaxRetries 2
            if (-not $s.Ok) { Add-A2CPbiProblem $pr 'scanner: start scan' $s.Error; continue }
            $scanId = [string]$s.Json.id
            $state = [string]$s.Json.status
            $started = [DateTime]::UtcNow
            while ($state -ne 'Succeeded' -and ([DateTime]::UtcNow - $started).TotalSeconds -lt 900) {
                Start-A2CSleep -Seconds 3
                $q = Invoke-A2CPbi ('/admin/workspaces/scanStatus/' + $scanId) -MaxRetries 2
                if (-not $q.Ok) { Add-A2CPbiProblem $pr 'scanner: scan status' $q.Error; break }
                $state = [string]$q.Json.status
            }
            if ($state -ne 'Succeeded') { continue }
            $res = Invoke-A2CPbi ('/admin/workspaces/scanResult/' + $scanId) -MaxRetries 2
            if (-not $res.Ok) { Add-A2CPbiProblem $pr 'scanner: scan result' $res.Error; continue }
            $a.ScanFiles++
            if (-not $cs) { Save-A2CExportFile ('PowerBI/scan/scanResult_{0:000}.json' -f $a.ScanFiles) $res.Content }   # the complete answer, for what the CSVs leave out
            foreach ($w in @($res.Json.workspaces)) { if ($null -ne $w) { $a.Scan.Add($w) } }
            foreach ($d in @($res.Json.datasourceInstances)) { if ($null -ne $d) { $a.DatasourceInstances.Add($d) } }
        }
        Write-A2CProgress -Activity 'Scanner API' -Completed
    }

    # refresh schedules of every model on a dedicated capacity (there is no tenant-wide schedule API for shared capacity)
    $skip = 0
    while ($true) {
        $r = Invoke-A2CPbi ('/admin/capacities/refreshables?$top=1000&$skip=' + $skip + '&$expand=capacity,group') -MaxRetries 2
        if (-not $r.Ok) { Add-A2CPbiProblem $pr 'refreshables (all capacities)' $r.Error; break }
        $page = @($r.Json.value | Where-Object { $null -ne $_ })
        foreach ($x in $page) { $a.Refreshables.Add($x) }
        if ($page.Count -lt 1000) { break }
        $skip += 1000
    }

    # workspaces changed in the last 30 days (the longest the API looks back)
    $since = $script:A2C.Now.AddDays(-30).AddHours(1).ToString('yyyy-MM-ddTHH:mm:ssZ', [Globalization.CultureInfo]::InvariantCulture)
    $rm = Invoke-A2CPbi ('/admin/workspaces/modified?modifiedSince=' + [uri]::EscapeDataString($since) + '&excludePersonalWorkspaces=False&excludeInActiveWorkspaces=False') -MaxRetries 2
    if ($rm.Ok) {
        $recent = @{}
        foreach ($x in @($rm.Json)) { if ($null -ne $x -and $x.id) { $recent[([string]$x.id).ToLowerInvariant()] = $true } }
        $script:A2C.Data['PbiAdminRecent'] = $recent
    }
    else { Add-A2CPbiProblem $pr 'scanner: workspaces changed in the last 30 days' $rm.Error }

    if (-not $cs) {
        $a.Apps = Get-A2CPbiItems $pr 'admin apps' '/admin/apps?$top=5000'
        $a.PublishedToWeb = Get-A2CPbiItems $pr 'published to web' '/admin/widelySharedArtifacts/publishedToWeb'
        $a.OrgLinks = Get-A2CPbiItems $pr 'links shared with the organization' '/admin/widelySharedArtifacts/linksSharedToWholeOrganization'
        $a.TenantSettings = Get-A2CPbiItems $pr 'tenant settings' '/admin/tenantsettings' 'fabric' 'tenantSettings'
        $a.Domains = Get-A2CPbiItems $pr 'domains' '/admin/domains' 'fabric' 'domains'
    }

    $status = $(if ($pr.Count -gt 0) { 'Partial' } else { 'Collected' })
    $note = Get-A2CPbiProblemText $pr
    Add-A2CCoverage -Area 'Power BI admin (tenant-wide)' -Status $status -Items $a.Workspaces.Count -Detail $note
    Write-A2CLog -Level Ok ('Tenant: {0} workspace(s), {1} scanned, {2} tenant setting(s)' -f $a.Workspaces.Count, $a.Scan.Count, @($a.TenantSettings).Count)
    return New-A2CDatasetResult -Status $status -Note $note
}

# ---- Principals: everyone and everything with access, looked up in Entra ID ----
$script:A2CGuidOnly = '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'

function Add-A2CPbiPrincipal {
    # Adds a role of a principal to the inventory: merged by Entra object ID, otherwise by identifier (UPN / e-mail).
    param($Inv, [string] $ObjectId, [string] $Identifier, [string] $Name, [string] $Type, [string] $UserType, [string] $AppId, [string] $Role, [string] $WorkspaceId, [string] $Source)
    $oid = $(if ($ObjectId -match $script:A2CGuidOnly) { $ObjectId.ToLowerInvariant() } else { '' })
    # Power BI names a group by its Entra object ID and a user by the UPN; an app's identifier is its app ID
    if (-not $oid -and $Type -ne 'App' -and $Identifier -match $script:A2CGuidOnly) { $oid = $Identifier.ToLowerInvariant() }
    $idn = $Identifier.Trim().ToLowerInvariant()
    if (-not $oid -and -not $idn) { return }
    $key = ''
    if ($oid -and $Inv.ByKey.ContainsKey('oid:' + $oid)) { $key = 'oid:' + $oid }
    elseif ($idn -and $Inv.Alias.ContainsKey($idn)) { $key = $Inv.Alias[$idn] }
    elseif ($oid) { $key = 'oid:' + $oid }
    else { $key = 'id:' + $idn }
    $p = $Inv.ByKey[$key]
    if ($null -eq $p) {
        $p = [pscustomobject]@{ ObjectId = $oid; Identifier = $Identifier; Name = $Name; Type = ''; UserType = ''; AppId = $AppId; Roles = [ordered]@{}
            Workspaces = [System.Collections.Generic.HashSet[string]]::new(); Sources = [System.Collections.Generic.HashSet[string]]::new(); EntraStatus = ''; Departed = '' }
        $Inv.ByKey[$key] = $p
        $Inv.List.Add($p)
    }
    if ($oid -and -not $p.ObjectId) { $p.ObjectId = $oid; $Inv.ByKey['oid:' + $oid] = $p }
    if ($idn) { $Inv.Alias[$idn] = $key }
    if (-not $p.Identifier -and $Identifier) { $p.Identifier = $Identifier }
    if (-not $p.Name -and $Name) { $p.Name = $Name }
    if (-not $p.AppId -and $AppId) { $p.AppId = $AppId }
    if ($UserType) { $p.UserType = $UserType }
    $t = $Type
    if ($t -in @('ServicePrincipal', 'ServicePrincipalProfile')) { $t = 'App' }
    if ($t -and -not $p.Type) { $p.Type = $t }
    if ($Role) {
        # counted per distinct workspace: the same assignment seen through two APIs counts once
        if (-not $p.Roles.Contains($Role)) { $p.Roles[$Role] = [System.Collections.Generic.HashSet[string]]::new() }
        [void]$p.Roles[$Role].Add($(if ($WorkspaceId) { $WorkspaceId.ToLowerInvariant() } else { $Source }))
    }
    if ($WorkspaceId) { [void]$p.Workspaces.Add($WorkspaceId.ToLowerInvariant()) }
    if ($Source) { [void]$p.Sources.Add($Source) }
}

function Get-A2CPbiPrincipalStatus {
    # 'Departed' when the Principals report found the account deleted or disabled in Entra ID, otherwise ''.
    param([string] $ObjectId, [string] $Identifier)
    $inv = $script:A2C.Data['PbiPrincipalInventory']
    if ($null -eq $inv) { return '' }
    $p = $null
    if ($ObjectId -match $script:A2CGuidOnly) { $p = $inv.ByKey['oid:' + $ObjectId.ToLowerInvariant()] }
    if ($null -eq $p -and $Identifier) { $k = $inv.Alias[$Identifier.Trim().ToLowerInvariant()]; if ($k) { $p = $inv.ByKey[$k] } }
    if ($null -ne $p -and $p.Departed -eq 'Yes') { return 'Departed' }
    return ''
}

function Get-A2CDatasetPrincipals {
    # Every principal in the data the other datasets collected (it runs after them), de-duplicated, then looked up
    # in Entra ID by object ID: one getByIds call per 1000 IDs, and one list of disabled users.
    $inv = [pscustomobject]@{ ByKey = @{}; Alias = @{}; List = [System.Collections.Generic.List[object]]::new() }
    $script:A2C.Data['PbiPrincipalInventory'] = $inv
    $script:A2C.Data['PbiPrincipals'] = $inv.List
    $cs = [bool]$script:A2C.Options.ClientScope
    $a = $script:A2C.Data['PbiAdmin']
    if ($null -ne $a -and (Test-A2CDataset 'PowerBIAdmin')) {
        foreach ($g in $a.Workspaces) {
            foreach ($u in @($g.users | Where-Object { $null -ne $_ })) {
                Add-A2CPbiPrincipal $inv ([string]$u.graphId) ([string](Get-A2CFirst $u @('identifier', 'emailAddress'))) ([string]$u.displayName) ([string]$u.principalType) ([string]$u.userType) `
                    $(if ([string]$u.principalType -eq 'App') { [string]$u.identifier } else { '' }) ('Workspace ' + [string]$u.groupUserAccessRight) ([string]$g.id) 'admin workspaces'
            }
        }
        if (-not $cs) {
            foreach ($sw in $a.Scan) {
                foreach ($prop in $sw.PSObject.Properties) {
                    if ($prop.Name -eq 'users' -or -not ($prop.Value -is [array])) { continue }
                    foreach ($x in $prop.Value) {
                        if ($null -eq $x) { continue }
                        if ($x.configuredBy -and [string]$x.configuredBy -match '@') { Add-A2CPbiPrincipal $inv '' ([string]$x.configuredBy) '' 'User' '' '' 'Item owner (configuredBy)' ([string]$sw.id) 'admin items' }
                        foreach ($u in @($x.users | Where-Object { $null -ne $_ })) {
                            Add-A2CPbiPrincipal $inv ([string]$u.graphId) ([string](Get-A2CFirst $u @('identifier', 'emailAddress'))) ([string]$u.displayName) ([string]$u.principalType) ([string]$u.userType) '' 'Item access' ([string]$sw.id) 'admin item users'
                        }
                    }
                }
            }
        }
        foreach ($c in $a.Capacities) { foreach ($m in @($c.admins | Where-Object { $_ })) { Add-A2CPbiPrincipal $inv '' ([string]$m) '' 'User' '' '' 'Capacity admin' '' 'admin capacities' } }
    }
    foreach ($w in @($script:A2C.Data['PbiWorkspaces'] | Where-Object { $null -ne $_ })) {
        foreach ($u in $w.Users) {
            Add-A2CPbiPrincipal $inv ([string]$u.graphId) ([string](Get-A2CFirst $u @('identifier', 'emailAddress'))) ([string]$u.displayName) ([string]$u.principalType) ([string]$u.userType) `
                $(if ([string]$u.principalType -eq 'App') { [string]$u.identifier } else { '' }) ('Workspace ' + [string]$u.groupUserAccessRight) $w.Id 'workspace users'
        }
        foreach ($ra in $w.RoleAssignments) {
            $pn = $ra.principal
            Add-A2CPbiPrincipal $inv ([string]$pn.id) ([string](Get-A2CFirst $pn @('userDetails.userPrincipalName', 'servicePrincipalDetails.aadAppId'))) ([string]$pn.displayName) ([string]$pn.type) '' ([string](Get-A2CFirst $pn @('servicePrincipalDetails.aadAppId'))) `
                ('Workspace ' + [string]$ra.role) $w.Id 'Fabric role assignments'
        }
        foreach ($x in (@($w.Datasets) + @($w.Dataflows))) {
            if ($null -ne $x -and [string]$x.configuredBy -match '@') { Add-A2CPbiPrincipal $inv '' ([string]$x.configuredBy) '' 'User' '' '' 'Item owner (configuredBy)' $w.Id 'semantic models and dataflows' }
        }
        foreach ($d in $w.Datasets) {
            foreach ($u in @(Get-A2CPbiSub $w 'datasetUsers' ([string]$d.id))) {
                Add-A2CPbiPrincipal $inv ([string]$u.graphId) ([string]$u.identifier) '' ([string]$u.principalType) '' '' 'Semantic model access' $w.Id 'semantic model users'
            }
        }
    }
    $t = $script:A2C.Data['PbiTenant']
    if ($null -ne $t) {
        foreach ($g in $t.Gateways) {
            foreach ($ds in @(Get-A2CPbiSub $t 'gatewayDatasources' ([string]$g.id))) {
                foreach ($u in @(Get-A2CPbiSub $t 'gatewayDatasourceUsers' ([string]$ds.id))) {
                    Add-A2CPbiPrincipal $inv ([string]$u.graphId) ([string](Get-A2CFirst $u @('identifier', 'emailAddress'))) ([string]$u.displayName) ([string]$u.principalType) '' '' 'Gateway data source' '' 'gateway users'
                }
            }
        }
        foreach ($pl in $t.Pipelines) { foreach ($u in @(Get-A2CPbiSub $t 'pipelineUsers' ([string]$pl.id))) { Add-A2CPbiPrincipal $inv '' ([string]$u.identifier) '' ([string]$u.principalType) '' '' 'Deployment pipeline' '' 'deployment pipelines' } }
    }
    $f = $script:A2C.Data['FabricTenant']
    if ($null -ne $f) {
        foreach ($dp in $f.DeploymentPipelines) {
            foreach ($ra in @(Get-A2CPbiSub $f 'dpRoles' ([string]$dp.id))) { Add-A2CPbiPrincipal $inv ([string]$ra.principal.id) '' '' ([string]$ra.principal.type) '' '' 'Deployment pipeline' '' 'deployment pipelines' }
        }
    }
    if ($inv.List.Count -eq 0) {
        Add-A2CCoverage -Area 'Principals' -Status Skipped -Detail 'No principals in the collected data'
        return New-A2CDatasetResult -Status 'Skipped' -Note 'No principals in the collected data'
    }

    # Entra ID, step 1: users known only by UPN (the Power BI APIs often give no object ID): /users/{upn} through
    # $batch, 20 per request. 404 = no such account any more.
    $lookupError = ''
    $byName = @($inv.List | Where-Object { -not $_.ObjectId -and $_.Identifier -match '@' -and $_.Type -ne 'Group' -and $_.Type -ne 'App' })
    $notFound = @{}
    $nameEnabled = @{}
    for ($i = 0; $i -lt $byName.Count; $i += 20) {
        $chunk = @($byName[$i..([math]::Min($i + 19, $byName.Count - 1))])
        $reqs = for ($k = 0; $k -lt $chunk.Count; $k++) {
            '{"id":"' + $k + '","method":"GET","url":' + (ConvertTo-A2CJsonLiteral ('/users/' + [uri]::EscapeDataString([string]$chunk[$k].Identifier) + '?$select=id,displayName,userType,accountEnabled')) + '}'
        }
        $r = Invoke-A2CPbi -Path ($script:A2CGraphBase + '/v1.0/$batch') -Method 'POST' -Body ('{"requests":[' + [string]::Join(',', @($reqs)) + ']}') -Kind 'graph' -MaxRetries 3
        if (-not $r.Ok) { $lookupError = Get-A2CShortText (Get-A2CErrorText $r.Error) 160; break }
        foreach ($resp in @($r.Json.responses)) {
            if ($null -eq $resp) { continue }
            $p = $chunk[[int]$resp.id]
            if ([int]$resp.status -eq 200 -and $resp.body.id) {
                $oid = ([string]$resp.body.id).ToLowerInvariant()
                $p.ObjectId = $oid
                if (-not $inv.ByKey.ContainsKey('oid:' + $oid)) { $inv.ByKey['oid:' + $oid] = $p }
                if ($null -ne $resp.body.PSObject.Properties['accountEnabled'] -and $null -ne $resp.body.accountEnabled) { $nameEnabled[$oid] = [bool]$resp.body.accountEnabled }
            }
            elseif ([int]$resp.status -eq 404) { $notFound[$p.Identifier.ToLowerInvariant()] = $true }
        }
    }

    # step 2: which object IDs still exist, what they are; step 3: which users are disabled
    $ids = @($inv.List | Where-Object { $_.ObjectId } | ForEach-Object { $_.ObjectId })
    $found = @{}
    for ($i = 0; $i -lt $ids.Count -and -not $lookupError; $i += 1000) {
        $chunk = @($ids[$i..([math]::Min($i + 999, $ids.Count - 1))])
        $body = '{"ids":[' + [string]::Join(',', @($chunk | ForEach-Object { '"' + $_ + '"' })) + '],"types":["user","group","servicePrincipal"]}'
        $r = Invoke-A2CPbi -Path ($script:A2CGraphBase + '/v1.0/directoryObjects/getByIds') -Method 'POST' -Body $body -Kind 'graph' -MaxRetries 3
        if (-not $r.Ok) { $lookupError = Get-A2CShortText (Get-A2CErrorText $r.Error) 160; break }
        foreach ($o in @($r.Json.value)) { if ($null -ne $o -and $o.id) { $found[([string]$o.id).ToLowerInvariant()] = $o } }
    }
    $disabled = $null
    if (-not $lookupError -and $ids.Count -gt 0) {
        $dr = Invoke-A2CPbiList -Path ($script:A2CGraphBase + '/v1.0/users?$filter=accountEnabled%20eq%20false&$select=id&$top=999') -Kind 'graph' -MaxRetries 3
        if ($dr.Ok) { $disabled = @{}; foreach ($o in $dr.Items) { if ($o.id) { $disabled[([string]$o.id).ToLowerInvariant()] = $true } } }
    }
    foreach ($p in $inv.List) {
        if ($p.Identifier -match '#EXT#') { $p.Type = 'Guest' }
        elseif ($p.Type -eq 'User' -and $p.UserType -eq 'Guest') { $p.Type = 'Guest' }
        if ($lookupError) { $p.EntraStatus = 'Not checked: ' + $lookupError; continue }
        if (-not $p.ObjectId) {
            if ($notFound.ContainsKey(([string]$p.Identifier).ToLowerInvariant())) { $p.EntraStatus = 'Not found (deleted)'; $p.Departed = 'Yes' }
            else { $p.EntraStatus = 'Not checked (no object ID)' }
            continue
        }
        $o = $found[$p.ObjectId]
        if ($null -eq $o) { $p.EntraStatus = 'Not found (deleted)'; $p.Departed = 'Yes'; continue }
        $ot = [string]$o.'@odata.type'
        if ($ot -like '*servicePrincipal') {
            $p.Type = $(if ([string]$o.servicePrincipalType -eq 'ManagedIdentity') { 'Managed identity' } else { 'Service principal' })
            if (-not $p.AppId) { $p.AppId = [string]$o.appId }
        }
        elseif ($ot -like '*group') { $p.Type = 'Group' }
        elseif ($ot -like '*user') { $p.Type = $(if ([string]$o.userType -eq 'Guest' -or $p.Type -eq 'Guest') { 'Guest' } else { 'User' }) }
        if (-not $p.Name) { $p.Name = [string]$o.displayName }
        if (($null -ne $disabled -and $disabled.ContainsKey($p.ObjectId)) -or $nameEnabled[$p.ObjectId] -eq $false) { $p.EntraStatus = 'Disabled'; $p.Departed = 'Yes' }
        else { $p.EntraStatus = $(if ($ot -like '*user' -and $null -eq $disabled) { 'Exists (enabled state not readable)' } else { 'Active' }); $p.Departed = 'No' }
    }
    $note = '{0} principal(s), {1} looked up in Entra ID' -f $inv.List.Count, ($ids.Count + $notFound.Count)
    $status = 'Collected'
    if ($lookupError) { $note += '; Entra ID lookup failed: ' + $lookupError; $status = 'Partial' }
    elseif ($null -eq $disabled -and $ids.Count -gt 0) { $note += '; disabled accounts not readable with this account' }
    Add-A2CCoverage -Area 'Principals' -Status $status -Items $inv.List.Count -Detail $note
    Write-A2CLog -Level Ok ('Principals: ' + $note)
    return New-A2CDatasetResult -Status $status -Note $note
}

function New-A2CReportPrincipals {
    $list = $script:A2C.Data['PbiPrincipals']
    if ($null -eq $list) { return }
    $cs = [bool]$script:A2C.Options.ClientScope
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($p in @($list | Sort-Object Type, Name, Identifier)) {
        $person = ($p.Type -eq 'User' -or $p.Type -eq 'Guest' -or -not $p.Type)
        # -ClientScope: people only when they own something (workspace or capacity admin); everyone else is a count
        if ($cs -and $person -and -not ($p.Roles.Contains('Workspace Admin') -or $p.Roles.Contains('Capacity admin'))) { continue }
        $rows.Add([ordered]@{ Principal = $p.Name; Identifier = $p.Identifier; ObjectId = $p.ObjectId; PrincipalType = $p.Type; AppId = $p.AppId; Workspaces = $p.Workspaces.Count
                Roles = (Join-A2C @($p.Roles.Keys | ForEach-Object { '{0} ({1})' -f $_, $p.Roles[$_].Count })); Sources = (Join-A2C @($p.Sources)); EntraStatus = $p.EntraStatus; Departed = $p.Departed })
    }
    Add-A2CRows 'PowerBI/principals.csv' $rows
}

function Get-A2CDatasetPowerBIActivity {
    $events = [System.Collections.Generic.List[object]]::new()
    $script:A2C.Data['PbiActivity'] = $events
    $days = [int]$script:A2C.Options.PbiActivityDays
    $failed = [System.Collections.Generic.List[string]]::new()
    $t0 = [DateTime]::Now
    for ($i = $days - 1; $i -ge 0; $i--) {
        # the API accepts one UTC day per request
        $day = $script:A2C.Now.Date.AddDays(-$i).ToString('yyyy-MM-dd', [Globalization.CultureInfo]::InvariantCulture)
        Write-A2CProgress -Activity 'Activity log' -Current ($days - $i) -Total $days -Started $t0 -Item $day
        $path = "/admin/activityevents?startDateTime='{0}T00:00:00.000Z'&endDateTime='{0}T23:59:59.999Z'" -f $day
        $r = Invoke-A2CPbiList -Path $path -ItemsProperty 'activityEventEntities' -MaxRetries 2
        if (-not $r.Ok) {
            if ($events.Count -eq 0 -and $failed.Count -eq 0 -and ($r.Error.Status -eq 401 -or $r.Error.Status -eq 403)) {
                Write-A2CProgress -Activity 'Activity log' -Completed
                return (Get-A2CPbiAdminSkip 'Power BI activity log' $r.Error)
            }
            $failed.Add($day + ': ' + (Get-A2CShortText (Get-A2CErrorText $r.Error) 120))
        }
        foreach ($e in $r.Items) { $events.Add($e) }
    }
    Write-A2CProgress -Activity 'Activity log' -Completed
    $status = $(if ($failed.Count -gt 0) { 'Partial' } else { 'Collected' })
    $note = [string]::Join(' | ', $failed)
    Add-A2CCoverage -Area 'Power BI activity log' -Status $status -Items $events.Count -Detail ('{0} day(s). {1}' -f $days, $note).Trim()
    Write-A2CLog -Level Ok ('{0} activity event(s) of the last {1} day(s)' -f $events.Count, $days)
    return New-A2CDatasetResult -Status $status -Note $note
}

function Get-A2CPbiExportDir {
    # PowerBI/<workspace>/<folder>/<item>, with names that are safe as folder names
    param($Ws, [string] $Folder, [string] $Item)
    $parts = @('PowerBI', (Get-A2CSafePathSegment $Ws.Name), (Get-A2CSafePathSegment $Folder))
    if ($Item) { $parts += (Get-A2CSafePathSegment $Item) }
    return ($parts -join '/')
}

function Add-A2CPbiExport {
    # One line of exportIndex.csv; a successful export is also remembered per item for the ExportedTo columns.
    param($Ws, [string] $ItemType, [string] $Item, [string] $ItemId, [string] $Kind, [string] $Status, [int] $Files = 0, [long] $Bytes = 0, [string] $Path = '', [string] $Detail = '')
    $script:A2C.Data['PbiExportIndex'].Add([pscustomobject]@{ Ws = $Ws; ItemType = $ItemType; Item = $Item; ItemId = $ItemId; Kind = $Kind; Status = $Status
            Files = $Files; Bytes = $Bytes; Path = $Path; Detail = $Detail })
    if ($Status -eq 'Exported') { $script:A2C.Data['PbiExportMap'][$ItemId.ToLowerInvariant()] = $Path }
}

function Export-A2CPbiReportFile {
    # Downloads a report as .pbix (or .rdl for a paginated report). Returns the number of bytes written.
    param($Ws, $Report)
    $id = [string]$Report.id
    $paginated = ([string]$Report.reportType -eq 'PaginatedReport')
    $ext = $(if ($paginated) { '.rdl' } else { '.pbix' })
    $kind = $ext.TrimStart('.')
    $rel = (Get-A2CPbiExportDir $Ws 'reports' '') + '/' + (Get-A2CSafePathSegment ([string]$Report.name)) + $ext
    $path = [IO.Path]::Combine($script:A2C.OutDir, $rel)
    if (Test-Path -LiteralPath $path) {
        # two reports with the same name in one workspace
        $rel = (Get-A2CPbiExportDir $Ws 'reports' '') + '/' + (Get-A2CSafePathSegment ([string]$Report.name)) + '_' + $id.Substring(0, [math]::Min(8, $id.Length)) + $ext
        $path = [IO.Path]::Combine($script:A2C.OutDir, $rel)
    }
    $dir = [IO.Path]::GetDirectoryName($path)
    if (-not [IO.Directory]::Exists($dir)) { [void][IO.Directory]::CreateDirectory($dir) }
    $r = Invoke-A2CRequest -Uri ($script:A2CPbiBase + $Ws.PbiPrefix + '/reports/' + $id + '/Export') -Kind 'pbi' -MaxRetries 2 -OutFile $path
    if ($r.StatusCode -lt 200 -or $r.StatusCode -ge 300 -or -not (Test-Path -LiteralPath $path)) {
        if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue }
        Add-A2CPbiExport $Ws 'Report' ([string]$Report.name) $id $kind 'Failed' -Detail (Get-A2CShortText (Get-A2CErrorText (Get-A2CError $r)) 200)
        return 0
    }
    $len = (Get-Item -LiteralPath $path).Length
    Add-A2CPbiExport $Ws 'Report' ([string]$Report.name) $id $kind 'Exported' -Files 1 -Bytes $len -Path $rel
    return $len
}

function Export-A2CPbiContent {
    # Writes what the inventory can download below PowerBI/<workspace>: item definitions (notebooks, pipelines,
    # dataflows, semantic models, ...), dataflow (gen1) definitions and the SQL code read by FabricSql.
    # -WithReports adds the report files (.pbix / .rdl, they contain data) and the report definitions.
    param([switch] $WithReports, [string] $Area)
    Initialize-A2CPbiExport
    $done = $script:A2C.Data['PbiExportDone']
    $budget = [long]$script:A2C.Options.PbiExportMaxSizeMB * 1MB
    $used = [long]0
    $stopped = $false
    $wss = @($script:A2C.Data['PbiWorkspaces'])
    $t0 = [DateTime]::Now
    $i = 0
    foreach ($w in $wss) {
        $i++
        Write-A2CProgress -Activity $Area -Current $i -Total $wss.Count -Started $t0 -Item $w.Name
        if ($WithReports) {
            foreach ($x in $w.Reports) {
                if ($budget -gt 0 -and $used -ge $budget) {
                    $stopped = $true
                    Add-A2CPbiExport $w 'Report' ([string]$x.name) ([string]$x.id) 'pbix' 'Skipped' -Detail '-PbiExportMaxSizeMB reached'
                    continue
                }
                $used += Export-A2CPbiReportFile $w $x
            }
        }
        # item definitions that were not written yet (by the ItemDefinitions dataset or an earlier export), 20 at a time
        $todo = @($w.Items | Where-Object { $script:A2CPbiDefinitionTypes -contains [string]$_.type -and -not $done.ContainsKey([string]$_.id) -and ($WithReports -or [string]$_.type -ne 'Report') })
        for ($k = 0; $k -lt $todo.Count; $k += 20) {
            $chunk = @($todo[$k..([math]::Min($k + 19, $todo.Count - 1))])
            $res = Invoke-A2CFabricLroBatch -Paths @($chunk | ForEach-Object { Get-A2CPbiDefinitionPath $w $_ })
            for ($n = 0; $n -lt $chunk.Count; $n++) { Save-A2CPbiDefinitionParts $w $chunk[$n] $res[$n] }
        }
        $gen1 = $script:A2C.Data['PbiDataflowDefs']
        foreach ($x in $w.Dataflows) {
            # dataflow (gen1) definition: model.json
            $id = [string]$x.objectId
            if ($done.ContainsKey($id)) { continue }
            $done[$id] = $true
            $text = $null
            if ($null -ne $gen1 -and $gen1.ContainsKey($id)) { $text = [string]$gen1[$id].Text }
            else {
                $r = Invoke-A2CPbi ($w.PbiPrefix + '/dataflows/' + $id) -MaxRetries 2
                if (-not $r.Ok) { Add-A2CPbiExport $w 'Dataflow' ([string]$x.name) $id 'model.json' 'Failed' -Detail (Get-A2CShortText (Get-A2CErrorText $r.Error) 200); continue }
                $text = $r.Content
            }
            $rel = (Get-A2CPbiExportDir $w 'dataflows' '') + '/' + (Get-A2CSafePathSegment ([string]$x.name)) + '.json'
            Save-A2CExportFile $rel $text
            Add-A2CPbiExport $w 'Dataflow' ([string]$x.name) $id 'model.json' 'Exported' -Files 1 -Bytes $text.Length -Path $rel
        }
        # code of views, procedures and functions read by FabricSql: one .sql file per object
        $sql = $script:A2C.Data['FabricSql']
        if ($null -ne $sql) {
            foreach ($db in (@($w.Lakehouses) + @($w.Warehouses))) {
                $s = $sql[[string]$db.id]
                $key = 'sql:' + [string]$db.id
                if ($null -eq $s -or -not $s.Ok -or @($s.Modules).Count -eq 0 -or $done.ContainsKey($key)) { continue }
                $done[$key] = $true
                $dir = Get-A2CPbiExportDir $w 'sql' ([string]$db.displayName)
                $bytes = [long]0
                foreach ($m in $s.Modules) {
                    $text = [string]$m['definition']
                    Save-A2CExportFile ($dir + '/' + (Get-A2CSafePathSegment (([string]$m['schema']) + '.' + ([string]$m['name']) + '.sql'))) $text
                    $bytes += $text.Length
                }
                Add-A2CPbiExport $w ([string]$db.type) ([string]$db.displayName) $key 'sql' 'Exported' -Files @($s.Modules).Count -Bytes $bytes -Path $dir
            }
        }
    }
    Write-A2CProgress -Activity $Area -Completed
    $index = $script:A2C.Data['PbiExportIndex']
    $ok = @($index | Where-Object { $_.Status -eq 'Exported' }).Count
    $failed = @($index | Where-Object { $_.Status -eq 'Failed' }).Count
    $note = '{0} exported, {1} not exported (see exportIndex.csv)' -f $ok, $failed
    if ($WithReports) { $note += ', {0:N0} MB of report files' -f ($used / 1MB) }
    if ($stopped) { $note += '; report downloads stopped at -PbiExportMaxSizeMB' }
    $status = $(if ($failed -gt 0 -or $stopped) { 'Partial' } else { 'Collected' })
    Add-A2CCoverage -Area $Area -Status $status -Items $ok -Detail $note
    Write-A2CLog -Level Ok ($Area + ': ' + $note)
    return New-A2CDatasetResult -Status $status -Note $note
}

function Get-A2CDatasetFabricCode { return (Export-A2CPbiContent -Area 'Fabric code download') }

function Get-A2CDatasetPowerBIExport { return (Export-A2CPbiContent -WithReports -Area 'Power BI export') }
# ---- end Power BI ----

#endregion

#region ============================== 7. REPORT BUILDERS ===============================================

# ---- Power BI ----
# All files are written to the PowerBI folder. Workspace columns come first in every per-workspace file.
$script:A2CPbiWs = 'Workspace:name WorkspaceId:guid '
$script:A2CPbiDsCols = 'DatasourceType Name:name Server:fqdn Database:name Url:url Path:path ConnectionString:text GatewayId:guid GatewayName:name DatasourceId:guid ConnectionDetails:json'

Register-A2CSchema 'PowerBI/workspaces.csv' ($script:A2CPbiWs + 'Type Description:text CapacityId:guid Capacity:name CapacitySku CapacityRegion IsOnDedicatedCapacity IsReadOnly DomainId:guid ' +
    'DefaultStorageFormat InPowerBI InFabric Reports:num SemanticModels:num Dashboards:num Dataflows:num FabricItems:num Users:num Admins:upn GitState ' +
    'OneLakeBlobEndpoint:url OneLakeDfsEndpoint:url WorkspaceIdentity:guid PowerBIStatus PowerBIProblems:note FabricStatus FabricProblems:note')
Register-A2CSchema 'PowerBI/workspaceUsers.csv' ($script:A2CPbiWs + 'DisplayName:name Identifier:upn EmailAddress:upn PrincipalType AccessRight UserType GraphId:guid')
Register-A2CSchema 'PowerBI/reports.csv' ($script:A2CPbiWs + 'Report:name ReportId:guid ReportType Format ModelId:guid Model:name ModelWorkspaceId:guid Description:text ' +
    'IsOwnedByMe IsFromPbix AppId:guid Pages:num WebUrl:url ExportedTo:path')
Register-A2CSchema 'PowerBI/reportPages.csv' ($script:A2CPbiWs + 'Report:name ReportId:guid Page:name DisplayName:name Order:num')
Register-A2CSchema 'PowerBI/reportDatasources.csv' ($script:A2CPbiWs + 'Report:name ReportId:guid ' + $script:A2CPbiDsCols)
Register-A2CSchema 'PowerBI/semanticModels.csv' ($script:A2CPbiWs + 'Model:name ModelId:guid ConfiguredBy:upn CreatedDate:date IsRefreshable TargetStorageMode ' +
    'IsEffectiveIdentityRequired IsEffectiveIdentityRolesRequired IsOnPremGatewayRequired AddRowsApiEnabled QueryScaleOut:json Description:text ' +
    'Reports:num Datasources:num DatasourceTypes StorageModes Measures:num RlsRoles:num ScheduleEnabled LastRefreshStatus LastRefreshEnd:date WebUrl:url ExportedTo:path')
Register-A2CSchema 'PowerBI/semanticModelDatasources.csv' ($script:A2CPbiWs + 'Model:name ModelId:guid ' + $script:A2CPbiDsCols)
Register-A2CSchema 'PowerBI/semanticModelUsers.csv' ($script:A2CPbiWs + 'Model:name ModelId:guid Identifier:upn PrincipalType AccessRight')
Register-A2CSchema 'PowerBI/semanticModelParameters.csv' ($script:A2CPbiWs + 'Model:name ModelId:guid Parameter:name Type IsRequired CurrentValue:text SuggestedValues:text')
Register-A2CSchema 'PowerBI/refreshSchedules.csv' ($script:A2CPbiWs + 'Model:name ModelId:guid Kind Enabled Days Times Frequency:num LocalTimeZoneId NotifyOption')
Register-A2CSchema 'PowerBI/refreshHistory.csv' ($script:A2CPbiWs + 'Model:name ModelId:guid RequestId:guid RefreshType Status StartTime:date EndTime:date DurationSec:num Attempts:num Error:note')
Register-A2CSchema 'PowerBI/dashboards.csv' ($script:A2CPbiWs + 'Dashboard:name DashboardId:guid IsReadOnly AppId:guid Tiles:num WebUrl:url')
Register-A2CSchema 'PowerBI/dashboardTiles.csv' ($script:A2CPbiWs + 'Dashboard:name DashboardId:guid Tile:name TileId:guid ReportId:guid ModelId:guid RowSpan:num ColSpan:num')
Register-A2CSchema 'PowerBI/dataflows.csv' ($script:A2CPbiWs + 'Dataflow:name DataflowId:guid Description:text ConfiguredBy:upn ModifiedBy:upn ModifiedDateTime:date Generation:num ' +
    'Datasources:num LastRunStatus LastRunEnd:date ExportedTo:path')
Register-A2CSchema 'PowerBI/dataflowDatasources.csv' ($script:A2CPbiWs + 'Dataflow:name DataflowId:guid ' + $script:A2CPbiDsCols)
Register-A2CSchema 'PowerBI/dataflowTransactions.csv' ($script:A2CPbiWs + 'Dataflow:name DataflowId:guid TransactionId RefreshType Status StartTime:date EndTime:date')
Register-A2CSchema 'PowerBI/lineage.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid DependsOnType DependsOn:name DependsOnId:guid DependsOnWorkspaceId:guid')
Register-A2CSchema 'PowerBI/apps.csv' 'App:name AppId:guid Description:text PublishedBy:name LastUpdate:date WorkspaceId:guid Workspace:name Reports:num Dashboards:num'
Register-A2CSchema 'PowerBI/appContent.csv' 'App:name AppId:guid ContentType Name:name Id:guid ModelId:guid WebUrl:url'
Register-A2CSchema 'PowerBI/capacities.csv' 'Capacity:name CapacityId:guid Sku Region State UserAccessRight Admins:upn TenantKeyId:guid Workspaces:num Source'
Register-A2CSchema 'PowerBI/capacityWorkloads.csv' 'Capacity:name CapacityId:guid Workload State MaxMemoryPercentageSetByUser:num'
Register-A2CSchema 'PowerBI/refreshables.csv' ('Capacity:name CapacityId:guid Workspace:name WorkspaceId:guid Name:name Id:guid Kind RefreshCount:num RefreshFailures:num ' +
    'AverageDurationSec:num MedianDurationSec:num RefreshesPerDay:num LastRefreshStatus LastRefreshStart:date LastRefreshEnd:date ScheduleEnabled ScheduleDays ScheduleTimes')
Register-A2CSchema 'PowerBI/gateways.csv' 'Gateway:name GatewayId:guid Type Status Version Members:num Datasources:num Annotation:text Source Properties:json'
Register-A2CSchema 'PowerBI/gatewayDatasources.csv' 'Gateway:name GatewayId:guid Datasource:name DatasourceId:guid DatasourceType CredentialType ConnectionDetails:json'
Register-A2CSchema 'PowerBI/gatewayDatasourceUsers.csv' 'Gateway:name GatewayId:guid Datasource:name DatasourceId:guid DisplayName:name Identifier:upn EmailAddress:upn PrincipalType AccessRight'
Register-A2CSchema 'PowerBI/deploymentPipelines.csv' 'Pipeline:name PipelineId:guid Description:text Stages:num Workspaces:name'
Register-A2CSchema 'PowerBI/pipelineStages.csv' 'Pipeline:name PipelineId:guid Order:num WorkspaceId:guid Workspace:name'
Register-A2CSchema 'PowerBI/pipelineStageItems.csv' 'Pipeline:name PipelineId:guid Order:num ItemType Item:name ItemId:guid SourceItemId:guid TargetItemId:guid LastDeploymentTime:date'
Register-A2CSchema 'PowerBI/pipelineUsers.csv' 'Pipeline:name PipelineId:guid Identifier:upn PrincipalType AccessRight'
Register-A2CSchema 'PowerBI/pipelineOperations.csv' ('Pipeline:name PipelineId:guid OperationId:guid Type Status ExecutionStartTime:date ExecutionEndTime:date ' +
    'SourceStageOrder:num TargetStageOrder:num PerformedBy:upn Note:text')
Register-A2CSchema 'PowerBI/availableFeatures.csv' 'Feature State ExtendedState AdditionalInfo:json'

Register-A2CSchema 'PowerBI/fabricItems.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid Description:text FolderPath:path FolderId:guid LogicalId:guid Tags:name Connections:num Shortcuts:num Schedules:num Runs:num LastRunStatus LastRunStart:date ' +
    'Properties:json ExportedTo:path')
Register-A2CSchema 'PowerBI/workspaceRoleAssignments.csv' ($script:A2CPbiWs + 'Principal:name PrincipalId:guid PrincipalType Role UserPrincipalName:upn AppId:guid GroupType')
Register-A2CSchema 'PowerBI/workspaceGit.csv' ($script:A2CPbiWs + 'State Provider Organization:name Project:name Repository:name Branch:name Directory:path Head LastSyncTime:date ' +
    'ItemStatus:note UncommittedItems:num IncomingItems:num Conflicts:num RemoteCommitHash')
Register-A2CSchema 'PowerBI/gitItemStatus.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid LogicalId:guid WorkspaceChange RemoteChange ConflictType')
Register-A2CSchema 'PowerBI/folders.csv' ($script:A2CPbiWs + 'Folder:name FolderId:guid ParentFolderId:guid Path:path Items:num')
Register-A2CSchema 'PowerBI/workspaceNetwork.csv' ($script:A2CPbiWs + 'InboundDefaultAction OutboundDefaultAction Policy:json')
Register-A2CSchema 'PowerBI/sparkSessions.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid LivyId:guid SparkApplicationId State Origin JobType SubmittedDateTime:date StartDateTime:date ' +
    'EndDateTime:date TotalDurationSec:num Submitter:guid RuntimeVersion Properties:json')
Register-A2CSchema 'PowerBI/onelakeRoles.csv' ($script:A2CPbiWs + 'Lakehouse:name LakehouseId:guid Role:name Paths:path Actions EntraMembers:guid FabricItemMembers:text Definition:json')
Register-A2CSchema 'PowerBI/deploymentStages.csv' 'Pipeline:name PipelineId:guid Stage:name StageId:guid Order:num Workspace:name WorkspaceId:guid IsPublic Items:num Description:text'
Register-A2CSchema 'PowerBI/deploymentStageItems.csv' 'Pipeline:name PipelineId:guid Stage:name Order:num Workspace:name ItemType Item:name ItemId:guid SourceItemId:guid TargetItemId:guid LastDeploymentTime:date'
Register-A2CSchema 'PowerBI/deploymentRoleAssignments.csv' 'Pipeline:name PipelineId:guid PrincipalId:guid PrincipalType Role'
Register-A2CSchema 'PowerBI/domains.csv' 'Domain:name DomainId:guid Description:text ParentDomainId:guid'
Register-A2CSchema 'PowerBI/workspaceSparkSettings.csv' ($script:A2CPbiWs + 'DefaultPool:name DefaultPoolType RuntimeVersion Environment:name AutomaticLogEnabled CustomizeComputeEnabled Settings:json')
Register-A2CSchema 'PowerBI/sparkPools.csv' ($script:A2CPbiWs + 'Pool:name PoolId:guid Type NodeFamily NodeSize AutoScale:json DynamicExecutorAllocation:json')
Register-A2CSchema 'PowerBI/managedPrivateEndpoints.csv' ($script:A2CPbiWs + 'Name:name Id:guid TargetPrivateLinkResourceId:id TargetSubresourceType ProvisioningState ConnectionState:json')
Register-A2CSchema 'PowerBI/lakehouses.csv' ($script:A2CPbiWs + 'Lakehouse:name LakehouseId:guid OneLakeTablesPath:url OneLakeFilesPath:url SqlEndpointId:guid SqlEndpointConnectionString:fqdn ' +
    'SqlEndpointStatus DefaultSchema Tables:num TablesSource Views:num StoredProcedures:num Functions:num Shortcuts:num SqlStatus:note')
Register-A2CSchema 'PowerBI/lakehouseTables.csv' ($script:A2CPbiWs + 'Lakehouse:name LakehouseId:guid Schema:name Table:name Type Format Location:url Source')
Register-A2CSchema 'PowerBI/warehouses.csv' ($script:A2CPbiWs + 'Warehouse:name WarehouseId:guid Description:text ConnectionString:fqdn CreatedDate:date LastUpdatedTime:date ' +
    'Tables:num Views:num StoredProcedures:num Functions:num SqlStatus:note')
Register-A2CSchema 'PowerBI/shortcuts.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid Shortcut:name Path:path TargetType TargetConnection:name TargetConnectionId:guid TargetLocation:url ' +
    'TargetSubpath:path TargetWorkspace:name TargetItem:name Target:json')
Register-A2CSchema 'PowerBI/itemSchedules.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid JobType ScheduleId:guid Enabled CreatedDateTime:date Owner:name Configuration:json')
Register-A2CSchema 'PowerBI/itemConnections.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid Connection:name ConnectionId:guid ConnectivityType ConnectionType Path:text GatewayId:guid')
Register-A2CSchema 'PowerBI/connections.csv' ('Connection:name ConnectionId:guid ConnectivityType ConnectionType Path:text GatewayId:guid PrivacyLevel CredentialType ' +
    'SingleSignOnType ConnectionEncryption SkipTestConnection')

Register-A2CSchema 'PowerBI/modelTables.csv' ($script:A2CPbiWs + 'Model:name ModelId:guid Table:name StorageMode IsHidden DataCategory Description:text Expression:text Properties:json')
Register-A2CSchema 'PowerBI/modelColumns.csv' ($script:A2CPbiWs + 'Model:name ModelId:guid Table:name Column:name DataType Type IsHidden IsKey FormatString DisplayFolder:name ' +
    'Description:text Expression:text Properties:json')
Register-A2CSchema 'PowerBI/modelMeasures.csv' ($script:A2CPbiWs + 'Model:name ModelId:guid Table:name Measure:name DataType FormatString IsHidden DisplayFolder:name State ' +
    'Description:text Expression:text Properties:json')
Register-A2CSchema 'PowerBI/modelRelationships.csv' ($script:A2CPbiWs + 'Model:name ModelId:guid Relationship:name FromTable:name FromColumn:name FromCardinality ToTable:name ToColumn:name ' +
    'ToCardinality IsActive CrossFilteringBehavior Properties:json')

Register-A2CSchema 'PowerBI/adminWorkspaces.csv' ($script:A2CPbiWs + 'Type Scope IsPersonal State IsOnDedicatedCapacity CapacityId:guid Description:text DefaultStorageFormat ' +
    'Users:num Admins:upn AdminsDeparted:num NoAdmin Items:num HasContent LastItemChange:date ChangedLast30Days RecentActivities:num ActiveUsers:num DatasourceTypes Datasources:num UsesGateway ' +
    'CreatedDate:date ModifiedDate:date')
Register-A2CSchema 'PowerBI/workspaceScopeSummary.csv' 'Measure:note Count:num Note:note'
Register-A2CSchema 'PowerBI/repointingMap.csv' ($script:A2CPbiWs + 'Scope DatasourceType Kind Host:fqdn Target:url Database:name GatewayId:guid UsesGateway Items:num ItemNames:name ' +
    'Models:num ScheduledModels:num RefreshSchedules ScheduleKnown')
Register-A2CSchema 'PowerBI/capacityFootprint.csv' ('Capacity:name CapacityId:guid Sku Region State Workspaces:num SharedWithContent:num Items:num ScheduledModels:num ' +
    'RefreshesPerDay:num AvgRefreshSec:num ActiveUsers:num Note:note')
Register-A2CSchema 'PowerBI/principals.csv' 'Principal:name Identifier:upn ObjectId:guid PrincipalType AppId:guid Workspaces:num Roles Sources EntraStatus:note Departed'
Register-A2CSchema 'PowerBI/adminWorkspaceUsers.csv' ($script:A2CPbiWs + 'DisplayName:name Identifier:upn EmailAddress:upn PrincipalType AccessRight UserType GraphId:guid')
Register-A2CSchema 'PowerBI/adminItems.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid ConfiguredBy:upn CreatedDate:date ModifiedDate:date ModifiedBy:upn Endorsement CertifiedBy:upn ' +
    'SensitivityLabelId:guid ModelId:guid TargetStorageMode ContentProviderType Description:text Users:num DatasourceIds:text Upstream:text')
Register-A2CSchema 'PowerBI/adminItemUsers.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid DisplayName:name Identifier:upn PrincipalType AccessRight')
Register-A2CSchema 'PowerBI/adminDatasourceInstances.csv' 'DatasourceId:guid DatasourceType GatewayId:guid Workspaces:num ConnectionDetails:json'
Register-A2CSchema 'PowerBI/adminModelTables.csv' ($script:A2CPbiWs + 'Model:name ModelId:guid Table:name IsHidden StorageMode Columns:num Measures:num Source:text')
Register-A2CSchema 'PowerBI/adminModelExpressions.csv' ($script:A2CPbiWs + 'Model:name ModelId:guid Kind Table:name Name:name Expression:text Description:text')
Register-A2CSchema 'PowerBI/tenantSettings.csv' ('SettingName Title:text TenantSettingGroup Enabled CanSpecifySecurityGroups EnabledSecurityGroups:text ExcludedSecurityGroups:text ' +
    'DelegateToCapacity DelegateToDomain DelegateToWorkspace Properties:json')
Register-A2CSchema 'PowerBI/adminCapacities.csv' 'Capacity:name CapacityId:guid Sku Region State Admins:upn TenantKeyId:guid'
Register-A2CSchema 'PowerBI/adminApps.csv' 'App:name AppId:guid Description:text PublishedBy:name LastUpdate:date WorkspaceId:guid'
Register-A2CSchema 'PowerBI/widelyShared.csv' 'Kind ArtifactType Artifact:name ArtifactId:guid Sharer:upn SharerType ShareType AccessRight'
Register-A2CSchema 'PowerBI/adminDomains.csv' 'Domain:name DomainId:guid Description:text ParentDomainId:guid ContributorsScope'
Register-A2CSchema 'PowerBI/activityEvents.csv' ('CreationTime:date UserId:upn Activity Operation Workspace:name WorkspaceId:guid ItemName:name ItemType ObjectId:text ' +
    'ModelName:name ReportName:name ClientIP:ip UserAgent:text IsSuccess RequestId:guid Id:guid Properties:json')
Register-A2CSchema 'PowerBI/exportIndex.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid Kind Status Files:num Bytes:num Path:path Detail:note')

# ---- Parsers of item definitions (pipelines, notebooks, dataflows, semantic models) ----
$script:A2CGuidRx = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'

function Initialize-A2CPbiItemIndex {
    # id -> @{ Type; Name; Ws; WsId } of every item, SQL endpoint and workspace in scope, to turn the IDs found in
    # definitions into names; and "server|database" -> lakehouse or warehouse for SQL endpoint references.
    if ($null -ne $script:A2C.Data['PbiItemIndex']) { return }
    $ix = @{}
    $sql = @{}
    foreach ($w in @($script:A2C.Data['PbiWorkspaces'])) {
        $ix[$w.Id.ToLowerInvariant()] = [pscustomobject]@{ Type = 'Workspace'; Name = $w.Name; Ws = $w.Name; WsId = $w.Id; Id = $w.Id }
        foreach ($x in $w.Datasets) { $ix[([string]$x.id).ToLowerInvariant()] = [pscustomobject]@{ Type = 'SemanticModel'; Name = [string]$x.name; Ws = $w.Name; WsId = $w.Id; Id = [string]$x.id } }
        foreach ($x in $w.Reports) { $ix[([string]$x.id).ToLowerInvariant()] = [pscustomobject]@{ Type = 'Report'; Name = [string]$x.name; Ws = $w.Name; WsId = $w.Id; Id = [string]$x.id } }
        foreach ($x in $w.Dataflows) { $ix[([string]$x.objectId).ToLowerInvariant()] = [pscustomobject]@{ Type = 'Dataflow'; Name = [string]$x.name; Ws = $w.Name; WsId = $w.Id; Id = [string]$x.objectId } }
        foreach ($x in $w.Items) { $ix[([string]$x.id).ToLowerInvariant()] = [pscustomobject]@{ Type = [string]$x.type; Name = [string]$x.displayName; Ws = $w.Name; WsId = $w.Id; Id = [string]$x.id } }
        foreach ($x in $w.Lakehouses) {
            $ref = [pscustomobject]@{ Type = 'Lakehouse'; Name = [string]$x.displayName; Ws = $w.Name; WsId = $w.Id; Id = [string]$x.id }
            $ep = $x.properties.sqlEndpointProperties
            if ($ep.id) { $ix[([string]$ep.id).ToLowerInvariant()] = $ref }
            if ($ep.connectionString) { $sql[([string]$ep.connectionString + '|' + [string]$x.displayName).ToLowerInvariant()] = $ref }
        }
        foreach ($x in $w.Warehouses) {
            $ref = [pscustomobject]@{ Type = 'Warehouse'; Name = [string]$x.displayName; Ws = $w.Name; WsId = $w.Id; Id = [string]$x.id }
            if ($x.properties.connectionString) { $sql[([string]$x.properties.connectionString + '|' + [string]$x.displayName).ToLowerInvariant()] = $ref }
        }
    }
    $conn = @{}
    $f = $script:A2C.Data['FabricTenant']
    if ($null -ne $f) { foreach ($c in $f.Connections) { $conn[([string]$c.id).ToLowerInvariant()] = $c } }
    $script:A2C.Data['PbiItemIndex'] = $ix
    $script:A2C.Data['PbiSqlIndex'] = $sql
    $script:A2C.Data['PbiConnectionIndex'] = $conn
}

function Get-A2CPbiItemRef {
    # The indexed item, workspace or SQL endpoint with this ID; $null when it is not in scope.
    param([string] $Id)
    if (-not $Id) { return $null }
    return $script:A2C.Data['PbiItemIndex'][$Id.ToLowerInvariant()]
}

function Get-A2CPbiItemRefs {
    # Every in-scope item (not workspace) whose ID occurs in a text, in order of appearance, without repeats.
    param([string] $Text, [string] $ExceptId = '')
    $out = [System.Collections.Generic.List[object]]::new()
    $seen = @{}
    foreach ($m in [regex]::Matches([string]$Text, $script:A2CGuidRx)) {
        $k = $m.Value.ToLowerInvariant()
        if ($seen.ContainsKey($k) -or $k -eq $ExceptId.ToLowerInvariant()) { continue }
        $seen[$k] = $true
        $ref = $script:A2C.Data['PbiItemIndex'][$k]
        if ($null -ne $ref -and $ref.Type -ne 'Workspace') { $out.Add($ref) }
    }
    return , $out
}

function Get-A2CPipelineActivities {
    # Activities of a pipeline definition with the ones inside If, ForEach, Until and Switch: @{ Activity; Parent }
    param($Activities, [string] $Parent = '')
    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($a in @($Activities)) {
        if ($null -eq $a) { continue }
        $out.Add([pscustomobject]@{ Activity = $a; Parent = $Parent })
        $tp = $a.typeProperties
        if ($null -eq $tp) { continue }
        $name = [string]$a.name
        foreach ($k in @('activities', 'ifTrueActivities', 'ifFalseActivities', 'defaultActivities')) {
            if ($null -ne $tp.PSObject.Properties[$k]) { foreach ($c in (Get-A2CPipelineActivities $tp.$k $name)) { $out.Add($c) } }
        }
        if ($null -ne $tp.PSObject.Properties['cases']) {
            foreach ($case in @($tp.cases)) { if ($null -ne $case) { foreach ($c in (Get-A2CPipelineActivities $case.activities ($name + ' / ' + [string]$case.value))) { $out.Add($c) } } }
        }
    }
    return , $out
}

$script:A2CPipelineRefKeys = [ordered]@{ notebookId = 'Notebook'; dataflowId = 'Dataflow'; datasetId = 'SemanticModel'; pipelineId = 'DataPipeline'
    sqlEndpointId = 'SQLEndpoint'; lakehouseId = 'Lakehouse'; warehouseId = 'Warehouse'; artifactId = 'Item'; sparkJobDefinitionId = 'SparkJobDefinition'; copyJobId = 'CopyJob' }
$script:A2CPipelineContainers = @('activities', 'ifTrueActivities', 'ifFalseActivities', 'defaultActivities', 'cases')

function Get-A2CPipelineActivityInfo {
    # What one activity runs or touches: the referenced items, stored procedure, parameters and connection.
    param($Activity)
    $tp = $Activity.typeProperties
    $own = [ordered]@{}      # typeProperties without the nested activities of a container
    if ($null -ne $tp) { foreach ($p in $tp.PSObject.Properties) { if ($script:A2CPipelineContainers -notcontains $p.Name) { $own[$p.Name] = $p.Value } } }
    $text = (ConvertTo-A2CJsonString $own) + ' ' + (ConvertTo-A2CJsonString $Activity.externalReferences) + ' ' + (ConvertTo-A2CJsonString $Activity.linkedService) +
        ' ' + (ConvertTo-A2CJsonString $Activity.connectionSettings)
    $refs = [System.Collections.Generic.List[object]]::new()
    foreach ($r in (Get-A2CPbiItemRefs $text)) { $refs.Add($r) }
    # an ID under a known key that is not in scope (another workspace, or no access): keep it, without a name
    foreach ($k in $script:A2CPipelineRefKeys.Keys) {
        $v = [string]$own[$k]
        if ($v -match ('^' + $script:A2CGuidRx + '$') -and $null -eq (Get-A2CPbiItemRef $v)) {
            $refs.Add([pscustomobject]@{ Type = $script:A2CPipelineRefKeys[$k]; Name = ''; Ws = ''; WsId = [string](Get-A2CFirst $tp @('workspaceId', 'groupId')); Id = $v })
        }
    }
    $conn = ''
    foreach ($m in [regex]::Matches($text, $script:A2CGuidRx)) {
        $c = $script:A2C.Data['PbiConnectionIndex'][$m.Value.ToLowerInvariant()]
        if ($null -ne $c) { $conn = [string]$c.displayName; if (-not $conn) { $conn = [string]$c.id }; break }
    }
    $params = [System.Collections.Generic.List[string]]::new()
    foreach ($pk in @('storedProcedureParameters', 'parameters')) {
        $po = $own[$pk]
        if ($null -eq $po -or $po -is [string]) { continue }
        foreach ($p in $po.PSObject.Properties) {
            $v = $p.Value
            if ($null -ne $v -and $v -isnot [string] -and $null -ne $v.PSObject.Properties['value']) { $v = $v.value }
            $params.Add(('{0}={1}' -f $p.Name, (Format-A2CCell $v)))
        }
    }
    $detail = ''
    if ($own.Contains('waitTimeInSeconds')) { $detail = 'wait ' + (Format-A2CCell $own['waitTimeInSeconds']) + ' s' }
    elseif ($null -ne $own['source'] -and $null -ne $own['sink']) { $detail = ('{0} -> {1}' -f $own['source'].type, $own['sink'].type) }
    elseif ($null -ne $own['expression']) { $detail = Get-A2CShortText ([string](Get-A2CFirst $own['expression'] @('value'))) 300 }
    return [pscustomobject]@{ Refs = $refs; Connection = $conn; Procedure = [string]$own['storedProcedureName']; Parameters = [string]::Join(', ', $params); Detail = $detail; Own = $own }
}

function Get-A2CNotebookCode {
    # Text of a notebook definition: the Fabric source format as it is, an .ipynb as its cell sources.
    param($Parts)
    foreach ($k in @($Parts.Keys)) {
        if ($k -like '*.ipynb') {
            try {
                $nb = ConvertFrom-A2CJson $Parts[$k]
                $sb = New-Object System.Text.StringBuilder
                foreach ($c in @($nb.cells)) { if ($null -ne $c -and [string]$c.cell_type -eq 'code') { [void]$sb.AppendLine((Join-A2C $c.source '')) } }
                $meta = ''
                if ($null -ne $nb.metadata) { $meta = '# META ' + (ConvertTo-A2CJsonString $nb.metadata) + "`n" }
                return ($meta + $sb.ToString())
            }
            catch { return '' }
        }
        return [string]$Parts[$k]
    }
    return ''
}

function Get-A2CNotebookDependencies {
    # What a notebook is attached to and what its code reads, writes and calls. The attached lakehouses come from
    # the notebook metadata (exact); tables, paths and calls are found by pattern in the code (comments skipped),
    # so names built at run time appear as written, e.g. "{schema}.{tablename}". Returns @{ Relation; Target; Id; WorkspaceId; Evidence }
    param([string] $Code)
    $out = [System.Collections.Generic.List[object]]::new()
    $seen = @{}
    $add = {
        param($Relation, $Target, $Id, $WorkspaceId, $Evidence)
        $key = ($Relation + '|' + $Target + '|' + $Id).ToLowerInvariant()
        if (-not $Target -and -not $Id) { return }
        if ($seen.ContainsKey($key)) { return }
        $seen[$key] = $true
        $out.Add([pscustomobject]@{ Relation = $Relation; Target = [string]$Target; Id = [string]$Id; WorkspaceId = [string]$WorkspaceId; Evidence = (Get-A2CShortText ([string]$Evidence) 160) })
    }
    $lines = $Code -split "`r?`n"
    # the first META block holds the notebook-level settings
    $meta = New-Object System.Text.StringBuilder
    $inMeta = $false
    foreach ($l in $lines) {
        if ($l -match '^\s*(?:#|--|//)\s*META(?:\s(.*))?$') { $inMeta = $true; [void]$meta.AppendLine([string]$Matches[1]); continue }   # not "# METADATA ****"
        if ($inMeta) { break }
    }
    if ($meta.Length -gt 0) {
        $m = $null
        try { $m = ConvertFrom-A2CJson $meta.ToString() } catch { $m = $null }
        $lh = $m.dependencies.lakehouse
        if ($null -ne $lh) {
            if ($lh.default_lakehouse) { & $add 'DefaultLakehouse' ([string]$lh.default_lakehouse_name) ([string]$lh.default_lakehouse) ([string]$lh.default_lakehouse_workspace_id) 'notebook metadata' }
            foreach ($k in @($lh.known_lakehouses)) { if ($null -ne $k -and $k.id -and [string]$k.id -ne [string]$lh.default_lakehouse) { & $add 'AttachedLakehouse' '' ([string]$k.id) '' 'notebook metadata' } }
        }
        $wh = $m.dependencies.warehouse
        if ($null -ne $wh) {
            if ($wh.default_warehouse) { & $add 'DefaultWarehouse' '' ([string]$wh.default_warehouse) '' 'notebook metadata' }
            foreach ($k in @($wh.known_warehouses)) { if ($null -ne $k -and $k.id -and [string]$k.id -ne [string]$wh.default_warehouse) { & $add 'AttachedWarehouse' '' ([string]$k.id) '' 'notebook metadata' } }
        }
        $env = $m.dependencies.environment
        if ($null -ne $env -and $env.environmentId) { & $add 'Environment' '' ([string]$env.environmentId) ([string]$env.workspaceId) 'notebook metadata' }
    }
    $q = '["'']'
    $name = 'f?' + $q + '([^"'']+)' + $q
    foreach ($raw in $lines) {
        $l = $raw.Trim()
        if (-not $l -or $l -match '^(?:#|--|//)\s*META(\s|$)') { continue }
        $magic = ($l -match '^(?:#|--|//)\s*MAGIC\s?(.*)$')
        if ($magic) { $l = $Matches[1].Trim() }
        elseif ($l.StartsWith('#') -or $l.StartsWith('//')) { continue }
        foreach ($m in [regex]::Matches($l, '(?:spark\.read\.table|spark\.table|DeltaTable\.forName\(\s*spark\s*,|\.read\.format\([^)]*\)\.table)\(?\s*' + $name)) { & $add 'ReadsTable' $m.Groups[1].Value '' '' $l }
        foreach ($m in [regex]::Matches($l, '(?:saveAsTable|insertInto)\(\s*' + $name)) { & $add 'WritesTable' $m.Groups[1].Value '' '' $l }
        if ($l -notmatch '^(?:from|import)\s') {
            foreach ($m in [regex]::Matches($l, '(?i)\b(?:INSERT\s+INTO|INSERT\s+OVERWRITE(?:\s+TABLE)?|MERGE\s+INTO|CREATE\s+(?:OR\s+REPLACE\s+)?TABLE(?:\s+IF\s+NOT\s+EXISTS)?)\s+([A-Za-z_][\w]*(?:\.[A-Za-z_][\w]*){0,2})')) {
                & $add 'WritesTable' $m.Groups[1].Value '' '' $l
            }
            # FROM / JOIN only with a schema.table name: a bare word after "from" is too often not a table
            foreach ($m in [regex]::Matches($l, '(?i)\b(?:FROM|JOIN)\s+([A-Za-z_][\w]*(?:\.[A-Za-z_][\w]*){1,2})\b')) { & $add 'ReadsTable' $m.Groups[1].Value '' '' $l }
        }
        foreach ($m in [regex]::Matches($l, '(?:notebookutils|mssparkutils)\.notebook\.run\(\s*' + $name)) { & $add 'RunsNotebook' $m.Groups[1].Value '' '' $l }
        if ($l -match '^%run\s+(\S+)') { & $add 'RunsNotebook' $Matches[1] '' '' $l }
        foreach ($m in [regex]::Matches($l, 'abfss://([^@"''\s]+)@onelake\.dfs\.fabric\.microsoft\.com/([^/"''\s{]+)(/[^"''\s]*)?')) {
            & $add 'OneLakePath' ($m.Groups[1].Value + '/' + $m.Groups[2].Value + $m.Groups[3].Value) $m.Groups[2].Value $m.Groups[1].Value $l
        }
        foreach ($m in [regex]::Matches($l, '(?:abfss?|wasbs?)://[^"''\s]+')) { if ($m.Value -notmatch 'onelake\.') { & $add 'StoragePath' $m.Value '' '' $l } }
        foreach ($m in [regex]::Matches($l, 'https?://[^"''\s)>\]]+')) { if ($m.Value -notmatch 'onelake\.') { & $add 'Url' $m.Value.TrimEnd('.', ',', ';') '' '' $l } }
    }
    return , $out
}

$script:A2CMConnectorRx = '\b((?:SharePoint|Sql|Web|Excel|Lakehouse|Warehouse|Fabric|PowerPlatform|PowerBI|AzureStorage|OData|Folder|File|Salesforce|Dataverse|CommonDataService|Odbc|OleDb|Oracle|Snowflake|' +
    'AnalysisServices|Databricks|GoogleBigQuery|PostgreSQL|MySQL|Sap\w*|AzureDataExplorer|Kusto|Exchange|ActiveDirectory|Access|Cube|DataLake|Hdfs|AmazonRedshift|Teradata|Sybase|DB2|Pdf|Xml|Csv|Json)' +
    '\.(?!Type\b|From\b|FromValue\b)[A-Za-z]+)\s*\(\s*("(?:[^"]|"")*")?(?:\s*,\s*("(?:[^"]|"")*"))?'
$script:A2CMParsers = @('Csv.Document', 'Json.Document', 'Xml.Document', 'Xml.Tables', 'Excel.Workbook', 'Pdf.Tables', 'Json.FromValue', 'Web.Page', 'Excel.CurrentWorkbook', 'File.Contents', 'Web.BrowserContents')

function Get-A2CMSources {
    # Data sources of a Power Query (M) document: one row per query and connector call, with the file, folder and
    # Fabric item that the query then navigates to. Returns @{ Query; Connector; Target; Database; File; FolderPath; WorkspaceId; ItemId }
    param([string] $Text, [string] $DefaultQuery = '')
    $out = [System.Collections.Generic.List[object]]::new()
    if (-not $Text) { return , $out }
    $queries = [System.Collections.Generic.List[object]]::new()
    $ms = [regex]::Matches($Text, '(?m)^\s*shared\s+(#"(?:[^"]|"")+"|[\w\.]+)\s*=')
    if ($ms.Count -eq 0) { $queries.Add(@($DefaultQuery, $Text)) }
    for ($i = 0; $i -lt $ms.Count; $i++) {
        $end = $(if ($i + 1 -lt $ms.Count) { $ms[$i + 1].Index } else { $Text.Length })
        $qn = $ms[$i].Groups[1].Value
        if ($qn.StartsWith('#"')) { $qn = $qn.Substring(2, $qn.Length - 3).Replace('""', '"') }
        $queries.Add(@($qn, $Text.Substring($ms[$i].Index, $end - $ms[$i].Index)))
    }
    $unq = { param($s) if ($s) { $s.Substring(1, $s.Length - 2).Replace('""', '"') } else { '' } }
    foreach ($q in $queries) {
        $body = [string]$q[1]
        $file = ''; $folder = ''; $navName = ''
        $m = [regex]::Match($body, '#"Folder Path"\s*=\s*("(?:[^"]|"")*")'); if ($m.Success) { $folder = & $unq $m.Groups[1].Value }
        $m = [regex]::Match($body, '\[\s*Name\s*=\s*("(?:[^"]|"")*")'); if ($m.Success) { $navName = & $unq $m.Groups[1].Value }
        if ($folder -or $navName -match '\.\w{2,5}$') { $file = $navName }
        $ws = ''; $item = ''
        $m = [regex]::Match($body, 'workspaceId\s*=\s*"(' + $script:A2CGuidRx + ')"'); if ($m.Success) { $ws = $m.Groups[1].Value }
        $m = [regex]::Match($body, '(?:lakehouseId|warehouseId|dataflowId|datamartId|artifactId|itemId)\s*=\s*"(' + $script:A2CGuidRx + ')"'); if ($m.Success) { $item = $m.Groups[1].Value }
        foreach ($c in [regex]::Matches($body, $script:A2CMConnectorRx)) {
            $connector = $c.Groups[1].Value
            if ($script:A2CMParsers -contains $connector -and -not $c.Groups[2].Success) { continue }   # a parser around another source
            $db = & $unq $c.Groups[3].Value
            if (-not $db -and -not $file -and $connector -like 'Sql.*') { $db = $navName }
            $out.Add([pscustomobject]@{ Query = [string]$q[0]; Connector = $connector; Target = (& $unq $c.Groups[2].Value); Database = $db; File = $file; FolderPath = $folder
                    WorkspaceId = $ws; ItemId = $item })
        }
    }
    return , $out
}

function Get-A2CTmdlObjects {
    # Objects of a TMDL file that the inventory uses: tables, measures, columns, partitions, expressions, roles and
    # table permissions, each with its expression (text after "=", also over several lines or in a ``` block) and
    # its properties. Returns @{ Kind; Name; Table; Value; Props }
    param([string] $Text)
    $out = [System.Collections.Generic.List[object]]::new()
    $lines = $Text -split "`r?`n"
    $n = $lines.Count
    $indentOf = { param($s) $t = 0; foreach ($ch in $s.ToCharArray()) { if ($ch -eq "`t") { $t += 4 } elseif ($ch -eq ' ') { $t++ } else { break } }; [int][math]::Floor($t / 4) }
    $rx = '^(table|measure|column|partition|expression|role|tablePermission)\s+(''(?:[^'']|'''')*''|[^\s=]+)\s*(?:=\s*(.*))?$'
    $table = ''
    $cur = $null
    $i = 0
    while ($i -lt $n) {
        $line = $lines[$i]
        $t = $line.Trim()
        if (-not $t) { $i++; continue }
        $ind = & $indentOf $line
        $m = [regex]::Match($t, $rx)
        # collects the lines of a block that is indented deeper than $Level (blank lines inside it included)
        $block = {
            param([int] $Level)
            $sb = [System.Collections.Generic.List[string]]::new()
            while ($i + 1 -lt $n) {
                $nx = $lines[$i + 1]
                if ($nx.Trim() -and (& $indentOf $nx) -le $Level) { break }
                $sb.Add($nx.Trim())
                $i++
            }
            while ($sb.Count -gt 0 -and -not $sb[$sb.Count - 1]) { $sb.RemoveAt($sb.Count - 1) }
            [string]::Join("`n", $sb)
        }
        if ($m.Success) {
            $name = $m.Groups[2].Value
            if ($name.StartsWith("'")) { $name = $name.Substring(1, $name.Length - 2).Replace("''", "'") }
            $kind = $m.Groups[1].Value
            if ($kind -eq 'table' -or $kind -eq 'role') { $table = $name }
            $value = $m.Groups[3].Value.Trim()
            if ($kind -ne 'table' -and $kind -ne 'role' -and $m.Groups[3].Success -and $kind -ne 'partition') {
                if ($value.StartsWith('```')) {
                    $sb = [System.Collections.Generic.List[string]]::new()
                    $rest = $value.Substring(3)
                    if ($rest.Trim()) { $sb.Add($rest) }
                    while ($i + 1 -lt $n) { $i++; $x = $lines[$i]; $p = $x.IndexOf('```'); if ($p -ge 0) { if ($x.Substring(0, $p).Trim()) { $sb.Add($x.Substring(0, $p).Trim()) }; break }; $sb.Add($x.Trim()) }
                    $value = [string]::Join("`n", $sb)
                }
                else {
                    $more = . $block ($ind + 1)
                    if ($more) { $value = ($value + "`n" + $more).Trim() }
                }
            }
            $cur = [pscustomobject]@{ Kind = $kind; Name = $name; Table = $(if ($kind -eq 'table' -or $kind -eq 'role') { '' } else { $table }); Value = $value; Props = @{}; Indent = $ind }
            $out.Add($cur)
            $i++
            continue
        }
        if ($null -ne $cur -and $ind -eq $cur.Indent + 1) {
            $pm = [regex]::Match($t, '^([A-Za-z_]\w*)\s*(?::\s*(.*)|=\s*(.*))?$')
            if ($pm.Success) {
                $pn = $pm.Groups[1].Value
                if ($pm.Groups[2].Success) { $cur.Props[$pn] = $pm.Groups[2].Value.Trim() }
                elseif ($pn -ne 'annotation' -and $pn -ne 'extendedProperty') {
                    $v = $pm.Groups[3].Value.Trim()
                    $more = . $block $ind
                    $cur.Props[$pn] = ($v + "`n" + $more).Trim()
                }
            }
        }
        $i++
    }
    return , $out
}

function Get-A2CBimObjects {
    # The same objects as Get-A2CTmdlObjects from a model.bim (TMSL JSON) definition.
    param([string] $Text)
    $out = [System.Collections.Generic.List[object]]::new()
    $model = $null
    try { $model = (ConvertFrom-A2CJson $Text).model } catch { return , $out }
    $txt = { param($v) if ($null -eq $v) { '' } elseif ($v -is [string]) { $v } else { [string]::Join("`n", @($v)) } }
    foreach ($e in @($model.expressions)) { if ($null -ne $e) { $out.Add([pscustomobject]@{ Kind = 'expression'; Name = [string]$e.name; Table = ''; Value = (& $txt $e.expression); Props = @{} }) } }
    foreach ($t in @($model.tables)) {
        if ($null -eq $t) { continue }
        $tn = [string]$t.name
        $out.Add([pscustomobject]@{ Kind = 'table'; Name = $tn; Table = ''; Value = ''; Props = @{} })
        foreach ($x in @($t.measures)) { if ($null -ne $x) { $out.Add([pscustomobject]@{ Kind = 'measure'; Name = [string]$x.name; Table = $tn; Value = (& $txt $x.expression); Props = @{ formatString = [string]$x.formatString } }) } }
        foreach ($x in @($t.columns)) { if ($null -ne $x -and $x.expression) { $out.Add([pscustomobject]@{ Kind = 'column'; Name = [string]$x.name; Table = $tn; Value = (& $txt $x.expression); Props = @{} }) } }
        foreach ($x in @($t.partitions)) {
            if ($null -eq $x) { continue }
            $src = $x.source
            $props = @{ mode = [string]$x.mode; source = (& $txt (Get-A2CFirst $src @('expression', 'query'))) }
            if ($src.entityName) { $props['source'] = ('entityName: {0}' -f $src.entityName) + "`n" + ('schemaName: {0}' -f $src.schemaName) + "`n" + ('expressionSource: {0}' -f $src.expressionSource) }
            $out.Add([pscustomobject]@{ Kind = 'partition'; Name = [string]$x.name; Table = $tn; Value = [string]$src.type; Props = $props })
        }
    }
    foreach ($r in @($model.roles)) {
        if ($null -eq $r) { continue }
        $out.Add([pscustomobject]@{ Kind = 'role'; Name = [string]$r.name; Table = ''; Value = ''; Props = @{ modelPermission = [string]$r.modelPermission } })
        foreach ($tp in @($r.tablePermissions)) { if ($null -ne $tp) { $out.Add([pscustomobject]@{ Kind = 'tablePermission'; Name = [string]$tp.name; Table = [string]$r.name; Value = (& $txt $tp.filterExpression); Props = @{} }) } }
    }
    return , $out
}

function Get-A2CPbiModelObjects {
    # Parsed definition of a semantic model (cached); an empty list when the definition was not read.
    param([string] $ModelId)
    $cache = $script:A2C.Data['PbiModelObjects']
    if ($null -eq $cache) { $cache = @{}; $script:A2C.Data['PbiModelObjects'] = $cache }
    if ($cache.ContainsKey($ModelId)) { return , $cache[$ModelId] }
    $all = [System.Collections.Generic.List[object]]::new()
    $defs = $script:A2C.Data['PbiDefs']
    if ($null -ne $defs -and $defs.ContainsKey($ModelId)) {
        $parts = $defs[$ModelId].Parts
        foreach ($k in @($parts.Keys | Sort-Object)) {
            $objs = $(if ($k -eq 'model.bim') { Get-A2CBimObjects $parts[$k] } else { Get-A2CTmdlObjects $parts[$k] })
            foreach ($o in $objs) { $all.Add($o) }
        }
    }
    $cache[$ModelId] = $all
    return , $all
}

function Get-A2CPbiTargetKind {
    # Kind of an external target and whether it lives in this Fabric tenant, from the connector and the address.
    param([string] $Connector, [string] $Target)
    $t = $Target.ToLowerInvariant()
    if ($t -match 'sharepoint\.com|sharepoint-df\.com') { return @('SharePoint', 'External') }
    if ($t -match 'datawarehouse\.fabric\.microsoft\.com|pbidedicated\.windows\.net|\.sql\.azuresynapse\.net') { return @('Fabric SQL endpoint', 'Fabric (this tenant)') }
    if ($t -match 'onelake\.') { return @('OneLake', 'Fabric (this tenant)') }
    if ($Connector -match '^(Lakehouse|Warehouse|Fabric|PowerPlatform|PowerBI)\.') { return @('Fabric item', 'Fabric (this tenant)') }
    if ($t -match 'database\.windows\.net|\.sql\.') { return @('SQL', 'External') }
    if ($t -match 'dfs\.core\.windows\.net|blob\.core\.windows\.net') { return @('Azure Storage', 'External') }
    if ($t -match '\.dynamics\.com|dataverse') { return @('Dataverse', 'External') }
    if ($Connector -match '^(Sql|Oracle|Odbc|OleDb|PostgreSQL|MySQL|Snowflake|Teradata|DB2|Sybase|AmazonRedshift|GoogleBigQuery|Databricks)') { return @('SQL', 'External') }
    if ($Connector -match '^(Dataverse|CommonDataService)') { return @('Dataverse', 'External') }
    if ($Connector -match '^Salesforce') { return @('Salesforce', 'External') }
    if ($Connector -match '^AnalysisServices') { return @('Analysis Services', 'External') }
    if ($t -match '^https?://') { return @('Web', 'External') }
    if ($t -match '^(\\\\|[a-z]:\\)') { return @('File', 'External') }
    return @('Other', 'External')
}

function Get-A2CEditDistance {
    # Levenshtein distance of two short texts (workspace names), case-insensitive.
    param([string] $A, [string] $B)
    $a1 = $A.ToLowerInvariant(); $b1 = $B.ToLowerInvariant()
    $prev = New-Object int[] ($b1.Length + 1)
    for ($j = 0; $j -le $b1.Length; $j++) { $prev[$j] = $j }
    for ($i = 1; $i -le $a1.Length; $i++) {
        $curr = New-Object int[] ($b1.Length + 1)
        $curr[0] = $i
        for ($j = 1; $j -le $b1.Length; $j++) {
            $cost = $(if ($a1[$i - 1] -eq $b1[$j - 1]) { 0 } else { 1 })
            $curr[$j] = [math]::Min([math]::Min($curr[$j - 1] + 1, $prev[$j] + 1), $prev[$j - 1] + $cost)
        }
        $prev = $curr
    }
    return $prev[$b1.Length]
}

function Read-A2CTable {
    # Rows of a .csv file or of a sheet of an .xlsx workbook (the sheet named -Sheet, otherwise the first one), as
    # ordered dictionaries keyed by the header row. No Excel and no module needed.
    param([string] $Path, [string] $Sheet = '')
    $rows = [System.Collections.Generic.List[object]]::new()
    $grid = [System.Collections.Generic.List[object]]::new()
    if ($Path -like '*.xlsx') {
        Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
        $zip = [IO.Compression.ZipFile]::OpenRead($Path)
        try {
            $read = { param($n) $e = $zip.GetEntry($n); if ($null -eq $e) { return $null }; $sr = New-Object IO.StreamReader($e.Open()); try { [xml]$sr.ReadToEnd() } finally { $sr.Dispose() } }
            $strings = [System.Collections.Generic.List[string]]::new()
            $ss = & $read 'xl/sharedStrings.xml'
            if ($null -ne $ss) { foreach ($si in $ss.sst.ChildNodes) { $strings.Add($si.InnerText) } }
            $wb = & $read 'xl/workbook.xml'
            $rels = & $read 'xl/_rels/workbook.xml.rels'
            $pick = $null
            foreach ($s in $wb.workbook.sheets.ChildNodes) { if ($null -eq $pick -or ($Sheet -and $s.name -eq $Sheet)) { $pick = $s } }
            $rid = $pick.GetAttribute('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
            $target = ''
            foreach ($r in $rels.Relationships.ChildNodes) { if ($r.Id -eq $rid) { $target = [string]$r.Target } }
            $doc = & $read ('xl/' + $target.TrimStart('/').Replace('xl/', ''))
            foreach ($r in $doc.GetElementsByTagName('row')) {
                $cells = @{}
                $max = -1
                foreach ($c in $r.ChildNodes) {
                    if ($c.LocalName -ne 'c') { continue }
                    $col = 0
                    foreach ($ch in ([string]$c.GetAttribute('r')).ToCharArray()) { if ($ch -ge 'A' -and $ch -le 'Z') { $col = $col * 26 + ([int]$ch - 64) } }
                    $v = ''
                    $vn = $c.GetElementsByTagName('v')
                    if ($c.GetAttribute('t') -eq 's' -and $vn.Count -gt 0) { $v = $strings[[int]$vn[0].InnerText] }
                    elseif ($c.GetAttribute('t') -eq 'inlineStr') { $v = $c.InnerText }
                    elseif ($vn.Count -gt 0) { $v = $vn[0].InnerText }
                    $cells[$col - 1] = [string]$v
                    if ($col - 1 -gt $max) { $max = $col - 1 }
                }
                $line = New-Object string[] ($max + 1)
                for ($k = 0; $k -le $max; $k++) { $line[$k] = [string]$cells[$k] }
                $grid.Add($line)
            }
        }
        finally { $zip.Dispose() }
    }
    else {
        $first = (Get-Content -LiteralPath $Path -TotalCount 1)
        $sep = ','
        if (([string]$first).Split(';').Count -gt ([string]$first).Split(',').Count) { $sep = ';' }
        foreach ($o in @(Import-Csv -LiteralPath $Path -Delimiter $sep -Encoding UTF8)) {
            $row = [ordered]@{}
            foreach ($p in $o.PSObject.Properties) { $row[$p.Name.Trim()] = ([string]$p.Value).Trim() }
            $rows.Add($row)
        }
        return , $rows
    }
    if ($grid.Count -eq 0) { return , $rows }
    $head = @($grid[0] | ForEach-Object { ([string]$_).Trim() })
    for ($i = 1; $i -lt $grid.Count; $i++) {
        $row = [ordered]@{}
        $any = $false
        for ($k = 0; $k -lt $head.Count; $k++) {
            if (-not $head[$k]) { continue }
            $v = ''
            if ($k -lt $grid[$i].Count) { $v = ([string]$grid[$i][$k]).Trim() }
            if ($v) { $any = $true }
            $row[$head[$k]] = $v
        }
        if ($any) { $rows.Add($row) }
    }
    return , $rows
}

function New-A2CPbiRow {
    # A report row that starts with the workspace columns.
    param($Ws, [System.Collections.IDictionary] $Values)
    $row = [ordered]@{ Workspace = $Ws.Name; WorkspaceId = $Ws.Id }
    foreach ($k in $Values.Keys) { $row[$k] = $Values[$k] }
    return $row
}

function Get-A2CPbiDatasourceValues {
    # Columns of $A2CPbiDsCols for a data source of a model, a paginated report or a dataflow.
    param($Ds)
    $c = $Ds.connectionDetails
    $gw = ''
    if ($Ds.gatewayId) { $gw = [string]$script:A2C.Data['PbiGatewayNames'][[string]$Ds.gatewayId] }
    return [ordered]@{ DatasourceType = [string]$Ds.datasourceType; Name = [string]$Ds.name; Server = [string]$c.server; Database = [string]$c.database
        Url = [string]$c.url; Path = [string]$c.path; ConnectionString = [string]$Ds.connectionString; GatewayId = [string]$Ds.gatewayId; GatewayName = $gw
        DatasourceId = [string]$Ds.datasourceId; ConnectionDetails = $c }
}

function Get-A2CPbiExportedTo {
    param([string] $ItemId)
    $m = $script:A2C.Data['PbiExportMap']
    if ($null -eq $m -or -not $ItemId) { return '' }
    return [string]$m[$ItemId.ToLowerInvariant()]
}

function Get-A2CPbiCapacityMap {
    # Capacity id -> @{ Name; Sku; Region; State; ... } from the Power BI and the Fabric capacity lists.
    $map = @{}
    $t = $script:A2C.Data['PbiTenant']
    if ($null -ne $t) {
        foreach ($c in $t.Capacities) {
            $map[([string]$c.id).ToLowerInvariant()] = [ordered]@{ Capacity = [string]$c.displayName; CapacityId = [string]$c.id; Sku = [string]$c.sku; Region = [string]$c.region
                State = [string]$c.state; UserAccessRight = [string]$c.capacityUserAccessRight; Admins = (Join-A2C $c.admins); TenantKeyId = [string]$c.tenantKeyId; Source = 'Power BI' }
        }
    }
    $f = $script:A2C.Data['FabricTenant']
    if ($null -ne $f) {
        foreach ($c in $f.Capacities) {
            $k = ([string]$c.id).ToLowerInvariant()
            if ($map.ContainsKey($k)) { $map[$k]['Source'] = 'Power BI; Fabric'; continue }
            $map[$k] = [ordered]@{ Capacity = [string]$c.displayName; CapacityId = [string]$c.id; Sku = [string]$c.sku; Region = [string]$c.region
                State = [string]$c.state; UserAccessRight = ''; Admins = ''; TenantKeyId = ''; Source = 'Fabric' }
        }
    }
    return $map
}

function New-A2CPbiSharedReports {
    # Files that both the PowerBI and the Fabric report feed: workspaces, capacities, gateways. Built once per run,
    # from whatever was collected.
    if ($script:A2C.Out.ContainsKey('PowerBI/workspaces.csv')) { return }
    $caps = Get-A2CPbiCapacityMap
    $perCap = @{}
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($w in @($script:A2C.Data['PbiWorkspaces'])) {
        $cap = $null
        if ($w.CapacityId) { $k = $w.CapacityId.ToLowerInvariant(); $cap = $caps[$k]; $perCap[$k] = [int]$perCap[$k] + 1 }
        $admins = @($w.Users | Where-Object { $_.groupUserAccessRight -eq 'Admin' } | ForEach-Object { $(if ($_.emailAddress) { $_.emailAddress } else { $_.displayName }) })
        if ($admins.Count -eq 0) {
            $admins = @($w.RoleAssignments | Where-Object { $_.role -eq 'Admin' } | ForEach-Object { [string](Get-A2CFirst $_.principal @('userDetails.userPrincipalName', 'displayName', 'id')) })
        }
        $users = @($w.Users).Count
        if ($users -eq 0) { $users = @($w.RoleAssignments).Count }
        $rows.Add((New-A2CPbiRow $w ([ordered]@{
                        Type = $w.Type; Description = $w.Description; CapacityId = $w.CapacityId
                        Capacity = $(if ($null -ne $cap) { $cap['Capacity'] } else { '' }); CapacitySku = $(if ($null -ne $cap) { $cap['Sku'] } else { '' })
                        CapacityRegion = $(if ($null -ne $cap) { $cap['Region'] } else { [string](Get-A2CFirst $w.Detail @('capacityRegion')) })
                        IsOnDedicatedCapacity = $(if ($null -ne $w.IsOnDedicatedCapacity) { $w.IsOnDedicatedCapacity } else { [bool]$w.CapacityId })
                        IsReadOnly = $w.IsReadOnly; DomainId = $w.DomainId; DefaultStorageFormat = $w.DefaultDatasetStorageFormat
                        InPowerBI = $w.InPbi; InFabric = $w.InFabric; Reports = @($w.Reports).Count; SemanticModels = @($w.Datasets).Count
                        Dashboards = @($w.Dashboards).Count; Dataflows = @($w.Dataflows).Count; FabricItems = @($w.Items).Count; Users = $users; Admins = (Join-A2C $admins)
                        GitState = [string](Get-A2CFirst $w.Git @('gitConnectionState'))
                        OneLakeBlobEndpoint = [string](Get-A2CFirst $w.Detail @('oneLakeEndpoints.blobEndpoint')); OneLakeDfsEndpoint = [string](Get-A2CFirst $w.Detail @('oneLakeEndpoints.dfsEndpoint'))
                        WorkspaceIdentity = [string](Get-A2CFirst $w.Detail @('workspaceIdentity.servicePrincipalId'))
                        PowerBIStatus = $(if ($w.Status -eq 'Collected' -and $w.Problems.Count -gt 0) { 'Partial' } else { $w.Status }); PowerBIProblems = (Get-A2CPbiProblemText $w.Problems)
                        FabricStatus = $(if ($w.FabricStatus -eq 'Collected' -and $w.FabricProblems.Count -gt 0) { 'Partial' } else { $w.FabricStatus }); FabricProblems = (Get-A2CPbiProblemText $w.FabricProblems)
                    })))
    }
    Add-A2CRows 'PowerBI/workspaces.csv' $rows

    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($k in @($caps.Keys | Sort-Object)) { $c = $caps[$k]; $c['Workspaces'] = [int]$perCap[$k]; $rows.Add($c) }
    Add-A2CRows 'PowerBI/capacities.csv' $rows

    $rows = [System.Collections.Generic.List[object]]::new()
    $seen = @{}
    $t = $script:A2C.Data['PbiTenant']
    if ($null -ne $t) {
        foreach ($g in $t.Gateways) {
            $seen[([string]$g.id).ToLowerInvariant()] = $true
            $rows.Add([ordered]@{ Gateway = [string]$g.name; GatewayId = [string]$g.id; Type = [string]$g.type; Status = [string]$g.gatewayStatus; Version = ''; Members = $null
                    Datasources = @(Get-A2CPbiSub $t 'gatewayDatasources' ([string]$g.id)).Count; Annotation = [string]$g.gatewayAnnotation; Source = 'Power BI'; Properties = $null })
        }
    }
    $f = $script:A2C.Data['FabricTenant']
    if ($null -ne $f) {
        foreach ($g in $f.Gateways) {
            if ($seen.ContainsKey(([string]$g.id).ToLowerInvariant())) { continue }
            $rows.Add([ordered]@{ Gateway = [string]$g.displayName; GatewayId = [string]$g.id; Type = [string]$g.type; Status = ''; Version = [string]$g.version
                    Members = $g.numberOfMemberGateways; Datasources = $null; Annotation = ''; Source = 'Fabric'; Properties = $g })
        }
    }
    Add-A2CRows 'PowerBI/gateways.csv' $rows
    New-A2CPbiWorkspacePlan
}

function New-A2CReportPowerBI {
    $wss = @($script:A2C.Data['PbiWorkspaces'])
    $t = $script:A2C.Data['PbiTenant']
    $gwNames = @{}
    if ($null -ne $t) { foreach ($g in $t.Gateways) { $gwNames[[string]$g.id] = [string]$g.name } }
    $script:A2C.Data['PbiGatewayNames'] = $gwNames
    New-A2CPbiSharedReports

    $files = @('workspaceUsers', 'reports', 'reportPages', 'reportDatasources', 'semanticModels', 'semanticModelDatasources', 'semanticModelUsers', 'semanticModelParameters',
        'refreshSchedules', 'refreshHistory', 'dashboards', 'dashboardTiles', 'dataflows', 'dataflowDatasources', 'dataflowTransactions', 'lineage', 'apps', 'appContent',
        'capacityWorkloads', 'refreshables', 'gatewayDatasources', 'gatewayDatasourceUsers', 'deploymentPipelines', 'pipelineStages', 'pipelineStageItems', 'pipelineUsers',
        'pipelineOperations', 'availableFeatures')
    $rows = @{}
    foreach ($f in $files) { $rows[$f] = [System.Collections.Generic.List[object]]::new() }
    $wsNames = @{}
    foreach ($w in $wss) { $wsNames[$w.Id.ToLowerInvariant()] = $w.Name }

    foreach ($w in $wss) {
        foreach ($u in $w.Users) {
            $rows['workspaceUsers'].Add((New-A2CPbiRow $w ([ordered]@{ DisplayName = [string]$u.displayName; Identifier = [string]$u.identifier; EmailAddress = [string]$u.emailAddress
                            PrincipalType = [string]$u.principalType; AccessRight = [string]$u.groupUserAccessRight; UserType = [string]$u.userType; GraphId = [string]$u.graphId })))
        }

        $modelNames = @{}
        $modelReports = @{}
        foreach ($d in $w.Datasets) { $modelNames[[string]$d.id] = [string]$d.name }
        foreach ($x in $w.Reports) {
            $id = [string]$x.id
            $mid = [string]$x.datasetId
            if ($mid) { $modelReports[$mid] = [int]$modelReports[$mid] + 1 }
            $pages = @(Get-A2CPbiSub $w 'reportPages' $id)
            $rows['reports'].Add((New-A2CPbiRow $w ([ordered]@{ Report = [string]$x.name; ReportId = $id; ReportType = [string]$x.reportType; Format = [string]$x.format; ModelId = $mid
                            Model = [string]$modelNames[$mid]; ModelWorkspaceId = [string]$x.datasetWorkspaceId; Description = [string]$x.description; IsOwnedByMe = $x.isOwnedByMe
                            IsFromPbix = $x.isFromPbix; AppId = [string]$x.appId; Pages = $pages.Count; WebUrl = [string]$x.webUrl; ExportedTo = (Get-A2CPbiExportedTo $id) })))
            foreach ($pg in $pages) {
                $rows['reportPages'].Add((New-A2CPbiRow $w ([ordered]@{ Report = [string]$x.name; ReportId = $id; Page = [string]$pg.name; DisplayName = [string]$pg.displayName; Order = $pg.order })))
            }
            foreach ($ds in @(Get-A2CPbiSub $w 'reportDatasources' $id)) {
                $v = [ordered]@{ Report = [string]$x.name; ReportId = $id }
                $dv = Get-A2CPbiDatasourceValues $ds
                foreach ($k in $dv.Keys) { $v[$k] = $dv[$k] }
                $rows['reportDatasources'].Add((New-A2CPbiRow $w $v))
            }
            if ($mid) {
                $rows['lineage'].Add((New-A2CPbiRow $w ([ordered]@{ ItemType = 'Report'; Item = [string]$x.name; ItemId = $id; DependsOnType = 'SemanticModel'; DependsOn = [string]$modelNames[$mid]
                                DependsOnId = $mid; DependsOnWorkspaceId = $(if ($x.datasetWorkspaceId) { [string]$x.datasetWorkspaceId } else { $w.Id }) })))
            }
        }

        foreach ($d in $w.Datasets) {
            $id = [string]$d.id
            $name = [string]$d.name
            $sources = @(Get-A2CPbiSub $w 'datasetDatasources' $id)
            $history = @(Get-A2CPbiSub $w 'datasetRefreshes' $id)
            $last = $null
            if ($history.Count -gt 0) { $last = $history[0] }
            $sched = Get-A2CPbiSub $w 'datasetSchedule' $id
            $dq = Get-A2CPbiSub $w 'datasetDqSchedule' $id
            $schedEnabled = $null
            # from the model definition, when the Dependencies report read it: storage modes of the partitions, measures, roles
            $modes = ''; $measures = $null; $roles = $null
            $objs = Get-A2CPbiModelObjects $id
            if ($objs.Count -gt 0) {
                $mc = @{}
                foreach ($o in $objs) { if ($o.Kind -eq 'partition') { $mk = [string]$o.Props['mode']; if (-not $mk) { $mk = [string]$o.Value }; $mc[$mk] = [int]$mc[$mk] + 1 } }
                $modes = Join-A2C @($mc.Keys | Sort-Object | ForEach-Object { '{0} ({1})' -f $_, $mc[$_] })
                $measures = @($objs | Where-Object { $_.Kind -eq 'measure' }).Count
                $roles = @($objs | Where-Object { $_.Kind -eq 'role' }).Count
            }
            if ($sched) { $schedEnabled = $sched.enabled }
            $rows['semanticModels'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; ConfiguredBy = [string]$d.configuredBy; CreatedDate = $d.createdDate
                            IsRefreshable = $d.isRefreshable; TargetStorageMode = [string]$d.targetStorageMode; IsEffectiveIdentityRequired = $d.isEffectiveIdentityRequired
                            IsEffectiveIdentityRolesRequired = $d.isEffectiveIdentityRolesRequired; IsOnPremGatewayRequired = $d.isOnPremGatewayRequired
                            AddRowsApiEnabled = $d.addRowsAPIEnabled; QueryScaleOut = $d.queryScaleOutSettings; Description = [string]$d.description
                            Reports = [int]$modelReports[$id]; Datasources = $sources.Count; DatasourceTypes = (Join-A2C (Get-A2CUnique @($sources | ForEach-Object { $_.datasourceType })))
                            StorageModes = $modes; Measures = $measures; RlsRoles = $roles
                            ScheduleEnabled = $schedEnabled; LastRefreshStatus = $(if ($null -ne $last) { [string]$last.status } else { '' })
                            LastRefreshEnd = $(if ($null -ne $last) { $last.endTime } else { $null }); WebUrl = [string]$d.webUrl; ExportedTo = (Get-A2CPbiExportedTo $id) })))
            foreach ($ds in $sources) {
                $v = [ordered]@{ Model = $name; ModelId = $id }
                $dv = Get-A2CPbiDatasourceValues $ds
                foreach ($k in $dv.Keys) { $v[$k] = $dv[$k] }
                $rows['semanticModelDatasources'].Add((New-A2CPbiRow $w $v))
            }
            foreach ($u in @(Get-A2CPbiSub $w 'datasetUsers' $id)) {
                $rows['semanticModelUsers'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Identifier = [string]$u.identifier; PrincipalType = [string]$u.principalType
                                AccessRight = [string]$u.datasetUserAccessRight })))
            }
            foreach ($pm in @(Get-A2CPbiSub $w 'datasetParameters' $id)) {
                $rows['semanticModelParameters'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Parameter = [string]$pm.name; Type = [string]$pm.type; IsRequired = $pm.isRequired
                                CurrentValue = [string]$pm.currentValue; SuggestedValues = (Join-A2C $pm.suggestedValues) })))
            }
            if ($sched) {
                $rows['refreshSchedules'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Kind = 'Import'; Enabled = $sched.enabled; Days = (Join-A2C $sched.days ', ')
                                Times = (Join-A2C $sched.times ', '); Frequency = $null; LocalTimeZoneId = [string]$sched.localTimeZoneId; NotifyOption = [string]$sched.notifyOption })))
            }
            if ($dq) {
                $rows['refreshSchedules'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Kind = 'DirectQuery cache'; Enabled = $null; Days = (Join-A2C $dq.days ', ')
                                Times = (Join-A2C $dq.times ', '); Frequency = $dq.frequency; LocalTimeZoneId = [string]$dq.localTimeZoneId; NotifyOption = '' })))
            }
            foreach ($h in $history) {
                $dur = $null
                $s = ConvertTo-A2CUtc $h.startTime
                $e = ConvertTo-A2CUtc $h.endTime
                if ($null -ne $s -and $null -ne $e) { $dur = [math]::Round(($e - $s).TotalSeconds, 0) }
                $rows['refreshHistory'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; RequestId = [string]$h.requestId; RefreshType = [string]$h.refreshType
                                Status = [string]$h.status; StartTime = $h.startTime; EndTime = $h.endTime; DurationSec = $dur; Attempts = @($h.refreshAttempts | Where-Object { $null -ne $_ }).Count
                                Error = [string]$h.serviceExceptionJson })))
            }
        }
        foreach ($l in $w.UpstreamDataflows) {
            $mid = [string]$l.datasetObjectId
            $rows['lineage'].Add((New-A2CPbiRow $w ([ordered]@{ ItemType = 'SemanticModel'; Item = [string]$modelNames[$mid]; ItemId = $mid; DependsOnType = 'Dataflow'; DependsOn = ''
                            DependsOnId = [string]$l.dataflowObjectId; DependsOnWorkspaceId = [string]$l.workspaceObjectId })))
        }

        foreach ($x in $w.Dashboards) {
            $id = [string]$x.id
            $tiles = @(Get-A2CPbiSub $w 'dashboardTiles' $id)
            $rows['dashboards'].Add((New-A2CPbiRow $w ([ordered]@{ Dashboard = [string]$x.displayName; DashboardId = $id; IsReadOnly = $x.isReadOnly; AppId = [string]$x.appId
                            Tiles = $tiles.Count; WebUrl = [string]$x.webUrl })))
            foreach ($tl in $tiles) {
                $rows['dashboardTiles'].Add((New-A2CPbiRow $w ([ordered]@{ Dashboard = [string]$x.displayName; DashboardId = $id; Tile = [string]$tl.title; TileId = [string]$tl.id
                                ReportId = [string]$tl.reportId; ModelId = [string]$tl.datasetId; RowSpan = $tl.rowSpan; ColSpan = $tl.colSpan })))
            }
        }

        foreach ($x in $w.Dataflows) {
            $id = [string]$x.objectId
            $name = [string]$x.name
            $sources = @(Get-A2CPbiSub $w 'dataflowDatasources' $id)
            $tx = @(Get-A2CPbiSub $w 'dataflowTransactions' $id)
            $last = $null
            if ($tx.Count -gt 0) { $last = $tx[0] }
            $rows['dataflows'].Add((New-A2CPbiRow $w ([ordered]@{ Dataflow = $name; DataflowId = $id; Description = [string]$x.description; ConfiguredBy = [string]$x.configuredBy
                            ModifiedBy = [string]$x.modifiedBy; ModifiedDateTime = $x.modifiedDateTime; Generation = $x.generation; Datasources = $sources.Count
                            LastRunStatus = $(if ($null -ne $last) { [string]$last.status } else { '' }); LastRunEnd = $(if ($null -ne $last) { $last.endTime } else { $null })
                            ExportedTo = (Get-A2CPbiExportedTo $id) })))
            foreach ($ds in $sources) {
                $v = [ordered]@{ Dataflow = $name; DataflowId = $id }
                $dv = Get-A2CPbiDatasourceValues $ds
                foreach ($k in $dv.Keys) { $v[$k] = $dv[$k] }
                $rows['dataflowDatasources'].Add((New-A2CPbiRow $w $v))
            }
            foreach ($tr in $tx) {
                $rows['dataflowTransactions'].Add((New-A2CPbiRow $w ([ordered]@{ Dataflow = $name; DataflowId = $id; TransactionId = [string]$tr.id; RefreshType = [string]$tr.refreshType
                                Status = [string]$tr.status; StartTime = $tr.startTime; EndTime = $tr.endTime })))
            }
            foreach ($l in @(Get-A2CPbiSub $w 'dataflowUpstream' $id)) {
                $rows['lineage'].Add((New-A2CPbiRow $w ([ordered]@{ ItemType = 'Dataflow'; Item = $name; ItemId = $id; DependsOnType = 'Dataflow'; DependsOn = ''
                                DependsOnId = [string]$l.targetDataflowId; DependsOnWorkspaceId = [string]$l.groupId })))
            }
        }
    }

    if ($null -ne $t) {
        foreach ($a in $t.Apps) {
            $id = [string]$a.id
            $ar = @(Get-A2CPbiSub $t 'appReports' $id)
            $ad = @(Get-A2CPbiSub $t 'appDashboards' $id)
            $rows['apps'].Add([ordered]@{ App = [string]$a.name; AppId = $id; Description = [string]$a.description; PublishedBy = [string]$a.publishedBy; LastUpdate = $a.lastUpdate
                    WorkspaceId = [string]$a.workspaceId; Workspace = [string]$wsNames[([string]$a.workspaceId).ToLowerInvariant()]; Reports = $ar.Count; Dashboards = $ad.Count })
            foreach ($x in $ar) {
                $rows['appContent'].Add([ordered]@{ App = [string]$a.name; AppId = $id; ContentType = 'Report'; Name = [string]$x.name; Id = [string]$x.id; ModelId = [string]$x.datasetId; WebUrl = [string]$x.webUrl })
            }
            foreach ($x in $ad) {
                $rows['appContent'].Add([ordered]@{ App = [string]$a.name; AppId = $id; ContentType = 'Dashboard'; Name = [string]$x.displayName; Id = [string]$x.id; ModelId = ''; WebUrl = [string]$x.webUrl })
            }
        }
        foreach ($c in $t.Capacities) {
            foreach ($wl in @(Get-A2CPbiSub $t 'workloads' ([string]$c.id))) {
                $rows['capacityWorkloads'].Add([ordered]@{ Capacity = [string]$c.displayName; CapacityId = [string]$c.id; Workload = [string]$wl.name; State = [string]$wl.state
                        MaxMemoryPercentageSetByUser = $wl.maxMemoryPercentageSetByUser })
            }
        }
        foreach ($x in $t.Refreshables) {
            $rows['refreshables'].Add([ordered]@{ Capacity = [string]$x.capacity.displayName; CapacityId = [string]$x.capacity.id; Workspace = [string]$x.group.name; WorkspaceId = [string]$x.group.id
                    Name = [string]$x.name; Id = [string]$x.id; Kind = [string]$x.kind; RefreshCount = $x.refreshCount; RefreshFailures = $x.refreshFailures
                    AverageDurationSec = $x.averageDuration; MedianDurationSec = $x.medianDuration; RefreshesPerDay = $x.refreshesPerDay
                    LastRefreshStatus = [string]$x.lastRefresh.status; LastRefreshStart = $x.lastRefresh.startTime; LastRefreshEnd = $x.lastRefresh.endTime
                    ScheduleEnabled = $x.refreshSchedule.enabled; ScheduleDays = (Join-A2C $x.refreshSchedule.days ', '); ScheduleTimes = (Join-A2C $x.refreshSchedule.times ', ') })
        }
        foreach ($g in $t.Gateways) {
            foreach ($ds in @(Get-A2CPbiSub $t 'gatewayDatasources' ([string]$g.id))) {
                $rows['gatewayDatasources'].Add([ordered]@{ Gateway = [string]$g.name; GatewayId = [string]$g.id; Datasource = [string]$ds.datasourceName; DatasourceId = [string]$ds.id
                        DatasourceType = [string]$ds.datasourceType; CredentialType = [string]$ds.credentialType; ConnectionDetails = [string]$ds.connectionDetails })
                foreach ($u in @(Get-A2CPbiSub $t 'gatewayDatasourceUsers' ([string]$ds.id))) {
                    $rows['gatewayDatasourceUsers'].Add([ordered]@{ Gateway = [string]$g.name; GatewayId = [string]$g.id; Datasource = [string]$ds.datasourceName; DatasourceId = [string]$ds.id
                            DisplayName = [string]$u.displayName; Identifier = [string]$u.identifier; EmailAddress = [string]$u.emailAddress; PrincipalType = [string]$u.principalType
                            AccessRight = [string]$u.datasourceAccessRight })
                }
            }
        }
        foreach ($pl in $t.Pipelines) {
            $id = [string]$pl.id
            $name = [string]$pl.displayName
            $stages = @(Get-A2CPbiSub $t 'pipelineStages' $id)
            $rows['deploymentPipelines'].Add([ordered]@{ Pipeline = $name; PipelineId = $id; Description = [string]$pl.description; Stages = $stages.Count
                    Workspaces = (Join-A2C @($stages | ForEach-Object { $_.workspaceName })) })
            foreach ($s in $stages) {
                $rows['pipelineStages'].Add([ordered]@{ Pipeline = $name; PipelineId = $id; Order = $s.order; WorkspaceId = [string]$s.workspaceId; Workspace = [string]$s.workspaceName })
                $arts = Get-A2CPbiSub $t 'pipelineArtifacts' ($id + '|' + [string]$s.order)
                if (-not $arts) { continue }
                foreach ($prop in $arts.PSObject.Properties) {
                    if ($prop.Name -like '@*') { continue }
                    foreach ($it in @($prop.Value)) {
                        if ($null -eq $it) { continue }
                        $rows['pipelineStageItems'].Add([ordered]@{ Pipeline = $name; PipelineId = $id; Order = $s.order; ItemType = $prop.Name; Item = [string]$it.artifactDisplayName
                                ItemId = [string]$it.artifactId; SourceItemId = [string]$it.sourceArtifactId; TargetItemId = [string]$it.targetArtifactId; LastDeploymentTime = $it.lastDeploymentTime })
                    }
                }
            }
            foreach ($u in @(Get-A2CPbiSub $t 'pipelineUsers' $id)) {
                $rows['pipelineUsers'].Add([ordered]@{ Pipeline = $name; PipelineId = $id; Identifier = [string]$u.identifier; PrincipalType = [string]$u.principalType; AccessRight = [string]$u.accessRight })
            }
            foreach ($op in @(Get-A2CPbiSub $t 'pipelineOperations' $id)) {
                $rows['pipelineOperations'].Add([ordered]@{ Pipeline = $name; PipelineId = $id; OperationId = [string]$op.id; Type = [string]$op.type; Status = [string]$op.status
                        ExecutionStartTime = $op.executionStartTime; ExecutionEndTime = $op.executionEndTime; SourceStageOrder = $op.sourceStageOrder
                        TargetStageOrder = $op.targetStageOrder; PerformedBy = [string](Get-A2CFirst $op @('performedBy.userPrincipalName', 'performedBy.principalObjectId'))
                        Note = [string](Get-A2CFirst $op @('note.content')) })
            }
        }
        foreach ($x in $t.Features) {
            $rows['availableFeatures'].Add([ordered]@{ Feature = [string]$x.name; State = [string]$x.state; ExtendedState = [string]$x.extendedState; AdditionalInfo = $x.additionalInfo })
        }
    }
    foreach ($f in $files) { Add-A2CRows ('PowerBI/' + $f + '.csv') $rows[$f] }
    Write-A2CLog -Level Ok ('Power BI: {0} reports, {1} semantic models, {2} dashboards, {3} dataflows -> PowerBI folder' -f $rows['reports'].Count, $rows['semanticModels'].Count,
        $rows['dashboards'].Count, $rows['dataflows'].Count)
}

function New-A2CReportFabric {
    $wss = @($script:A2C.Data['PbiWorkspaces'])
    Initialize-A2CPbiItemIndex
    New-A2CPbiSharedReports
    New-A2CPbiEnvironmentMatrix
    $files = @('fabricItems', 'workspaceRoleAssignments', 'workspaceGit', 'workspaceSparkSettings', 'sparkPools', 'managedPrivateEndpoints', 'lakehouses', 'lakehouseTables',
        'warehouses', 'shortcuts', 'itemSchedules', 'itemRuns', 'itemConnections', 'connections', 'folders', 'workspaceNetwork', 'gitItemStatus', 'sparkSessions',
        'onelakeRoles', 'deploymentStages', 'deploymentStageItems', 'deploymentRoleAssignments', 'domains')
    $rows = @{}
    foreach ($f in $files) { $rows[$f] = [System.Collections.Generic.List[object]]::new() }

    foreach ($w in $wss) {
        $props = @{}   # item id -> type specific properties (lakehouse, warehouse, and the types of $A2CPbiTypedLists)
        if ($w.Sub.ContainsKey('typedProps')) { foreach ($k in $w.Sub['typedProps'].Keys) { $props[$k] = $w.Sub['typedProps'][$k] } }
        foreach ($x in $w.Lakehouses) { $props[[string]$x.id] = $x.properties }
        foreach ($x in $w.Warehouses) { $props[[string]$x.id] = $x.properties }
        # folders: id -> path from the workspace root
        $folders = @{}
        foreach ($f in @(Get-A2CPbiSub $w 'folders' 'ws')) { $folders[[string]$f.id] = $f }
        $folderPath = @{}
        foreach ($fid in @($folders.Keys)) {
            $parts = [System.Collections.Generic.List[string]]::new()
            $cur = $folders[$fid]
            $guard = 0
            while ($null -ne $cur -and $guard -lt 50) { $parts.Insert(0, [string]$cur.displayName); $cur = $folders[[string]$cur.parentFolderId]; $guard++ }
            $folderPath[$fid] = [string]::Join('/', $parts)
        }
        $inFolder = @{}
        foreach ($x in $w.Items) { if ($x.folderId) { $inFolder[[string]$x.folderId] = [int]$inFolder[[string]$x.folderId] + 1 } }
        foreach ($fid in @($folders.Keys | Sort-Object { $folderPath[$_] })) {
            $rows['folders'].Add((New-A2CPbiRow $w ([ordered]@{ Folder = [string]$folders[$fid].displayName; FolderId = $fid; ParentFolderId = [string]$folders[$fid].parentFolderId
                            Path = $folderPath[$fid]; Items = [int]$inFolder[$fid] })))
        }
        $net = Get-A2CPbiSub $w 'network' 'ws'
        if ($net) {
            $rows['workspaceNetwork'].Add((New-A2CPbiRow $w ([ordered]@{ InboundDefaultAction = [string](Get-A2CFirst $net @('inbound.publicAccessRules.defaultAction'))
                            OutboundDefaultAction = [string](Get-A2CFirst $net @('outbound.publicAccessRules.defaultAction')); Policy = $net })))
        }
        foreach ($ss in @(Get-A2CPbiSub $w 'sparkSessions' 'ws')) {
            $ref = Get-A2CPbiItemRef ([string]$ss.item.itemId)
            $secs = $null
            if ($null -ne $ss.totalDuration -and $null -ne $ss.totalDuration.value) {
                $secs = [double]$ss.totalDuration.value
                switch -Wildcard ([string]$ss.totalDuration.timeUnit) { 'Min*' { $secs = $secs * 60 } 'Hour*' { $secs = $secs * 3600 } 'Milli*' { $secs = $secs / 1000 } }
                $secs = [math]::Round($secs, 0)
            }
            $rows['sparkSessions'].Add((New-A2CPbiRow $w ([ordered]@{ ItemType = $(if ($null -ne $ref) { $ref.Type } else { [string]$ss.item.itemType }); Item = $(if ($null -ne $ref) { $ref.Name } else { [string]$ss.itemName })
                            ItemId = [string]$ss.item.itemId; LivyId = [string]$ss.livyId; SparkApplicationId = [string]$ss.sparkApplicationId; State = [string]$ss.state; Origin = [string]$ss.origin
                            JobType = [string]$ss.jobType; SubmittedDateTime = (ConvertTo-A2CUtc $ss.submittedDateTime); StartDateTime = (ConvertTo-A2CUtc $ss.startDateTime)
                            EndDateTime = (ConvertTo-A2CUtc $ss.endDateTime); TotalDurationSec = $secs; Submitter = [string]$ss.submitter.id; RuntimeVersion = [string]$ss.runtimeVersion; Properties = $ss })))
        }
        foreach ($x in $w.Lakehouses) {
            foreach ($role in @(Get-A2CPbiSub $w 'dataAccessRoles' ([string]$x.id))) {
                $paths = [System.Collections.Generic.List[string]]::new()
                $actions = [System.Collections.Generic.List[string]]::new()
                foreach ($rule in @($role.decisionRules)) {
                    foreach ($pm in @($rule.permission)) {
                        if ($null -eq $pm) { continue }
                        foreach ($v in @($pm.attributeValueIncludedIn)) { if ([string]$pm.attributeName -eq 'Path') { $paths.Add([string]$v) } else { $actions.Add([string]$v) } }
                    }
                }
                $rows['onelakeRoles'].Add((New-A2CPbiRow $w ([ordered]@{ Lakehouse = [string]$x.displayName; LakehouseId = [string]$x.id; Role = [string]$role.name; Paths = (Join-A2C (Get-A2CUnique $paths))
                                Actions = (Join-A2C (Get-A2CUnique $actions)); EntraMembers = (Join-A2C @($role.members.microsoftEntraMembers | ForEach-Object { $_.objectId }))
                                FabricItemMembers = (Join-A2C @($role.members.fabricItemMembers | ForEach-Object { ([string]$_.sourcePath) + ' [' + (Join-A2C $_.itemAccess ',') + ']' })); Definition = $role })))
            }
        }
        foreach ($x in @($w.Items | Sort-Object type, displayName)) {
            $id = [string]$x.id
            $type = [string]$x.type
            $name = [string]$x.displayName
            $conns = @(Get-A2CPbiSub $w 'itemConnections' $id)
            $shorts = @(Get-A2CPbiSub $w 'shortcuts' $id)
            $scheds = @(Get-A2CPbiSub $w 'itemSchedules' $id)
            $runs = @(Get-A2CPbiSub $w 'itemRuns' $id)
            $lastRun = $null
            if ($runs.Count -gt 0) { $lastRun = @($runs | Sort-Object { [string]$_.startTimeUtc } -Descending)[0] }
            $rows['fabricItems'].Add((New-A2CPbiRow $w ([ordered]@{ ItemType = $type; Item = $name; ItemId = $id; Description = [string]$x.description; FolderPath = [string]$folderPath[[string]$x.folderId]
                            FolderId = [string]$x.folderId; LogicalId = [string]$x.logicalId; Tags = (Join-A2C @($x.tags | ForEach-Object { $_.displayName }))
                            Connections = $conns.Count; Shortcuts = $shorts.Count; Schedules = $scheds.Count; Runs = $(if ($script:A2CPbiRunTypes -contains $type) { $runs.Count } else { $null })
                            LastRunStatus = [string]$lastRun.status; LastRunStart = (ConvertTo-A2CUtc $lastRun.startTimeUtc); Properties = $props[$id]; ExportedTo = (Get-A2CPbiExportedTo $id) })))
            foreach ($rn in $runs) {
                $dur = $null
                $rs = ConvertTo-A2CUtc $rn.startTimeUtc
                $re = ConvertTo-A2CUtc $rn.endTimeUtc
                if ($null -ne $rs -and $null -ne $re) { $dur = [math]::Round(($re - $rs).TotalSeconds, 0) }
                $rows['itemRuns'].Add((New-A2CPbiRow $w ([ordered]@{ ItemType = $type; Item = $name; ItemId = $id; RunId = [string]$rn.id; JobType = [string]$rn.jobType; InvokeType = [string]$rn.invokeType
                                Status = [string]$rn.status; StartTimeUtc = $rs; EndTimeUtc = $re; DurationSec = $dur; FailureReason = $rn.failureReason })))
            }
            foreach ($c in $conns) {
                $cd = $c.connectionDetails
                $rows['itemConnections'].Add((New-A2CPbiRow $w ([ordered]@{ ItemType = $type; Item = $name; ItemId = $id; Connection = [string]$c.displayName; ConnectionId = [string]$c.id
                                ConnectivityType = [string]$c.connectivityType; ConnectionType = [string]$cd.type; Path = [string]$cd.path; GatewayId = [string]$c.gatewayId })))
            }
            foreach ($s in $shorts) {
                $tt = [string]$s.target.type
                if (-not $tt -and $null -ne $s.target) { $tt = (Join-A2C @($s.target.PSObject.Properties | ForEach-Object { $_.Name }) ', ') }
                $st = Get-A2CPbiShortcutTarget $s
                $rows['shortcuts'].Add((New-A2CPbiRow $w ([ordered]@{ ItemType = $type; Item = $name; ItemId = $id; Shortcut = [string]$s.name; Path = [string]$s.path; TargetType = $tt
                                TargetConnection = $st.Connection; TargetConnectionId = $st.ConnectionId; TargetLocation = $st.Location; TargetSubpath = $st.Subpath
                                TargetWorkspace = $(if ($null -ne $st.Ref) { $st.Ref.Ws } else { '' }); TargetItem = $(if ($null -ne $st.Ref) { $st.Ref.Name } else { '' }); Target = $s.target })))
            }
            foreach ($s in $scheds) {
                $rows['itemSchedules'].Add((New-A2CPbiRow $w ([ordered]@{ ItemType = $type; Item = $name; ItemId = $id; JobType = [string]$script:A2CPbiJobTypes[$type]; ScheduleId = [string]$s.id
                                Enabled = $s.enabled; CreatedDateTime = $s.createdDateTime; Owner = [string](Get-A2CFirst $s.owner @('userDetails.userPrincipalName', 'displayName', 'id'))
                                Configuration = $s.configuration })))
            }
        }
        foreach ($ra in $w.RoleAssignments) {
            $pn = $ra.principal
            $rows['workspaceRoleAssignments'].Add((New-A2CPbiRow $w ([ordered]@{ Principal = [string]$pn.displayName; PrincipalId = [string]$pn.id; PrincipalType = [string]$pn.type; Role = [string]$ra.role
                            UserPrincipalName = [string](Get-A2CFirst $pn @('userDetails.userPrincipalName')); AppId = [string](Get-A2CFirst $pn @('servicePrincipalDetails.aadAppId'))
                            GroupType = [string](Get-A2CFirst $pn @('groupDetails.groupType')) })))
        }
        if ($null -ne $w.Git) {
            $gp = $w.Git.gitProviderDetails
            $gst = ''; $gUn = $null; $gIn = $null; $gCf = $null; $gHash = ''
            $gs = Get-A2CPbiSub $w 'gitStatus' 'ws'
            if ($gs) {
                if (-not $gs.Ok) { $gst = 'Not read: ' + $gs.Error }
                else {
                    $changes = @($gs.Json.changes | Where-Object { $null -ne $_ })
                    $gst = 'Read'; $gHash = [string]$gs.Json.remoteCommitHash
                    $gUn = @($changes | Where-Object { $_.workspaceChange }).Count; $gIn = @($changes | Where-Object { $_.remoteChange }).Count
                    $gCf = @($changes | Where-Object { $_.conflictType -and [string]$_.conflictType -ne 'None' }).Count
                    foreach ($c in $changes) {
                        $im = $c.itemMetadata
                        $rows['gitItemStatus'].Add((New-A2CPbiRow $w ([ordered]@{ ItemType = [string]$im.itemType; Item = [string]$im.displayName; ItemId = [string]$im.itemIdentifier.objectId
                                        LogicalId = [string]$im.itemIdentifier.logicalId; WorkspaceChange = [string]$c.workspaceChange; RemoteChange = [string]$c.remoteChange; ConflictType = [string]$c.conflictType })))
                    }
                }
            }
            $rows['workspaceGit'].Add((New-A2CPbiRow $w ([ordered]@{ State = [string]$w.Git.gitConnectionState; Provider = [string]$gp.gitProviderType
                            Organization = [string](Get-A2CFirst $gp @('organizationName', 'ownerName')); Project = [string]$gp.projectName; Repository = [string]$gp.repositoryName
                            Branch = [string]$gp.branchName; Directory = [string]$gp.directoryName; Head = [string]$w.Git.gitSyncDetails.head; LastSyncTime = $w.Git.gitSyncDetails.lastSyncTime
                            ItemStatus = $gst; UncommittedItems = $gUn; IncomingItems = $gIn; Conflicts = $gCf; RemoteCommitHash = $gHash })))
        }
        if ($null -ne $w.SparkSettings) {
            $ss = $w.SparkSettings
            $rows['workspaceSparkSettings'].Add((New-A2CPbiRow $w ([ordered]@{ DefaultPool = [string]$ss.pool.defaultPool.name; DefaultPoolType = [string]$ss.pool.defaultPool.type
                            RuntimeVersion = [string]$ss.environment.runtimeVersion; Environment = [string]$ss.environment.name; AutomaticLogEnabled = $ss.automaticLog.enabled
                            CustomizeComputeEnabled = $ss.pool.customizeComputeEnabled; Settings = $ss })))
        }
        foreach ($x in $w.SparkPools) {
            $rows['sparkPools'].Add((New-A2CPbiRow $w ([ordered]@{ Pool = [string]$x.name; PoolId = [string]$x.id; Type = [string]$x.type; NodeFamily = [string]$x.nodeFamily; NodeSize = [string]$x.nodeSize
                            AutoScale = $x.autoScale; DynamicExecutorAllocation = $x.dynamicExecutorAllocation })))
        }
        foreach ($x in $w.ManagedPrivateEndpoints) {
            $rows['managedPrivateEndpoints'].Add((New-A2CPbiRow $w ([ordered]@{ Name = [string]$x.name; Id = [string]$x.id; TargetPrivateLinkResourceId = [string]$x.targetPrivateLinkResourceId
                            TargetSubresourceType = [string]$x.targetSubresourceType; ProvisioningState = [string]$x.provisioningState; ConnectionState = $x.connectionState })))
        }
        foreach ($x in $w.Lakehouses) {
            $id = [string]$x.id
            $lp = $x.properties
            $tables = @(Get-A2CPbiSub $w 'lakehouseTables' $id)
            $listed = ($w.Sub.ContainsKey('lakehouseTables') -and $w.Sub['lakehouseTables'].ContainsKey($id))
            $sql = Get-A2CSqlSummary $id
            # the table list API does not work for lakehouses with schemas: there the SQL endpoint is the source
            $count = $null; $source = ''
            if ($listed) { $count = $tables.Count; $source = 'Fabric API' }
            elseif ($sql.Status -eq 'Read') { $count = $sql.Tables; $source = 'SQL endpoint' }
            $rows['lakehouses'].Add((New-A2CPbiRow $w ([ordered]@{ Lakehouse = [string]$x.displayName; LakehouseId = $id; OneLakeTablesPath = [string]$lp.oneLakeTablesPath
                            OneLakeFilesPath = [string]$lp.oneLakeFilesPath; SqlEndpointId = [string]$lp.sqlEndpointProperties.id
                            SqlEndpointConnectionString = [string]$lp.sqlEndpointProperties.connectionString; SqlEndpointStatus = [string]$lp.sqlEndpointProperties.provisioningStatus
                            DefaultSchema = [string]$lp.defaultSchema; Tables = $count; TablesSource = $source; Views = $sql.Views; StoredProcedures = $sql.StoredProcedures
                            Functions = $sql.Functions; Shortcuts = @(Get-A2CPbiSub $w 'shortcuts' $id).Count; SqlStatus = $sql.Status })))
            foreach ($tb in $tables) {
                $rows['lakehouseTables'].Add((New-A2CPbiRow $w ([ordered]@{ Lakehouse = [string]$x.displayName; LakehouseId = $id; Schema = ''; Table = [string]$tb.name; Type = [string]$tb.type
                                Format = [string]$tb.format; Location = [string]$tb.location; Source = 'Fabric API' })))
            }
            if (-not $listed) {
                foreach ($o in $sql.Objects) {
                    if ([string]$o['type'] -ne 'U') { continue }
                    $rows['lakehouseTables'].Add((New-A2CPbiRow $w ([ordered]@{ Lakehouse = [string]$x.displayName; LakehouseId = $id; Schema = [string]$o['schema']; Table = [string]$o['name']
                                    Type = ''; Format = ''; Location = ''; Source = 'SQL endpoint' })))
                }
            }
        }
        foreach ($x in $w.Warehouses) {
            $id = [string]$x.id
            $sql = Get-A2CSqlSummary $id
            $rows['warehouses'].Add((New-A2CPbiRow $w ([ordered]@{ Warehouse = [string]$x.displayName; WarehouseId = $id; Description = [string]$x.description
                            ConnectionString = [string]$x.properties.connectionString; CreatedDate = $x.properties.createdDate; LastUpdatedTime = $x.properties.lastUpdatedTime
                            Tables = $sql.Tables; Views = $sql.Views; StoredProcedures = $sql.StoredProcedures; Functions = $sql.Functions; SqlStatus = $sql.Status })))
        }
    }
    $f = $script:A2C.Data['FabricTenant']
    if ($null -ne $f) {
        foreach ($x in $f.Domains) {
            $rows['domains'].Add([ordered]@{ Domain = [string]$x.displayName; DomainId = [string]$x.id; Description = [string]$x.description; ParentDomainId = [string]$x.parentDomainId })
        }
        foreach ($dp in $f.DeploymentPipelines) {
            $dpid = [string]$dp.id
            $dpn = [string]$dp.displayName
            foreach ($st in @(Get-A2CPbiSub $f 'dpStages' $dpid)) {
                $its = @(Get-A2CPbiSub $f 'dpStageItems' ([string]$st.id))
                $rows['deploymentStages'].Add([ordered]@{ Pipeline = $dpn; PipelineId = $dpid; Stage = [string]$st.displayName; StageId = [string]$st.id; Order = $st.order
                        Workspace = [string]$st.workspaceName; WorkspaceId = [string]$st.workspaceId; IsPublic = $st.isPublic; Items = $(if ($st.workspaceId) { $its.Count } else { $null })
                        Description = [string]$st.description })
                foreach ($it in $its) {
                    $rows['deploymentStageItems'].Add([ordered]@{ Pipeline = $dpn; PipelineId = $dpid; Stage = [string]$st.displayName; Order = $st.order; Workspace = [string]$st.workspaceName
                            ItemType = [string]$it.itemType; Item = [string]$it.itemDisplayName; ItemId = [string]$it.itemId; SourceItemId = [string]$it.sourceItemId
                            TargetItemId = [string]$it.targetItemId; LastDeploymentTime = $it.lastDeploymentTime })
                }
            }
            foreach ($ra in @(Get-A2CPbiSub $f 'dpRoles' $dpid)) {
                $rows['deploymentRoleAssignments'].Add([ordered]@{ Pipeline = $dpn; PipelineId = $dpid; PrincipalId = [string]$ra.principal.id; PrincipalType = [string]$ra.principal.type; Role = [string]$ra.role })
            }
        }
        foreach ($c in $f.Connections) {
            $cd = $c.connectionDetails
            $cr = $c.credentialDetails
            $rows['connections'].Add([ordered]@{ Connection = [string]$c.displayName; ConnectionId = [string]$c.id; ConnectivityType = [string]$c.connectivityType; ConnectionType = [string]$cd.type
                    Path = [string]$cd.path; GatewayId = [string]$c.gatewayId; PrivacyLevel = [string]$c.privacyLevel; CredentialType = [string]$cr.credentialType
                    SingleSignOnType = [string]$cr.singleSignOnType; ConnectionEncryption = [string]$cr.connectionEncryption; SkipTestConnection = $cr.skipTestConnection })
        }
    }
    foreach ($f in $files) { Add-A2CRows ('PowerBI/' + $f + '.csv') $rows[$f] }
    Write-A2CLog -Level Ok ('Fabric: {0} items, {1} lakehouse table(s), {2} shortcut(s) -> PowerBI folder' -f $rows['fabricItems'].Count, $rows['lakehouseTables'].Count, $rows['shortcuts'].Count)
}

Register-A2CSchema 'PowerBI/sqlObjects.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid Schema:name Object:name ObjectType Created:date Modified:date Columns:num DefinitionLength:num')
Register-A2CSchema 'PowerBI/sqlColumns.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid Schema:name Table:name Column:name Ordinal:num DataType Length:num Precision:num Scale:num Nullable')
Register-A2CSchema 'PowerBI/sqlModules.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid Schema:name Object:name ObjectType DefinitionLength:num ReferencedDatabases:name Definition:text')
Register-A2CSchema 'PowerBI/sqlReferences.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid Schema:name Object:name ReferencedDatabase:name ReferencedSchema:name ReferencedObject:name ' +
    'ReferencedItemType ReferencedWorkspace:name')
Register-A2CSchema 'PowerBI/itemRuns.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid RunId:guid JobType InvokeType Status StartTimeUtc:date EndTimeUtc:date DurationSec:num FailureReason:note')
Register-A2CSchema 'PowerBI/environmentMatrix.csv' 'Solution:name ItemType Item:name DEV:name TEST:name UAT:name PROD:name Environments:num MissingIn'
Register-A2CSchema 'PowerBI/pipelineActivities.csv' ($script:A2CPbiWs + 'Pipeline:name PipelineId:guid Activity:name Parent:name ActivityType State DependsOn:name RefType RefItem:name RefItemId:guid ' +
    'RefWorkspace:name References:name StoredProcedure:name Parameters:text Connection:name Detail:text TypeProperties:json')
Register-A2CSchema 'PowerBI/notebookDependencies.csv' ($script:A2CPbiWs + 'Notebook:name NotebookId:guid Relation Target:name TargetType TargetItem:name TargetItemId:guid TargetWorkspace:name Evidence:text')
Register-A2CSchema 'PowerBI/dataflowSources.csv' ($script:A2CPbiWs + 'Dataflow:name DataflowId:guid Generation Query:name Role Connector Target:url Database:name File:name FolderPath:url ' +
    'TargetItemType TargetItem:name TargetWorkspace:name Kind Scope')
Register-A2CSchema 'PowerBI/modelPartitions.csv' ($script:A2CPbiWs + 'Model:name ModelId:guid Table:name Partition:name SourceType Mode Source:text')
Register-A2CSchema 'PowerBI/modelExpressions.csv' ($script:A2CPbiWs + 'Model:name ModelId:guid Kind Table:name Name:name FormatString DisplayFolder:name Expression:text')
Register-A2CSchema 'PowerBI/modelRoles.csv' ($script:A2CPbiWs + 'Model:name ModelId:guid Role:name ModelPermission Table:name FilterExpression:text')
Register-A2CSchema 'PowerBI/modelSources.csv' ($script:A2CPbiWs + 'Model:name ModelId:guid Query:name Connector Target:url Database:name File:name FolderPath:url TargetItemType TargetItem:name ' +
    'TargetWorkspace:name Kind Scope')
Register-A2CSchema 'PowerBI/itemDependencies.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid Relation DependsOnType DependsOn:name DependsOnId:guid DependsOnWorkspace:name Via:text')
Register-A2CSchema 'PowerBI/externalDependencies.csv' 'Kind Scope Target:url Detail:text Workspace:name UsedByType UsedBy:name UsedById:guid Via'
Register-A2CSchema 'PowerBI/onelakeStorage.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid Tables:num TableFiles:num TableBytes:num TableGB:num FilesAreaFiles:num FilesAreaBytes:num FilesAreaGB:num ' +
    'LastModified:date Status:note')
Register-A2CSchema 'PowerBI/onelakeTableSizes.csv' ($script:A2CPbiWs + 'ItemType Item:name ItemId:guid Schema:name Table:name Files:num Bytes:num MB:num LastModified:date')
Register-A2CSchema 'PowerBI/flowEnvironments.csv' 'Environment:name EnvironmentId Location Sku IsDefault CreatedTime:date Flows:num'
Register-A2CSchema 'PowerBI/flows.csv' ('Environment:name EnvironmentId Flow:name FlowId:guid State Scope CreatedTime:date LastModifiedTime:date Creator:guid Triggers Connectors UsesPowerBI ' +
    'Properties:json')

function Add-A2CPbiDependency {
    # One line of itemDependencies.csv (collected in Data['PbiDependencies'] while the other files are built).
    param($Ws, [string] $ItemType, [string] $Item, [string] $ItemId, [string] $Relation, $Ref, [string] $Via)
    if ($null -eq $Ref) { return }
    $l = $script:A2C.Data['PbiDependencies']
    $key = ($ItemId + '|' + $Relation + '|' + [string]$Ref.Id).ToLowerInvariant()
    if ($script:A2C.Data['PbiDependencyKeys'].ContainsKey($key)) { return }
    $script:A2C.Data['PbiDependencyKeys'][$key] = $true
    $l.Add((New-A2CPbiRow $Ws ([ordered]@{ ItemType = $ItemType; Item = $Item; ItemId = $ItemId; Relation = $Relation; DependsOnType = [string]$Ref.Type; DependsOn = [string]$Ref.Name
                    DependsOnId = [string]$Ref.Id; DependsOnWorkspace = [string]$Ref.Ws; Via = $Via })))
}

function Add-A2CPbiExternal {
    # One line of externalDependencies.csv, without repeats.
    param([string] $Kind, [string] $Scope, [string] $Target, [string] $Detail, [string] $Workspace, [string] $UsedByType, [string] $UsedBy, [string] $UsedById, [string] $Via)
    if (-not $Target) { return }
    $key = ($Kind + '|' + $Target + '|' + $Detail + '|' + $UsedById + '|' + $Via).ToLowerInvariant()
    if ($script:A2C.Data['PbiExternalKeys'].ContainsKey($key)) { return }
    $script:A2C.Data['PbiExternalKeys'][$key] = $true
    $script:A2C.Data['PbiExternal'].Add([ordered]@{ Kind = $Kind; Scope = $Scope; Target = $Target; Detail = $Detail; Workspace = $Workspace; UsedByType = $UsedByType; UsedBy = $UsedBy
            UsedById = $UsedById; Via = $Via })
}

function Resolve-A2CMSource {
    # The Fabric item an M source points to (by item ID, or by SQL endpoint server + database name), with the kind
    # and scope of the target. Returns @{ Ref; Kind; Scope }
    param($Source)
    $ref = Get-A2CPbiItemRef $Source.ItemId
    if ($null -eq $ref -and $Source.Target -and $Source.Database) { $ref = $script:A2C.Data['PbiSqlIndex'][([string]$Source.Target + '|' + [string]$Source.Database).ToLowerInvariant()] }
    if ($null -eq $ref) { foreach ($r in (Get-A2CPbiItemRefs ([string]$Source.Target + ' ' + [string]$Source.Database))) { $ref = $r; break } }
    $ks = Get-A2CPbiTargetKind $Source.Connector ([string]$Source.Target)
    return [pscustomobject]@{ Ref = $ref; Kind = $ks[0]; Scope = $ks[1] }
}

function New-A2CReportDependencies {
    Initialize-A2CPbiItemIndex
    $script:A2C.Data['PbiDependencies'] = [System.Collections.Generic.List[object]]::new()
    $script:A2C.Data['PbiDependencyKeys'] = @{}
    $script:A2C.Data['PbiExternal'] = [System.Collections.Generic.List[object]]::new()
    $script:A2C.Data['PbiExternalKeys'] = @{}
    $files = @('pipelineActivities', 'notebookDependencies', 'dataflowSources', 'modelPartitions', 'modelExpressions', 'modelRoles', 'modelSources')
    $rows = @{}
    foreach ($f in $files) { $rows[$f] = [System.Collections.Generic.List[object]]::new() }
    $defs = $script:A2C.Data['PbiDefs']
    if ($null -eq $defs) { $defs = @{} }

    foreach ($id in @($defs.Keys)) {
        $d = $defs[$id]
        $w = $d.Ws
        $type = [string]$d.Item.type
        $name = [string]$d.Item.displayName
        if ($type -eq 'DataPipeline') {
            $json = $null
            try { $json = ConvertFrom-A2CJson ([string]$d.Parts['pipeline-content.json']) } catch { $json = $null }
            if ($null -eq $json) { continue }
            foreach ($e in (Get-A2CPipelineActivities $json.properties.activities)) {
                $a = $e.Activity
                $info = Get-A2CPipelineActivityInfo $a
                $first = $null
                if ($info.Refs.Count -gt 0) { $first = $info.Refs[0] }
                $rows['pipelineActivities'].Add((New-A2CPbiRow $w ([ordered]@{ Pipeline = $name; PipelineId = $id; Activity = [string]$a.name; Parent = $e.Parent; ActivityType = [string]$a.type
                                State = [string]$a.state; DependsOn = (Join-A2C @($a.dependsOn | ForEach-Object { $_.activity }) ', ')
                                RefType = $(if ($null -ne $first) { $first.Type } else { '' }); RefItem = $(if ($null -ne $first) { $first.Name } else { '' })
                                RefItemId = $(if ($null -ne $first) { $first.Id } else { '' }); RefWorkspace = $(if ($null -ne $first) { $first.Ws } else { '' })
                                References = (Join-A2C @($info.Refs | ForEach-Object { $_.Type + ': ' + $(if ($_.Name) { $_.Name } else { $_.Id + ' (not in scope)' }) }))
                                StoredProcedure = $info.Procedure; Parameters = $info.Parameters; Connection = $info.Connection; Detail = $info.Detail; TypeProperties = $info.Own })))
                foreach ($r in $info.Refs) {
                    $rel = $(if ($r.Type -eq 'SemanticModel') { 'Refreshes' } elseif ($r.Type -in @('Lakehouse', 'Warehouse', 'SQLEndpoint')) { 'Uses' } else { 'Runs' })
                    Add-A2CPbiDependency $w 'DataPipeline' $name $id $rel $r ('activity ' + [string]$a.name)
                }
            }
        }
        elseif ($type -eq 'Notebook') {
            foreach ($dep in (Get-A2CNotebookDependencies (Get-A2CNotebookCode $d.Parts))) {
                $ref = Get-A2CPbiItemRef $dep.Id
                $target = $dep.Target
                $ttype = ''
                switch -Wildcard ($dep.Relation) {
                    '*Lakehouse'   { $ttype = 'Lakehouse' }
                    '*Warehouse'   { $ttype = 'Warehouse' }
                    'Environment'  { $ttype = 'Environment' }
                    '*Table'       { $ttype = 'Table' }
                    'RunsNotebook' { $ttype = 'Notebook' }
                    'OneLakePath'  { $ttype = 'OneLake path' }
                    'StoragePath'  { $ttype = 'Storage path' }
                    'Url'          { $ttype = 'Url' }
                }
                if ($dep.Relation -eq 'RunsNotebook' -and $null -eq $ref) {
                    # a notebook called by name: the one with that name in the same workspace
                    foreach ($x in $w.Items) { if ([string]$x.type -eq 'Notebook' -and [string]$x.displayName -eq $dep.Target) { $ref = Get-A2CPbiItemRef ([string]$x.id); break } }
                }
                $tws = ''
                if ($null -ne $ref) { $tws = $ref.Ws; if (-not $target) { $target = $ref.Name } }
                elseif ($dep.WorkspaceId) { $wr = Get-A2CPbiItemRef $dep.WorkspaceId; if ($null -ne $wr) { $tws = $wr.Name } }
                $rows['notebookDependencies'].Add((New-A2CPbiRow $w ([ordered]@{ Notebook = $name; NotebookId = $id; Relation = $dep.Relation; Target = $target; TargetType = $ttype
                                TargetItem = $(if ($null -ne $ref) { $ref.Name } else { '' }); TargetItemId = $(if ($null -ne $ref) { $ref.Id } else { $dep.Id }); TargetWorkspace = $tws; Evidence = $dep.Evidence })))
                if ($null -ne $ref) { Add-A2CPbiDependency $w 'Notebook' $name $id $(if ($dep.Relation -eq 'RunsNotebook') { 'Runs' } else { 'Uses' }) $ref $dep.Relation }
                if ($dep.Relation -eq 'Url' -or $dep.Relation -eq 'StoragePath') {
                    $ks = Get-A2CPbiTargetKind '' $dep.Target
                    Add-A2CPbiExternal $ks[0] $ks[1] $dep.Target '' $w.Name 'Notebook' $name $id 'notebook code'
                }
            }
        }
    }

    # dataflows: Fabric items (mashup.pq) and gen1 (the mashup document inside model.json)
    $flows = [System.Collections.Generic.List[object]]::new()
    foreach ($id in @($defs.Keys)) { if ([string]$defs[$id].Item.type -eq 'Dataflow') { $flows.Add(@($defs[$id].Ws, [string]$defs[$id].Item.displayName, $id, 'Gen2', [string]$defs[$id].Parts['mashup.pq'])) } }
    $gen1 = $script:A2C.Data['PbiDataflowDefs']
    if ($null -ne $gen1) {
        foreach ($id in @($gen1.Keys)) {
            $doc = ''
            try { $doc = [string](ConvertFrom-A2CJson $gen1[$id].Text).'pbi:mashup'.document } catch { $doc = '' }
            $flows.Add(@($gen1[$id].Ws, [string]$gen1[$id].Dataflow.name, $id, 'Gen1', $doc))
        }
    }
    foreach ($fl in $flows) {
        $w = $fl[0]
        foreach ($s in (Get-A2CMSources $fl[4])) {
            $res = Resolve-A2CMSource $s
            $role = $(if ($s.Query -like '*_DataDestination') { 'Destination' } else { 'Source' })
            $rows['dataflowSources'].Add((New-A2CPbiRow $w ([ordered]@{ Dataflow = $fl[1]; DataflowId = $fl[2]; Generation = $fl[3]; Query = $s.Query; Role = $role; Connector = $s.Connector
                            Target = $s.Target; Database = $s.Database; File = $s.File; FolderPath = $s.FolderPath
                            TargetItemType = $(if ($null -ne $res.Ref) { $res.Ref.Type } else { '' }); TargetItem = $(if ($null -ne $res.Ref) { $res.Ref.Name } else { '' })
                            TargetWorkspace = $(if ($null -ne $res.Ref) { $res.Ref.Ws } else { '' }); Kind = $res.Kind; Scope = $res.Scope })))
            if ($null -ne $res.Ref) { Add-A2CPbiDependency $w 'Dataflow' $fl[1] $fl[2] $(if ($role -eq 'Destination') { 'WritesTo' } else { 'ReadsFrom' }) $res.Ref ('query ' + $s.Query) }
            elseif ($s.Target) { Add-A2CPbiExternal $res.Kind $res.Scope $s.Target (Join-A2C @($s.Database, $s.FolderPath, $s.File) ' | ') $w.Name 'Dataflow' $fl[1] $fl[2] ('dataflow query ' + $s.Query) }
        }
    }

    # semantic models: partitions, DAX and M expressions, roles, sources
    foreach ($id in @($defs.Keys)) {
        $d = $defs[$id]
        if ([string]$d.Item.type -ne 'SemanticModel') { continue }
        $w = $d.Ws
        $name = [string]$d.Item.displayName
        $objs = Get-A2CPbiModelObjects $id
        $mtexts = [System.Collections.Generic.List[object]]::new()
        $calcTables = @{}
        foreach ($o in $objs) {
            switch ($o.Kind) {
                'partition' {
                    $src = [string]$o.Props['source']
                    $rows['modelPartitions'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Table = $o.Table; Partition = $o.Name; SourceType = $o.Value; Mode = [string]$o.Props['mode']; Source = $src })))
                    if ($o.Value -eq 'm') { $mtexts.Add(@($o.Table, $src)) }
                    elseif ($o.Value -eq 'calculated') { $calcTables[$o.Table] = $src }
                }
                'measure' {
                    $rows['modelExpressions'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Kind = 'Measure'; Table = $o.Table; Name = $o.Name; FormatString = [string]$o.Props['formatString']
                                    DisplayFolder = [string]$o.Props['displayFolder']; Expression = $o.Value })))
                }
                'column' {
                    if ($o.Value) {
                        $rows['modelExpressions'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Kind = 'CalculatedColumn'; Table = $o.Table; Name = $o.Name
                                        FormatString = [string]$o.Props['formatString']; DisplayFolder = [string]$o.Props['displayFolder']; Expression = $o.Value })))
                    }
                }
                'expression' {
                    $rows['modelExpressions'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Kind = 'Expression (M)'; Table = ''; Name = $o.Name; FormatString = ''; DisplayFolder = ''; Expression = $o.Value })))
                    $mtexts.Add(@($o.Name, $o.Value))
                }
                'role' {
                    $rows['modelRoles'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Role = $o.Name; ModelPermission = [string]$o.Props['modelPermission']; Table = ''; FilterExpression = '' })))
                }
                'tablePermission' {
                    $rows['modelRoles'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Role = $o.Table; ModelPermission = ''; Table = $o.Name; FilterExpression = $o.Value })))
                }
            }
        }
        foreach ($t in $calcTables.Keys) {
            $rows['modelExpressions'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Kind = 'CalculatedTable'; Table = $t; Name = $t; FormatString = ''; DisplayFolder = ''; Expression = $calcTables[$t] })))
        }
        foreach ($mt in $mtexts) {
            foreach ($s in (Get-A2CMSources $mt[1] $mt[0])) {
                $res = Resolve-A2CMSource $s
                $rows['modelSources'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Query = $s.Query; Connector = $s.Connector; Target = $s.Target; Database = $s.Database; File = $s.File
                                FolderPath = $s.FolderPath; TargetItemType = $(if ($null -ne $res.Ref) { $res.Ref.Type } else { '' }); TargetItem = $(if ($null -ne $res.Ref) { $res.Ref.Name } else { '' })
                                TargetWorkspace = $(if ($null -ne $res.Ref) { $res.Ref.Ws } else { '' }); Kind = $res.Kind; Scope = $res.Scope })))
                if ($null -ne $res.Ref) { Add-A2CPbiDependency $w 'SemanticModel' $name $id 'ReadsFrom' $res.Ref ('query ' + $s.Query) }
                elseif ($s.Target) { Add-A2CPbiExternal $res.Kind $res.Scope $s.Target (Join-A2C @($s.Database, $s.FolderPath, $s.File) ' | ') $w.Name 'SemanticModel' $name $id ('model query ' + $s.Query) }
            }
        }
    }

    # what the APIs already told: report -> model, model -> dataflow, data sources, shortcuts, connections
    foreach ($w in @($script:A2C.Data['PbiWorkspaces'])) {
        foreach ($x in $w.Reports) { if ($x.datasetId) { Add-A2CPbiDependency $w 'Report' ([string]$x.name) ([string]$x.id) 'Uses' (Get-A2CPbiItemRef ([string]$x.datasetId)) 'report' } }
        foreach ($l in $w.UpstreamDataflows) {
            $mr = Get-A2CPbiItemRef ([string]$l.datasetObjectId)
            if ($null -ne $mr) { Add-A2CPbiDependency $w 'SemanticModel' $mr.Name $mr.Id 'ReadsFrom' (Get-A2CPbiItemRef ([string]$l.dataflowObjectId)) 'upstream dataflow' }
        }
        foreach ($pair in @(@('SemanticModel', $w.Datasets, 'datasetDatasources', 'id', 'name'), @('Dataflow', $w.Dataflows, 'dataflowDatasources', 'objectId', 'name'), @('Report', $w.Reports, 'reportDatasources', 'id', 'name'))) {
            foreach ($x in $pair[1]) {
                $xid = [string]$x.($pair[3])
                foreach ($ds in @(Get-A2CPbiSub $w $pair[2] $xid)) {
                    $c = $ds.connectionDetails
                    $target = [string](Get-A2CFirst $c @('server', 'url', 'path', 'account', 'domain'))
                    if (-not $target) { $target = [string]$ds.connectionString }
                    $ks = Get-A2CPbiTargetKind ([string]$ds.datasourceType) $target
                    $ref = $null
                    if ($c.server -and $c.database) { $ref = $script:A2C.Data['PbiSqlIndex'][([string]$c.server + '|' + [string]$c.database).ToLowerInvariant()] }
                    if ($null -ne $ref) { Add-A2CPbiDependency $w $pair[0] ([string]$x.($pair[4])) $xid 'ReadsFrom' $ref 'data source' }
                    else { Add-A2CPbiExternal $ks[0] $ks[1] $target ([string]$c.database) $w.Name $pair[0] ([string]$x.($pair[4])) $xid ('data source ' + [string]$ds.datasourceType) }
                }
            }
        }
        foreach ($x in $w.Items) {
            foreach ($s in @(Get-A2CPbiSub $w 'shortcuts' ([string]$x.id))) {
                $sc = Get-A2CPbiShortcutTarget $s
                if ($null -ne $sc.Ref) { Add-A2CPbiDependency $w ([string]$x.type) ([string]$x.displayName) ([string]$x.id) 'ShortcutTo' $sc.Ref ('shortcut ' + [string]$s.name) }
                else {
                    $ks = Get-A2CPbiTargetKind ([string]$s.target.type) $sc.Location
                    $kind = $(if ($ks[0] -eq 'Other') { [string]$s.target.type } else { $ks[0] })
                    # a Dataverse link makes one shortcut per table: one line per environment is enough here (shortcuts.csv has them all)
                    $sub = $(if ([string]$s.target.type -eq 'Dataverse') { '' } else { $sc.Subpath })
                    Add-A2CPbiExternal $kind $ks[1] $sc.Location $sub $w.Name ([string]$x.type) ([string]$x.displayName) ([string]$x.id) 'shortcut'
                }
            }
        }
    }
    $ft = $script:A2C.Data['FabricTenant']
    if ($null -ne $ft) {
        foreach ($c in $ft.Connections) {
            $path = [string]$c.connectionDetails.path
            $ks = Get-A2CPbiTargetKind ([string]$c.connectionDetails.type) $path
            $kind = $(if ($ks[0] -eq 'Other') { [string]$c.connectionDetails.type } else { $ks[0] })
            Add-A2CPbiExternal $kind $ks[1] $path ([string]$c.connectivityType) '' 'Connection' ([string]$c.displayName) ([string]$c.id) 'connection'
        }
    }

    foreach ($f in $files) { Add-A2CRows ('PowerBI/' + $f + '.csv') $rows[$f] }
    Add-A2CRows 'PowerBI/itemDependencies.csv' @($script:A2C.Data['PbiDependencies'] | Sort-Object { $_['Workspace'] }, { $_['ItemType'] }, { $_['Item'] })
    Add-A2CRows 'PowerBI/externalDependencies.csv' @($script:A2C.Data['PbiExternal'] | Sort-Object { $_['Kind'] }, { $_['Target'] }, { $_['UsedBy'] })
    Write-A2CLog -Level Ok ('Dependencies: {0} pipeline activities, {1} notebook links, {2} dataflow sources, {3} model expressions, {4} item links, {5} external targets' -f $rows['pipelineActivities'].Count,
        $rows['notebookDependencies'].Count, $rows['dataflowSources'].Count, $rows['modelExpressions'].Count, $script:A2C.Data['PbiDependencies'].Count, $script:A2C.Data['PbiExternal'].Count)
}

function Get-A2CPbiShortcutTarget {
    # Where a shortcut points: connection, location and subpath for external targets, the item for OneLake ones.
    param($Shortcut)
    $t = $Shortcut.target
    $inner = $null
    if ($null -ne $t) { foreach ($p in $t.PSObject.Properties) { if ($p.Name -ne 'type' -and $null -ne $p.Value -and $p.Value -isnot [string]) { $inner = $p.Value; break } } }
    $ref = $null
    if ($null -ne $script:A2C.Data['PbiItemIndex']) { $ref = Get-A2CPbiItemRef ([string]$inner.itemId) }
    $conn = ''
    $ci = $script:A2C.Data['PbiConnectionIndex']
    if ($null -ne $ci -and $inner.connectionId) { $c = $ci[([string]$inner.connectionId).ToLowerInvariant()]; if ($null -ne $c) { $conn = [string]$c.displayName } }
    $loc = [string](Get-A2CFirst $inner @('location', 'environmentDomain', 'url'))
    if (-not $loc -and $inner.workspaceId) { $loc = 'onelake:' + [string]$inner.workspaceId + '/' + [string]$inner.itemId }
    return [pscustomobject]@{ Ref = $ref; ConnectionId = [string]$inner.connectionId; Connection = $conn; Location = $loc
        Subpath = [string](Get-A2CFirst $inner @('subpath', 'path', 'tableName', 'deltaLakeFolder')) }
}

function Get-A2CPbiEnvironment {
    # "Finance - DEV" -> @('finance', 'DEV'); a name without an environment suffix counts as PROD.
    param([string] $Name)
    $m = [regex]::Match($Name, '(?i)^(.*?)[\s_\-]+(DEV|TEST|UAT|PROD)\s*$')
    if ($m.Success) { return @($m.Groups[1].Value.Trim(), $m.Groups[2].Value.ToUpperInvariant()) }
    return @($Name.Trim(), 'PROD')
}

function New-A2CPbiEnvironmentMatrix {
    # Items of the same solution side by side per environment (workspaces named "<solution> - DEV/TEST/UAT/PROD"):
    # shows what exists only in DEV or is missing in PROD. Names are matched without their environment suffix.
    $rows = [System.Collections.Generic.List[object]]::new()
    $types = @('Report', 'SemanticModel', 'Notebook', 'DataPipeline', 'Dataflow', 'Lakehouse', 'Warehouse')
    $sol = @{}
    foreach ($w in @($script:A2C.Data['PbiWorkspaces'])) {
        $e = Get-A2CPbiEnvironment $w.Name
        $k = $e[0].ToLowerInvariant()
        if (-not $sol.ContainsKey($k)) { $sol[$k] = [pscustomobject]@{ Name = $e[0]; Envs = @{}; Items = @{} } }
        $sol[$k].Envs[$e[1]] = $true
        foreach ($x in $w.Items) {
            $type = [string]$x.type
            $n = [string]$x.displayName
            if ($types -notcontains $type -or $n -match '(?i)^(DataflowsStaging|StagingLakehouseForDataflows|StagingWarehouseForDataflows)') { continue }
            $base = ($n -replace '(?i)[\s_\-]*(dev|test|uat|prod)$', '')
            $ik = ($type + '|' + $base).ToLowerInvariant()
            if (-not $sol[$k].Items.ContainsKey($ik)) { $sol[$k].Items[$ik] = [pscustomobject]@{ Type = $type; Name = $base; In = @{} } }
            $have = $sol[$k].Items[$ik].In
            if (-not $have.ContainsKey($e[1])) { $have[$e[1]] = [System.Collections.Generic.List[string]]::new() }
            $have[$e[1]].Add($n)
        }
    }
    foreach ($k in @($sol.Keys | Sort-Object)) {
        $s = $sol[$k]
        if ($s.Envs.Count -lt 2) { continue }
        foreach ($ik in @($s.Items.Keys | Sort-Object)) {
            $it = $s.Items[$ik]
            $row = [ordered]@{ Solution = $s.Name; ItemType = $it.Type; Item = $it.Name }
            $missing = [System.Collections.Generic.List[string]]::new()
            foreach ($env in @('DEV', 'TEST', 'UAT', 'PROD')) {
                $row[$env] = $(if ($it.In.ContainsKey($env)) { Join-A2C $it.In[$env] } else { '' })
                if ($s.Envs.ContainsKey($env) -and -not $it.In.ContainsKey($env)) { $missing.Add($env) }
            }
            $row['Environments'] = $it.In.Count
            $row['MissingIn'] = [string]::Join(', ', $missing)
            $rows.Add($row)
        }
    }
    Add-A2CRows 'PowerBI/environmentMatrix.csv' $rows
}

function New-A2CPbiWorkspacePlan {
    # -WorkspaceListFile: a list of workspaces somebody expects (e.g. a migration scope sheet) checked against what
    # this account sees. Its columns are carried over, so owner / purpose / "to migrate" end up next to the facts.
    $path = [string]$script:A2C.Options.WorkspaceListFile
    if (-not $path -or $script:A2C.Out.ContainsKey('PowerBI/workspacePlan.csv')) { return }
    $listed = Read-A2CTable -Path $path -Sheet 'Workspaces'
    if ($listed.Count -eq 0) { Write-A2CLog -Level Warn ('-WorkspaceListFile has no rows: ' + $path); return }
    $cols = @($listed[0].Keys)
    $nameCol = $cols[0]
    foreach ($c in $cols) { if ($c -match '(?i)^workspace') { $nameCol = $c; break } }
    $extra = @($cols | Where-Object { $_ -ne $nameCol } | ForEach-Object { @{ Source = $_; Column = 'Listed_' + ($_ -replace '[^A-Za-z0-9]+', '_').Trim('_') } })
    Register-A2CSchema 'PowerBI/workspacePlan.csv' ('ListedWorkspace:name Status Check:note Workspace:name WorkspaceId:guid Type State Items:num HasContent ChangedLast30Days ' +
        'CapacityId:guid InScope Source ' + (Join-A2C @($extra | ForEach-Object { $_.Column + ':text' }) ' '))
    # the admin list when the account is a Fabric administrator (the whole tenant), otherwise the member's workspaces
    $facts = Get-A2CPbiAdminFacts
    $adminSource = ($facts.Count -gt 0)
    $all = [System.Collections.Generic.List[object]]::new()
    if ($adminSource) {
        foreach ($f in $facts) { $all.Add([pscustomobject]@{ Name = $f.Name; Id = $f.Id; Type = $f.Type; CapacityId = $f.CapacityId; IsPersonal = $f.IsPersonal; State = $f.State; Active = $f.Active
                    Items = $f.Items; HasContent = $f.HasContent; ChangedRecently = $f.ChangedRecently }) }
    }
    else {
        foreach ($w in @($script:A2C.Data['PbiWorkspacesAll'])) {
            $n = $null
            if ($w.Status -eq 'Collected' -or $w.FabricStatus -eq 'Collected') { $n = [math]::Max(@($w.Items).Count, @($w.Reports).Count + @($w.Datasets).Count + @($w.Dashboards).Count + @($w.Dataflows).Count) }
            $all.Add([pscustomobject]@{ Name = $w.Name; Id = $w.Id; Type = $w.Type; CapacityId = $w.CapacityId; IsPersonal = ($w.Type -eq 'Personal'); State = ''; Active = $true
                    Items = $n; HasContent = $(if ($null -ne $n) { $n -gt 0 } else { $null }); ChangedRecently = $null })
        }
    }
    $scopeKnown = ($null -ne $script:A2C.Data['PbiWorkspaces'])
    $scope = @{}
    if ($scopeKnown) { foreach ($w in $script:A2C.Data['PbiWorkspaces']) { $scope[$w.Id] = $true } }
    $notFound = $(if ($adminSource) { 'Not in the tenant' } else { 'Not visible to this account' })
    $norm = { param($s) (([string]$s) -replace '[^A-Za-z0-9]+', '').ToLowerInvariant() }
    $used = @{}
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($l in $listed) {
        $name = [string]$l[$nameCol]
        if (-not $name) { continue }
        $hit = $null
        $status = 'Found'
        foreach ($w in $all) { if ((& $norm $w.Name) -eq (& $norm $name)) { $hit = $w; break } }
        if ($null -ne $hit -and $hit.Name -ne $name) { $status = 'Found (spelled differently)' }
        if ($null -eq $hit) {
            # closest name within a few typing errors, e.g. "Finance Wld" for "Finance Wild"
            $best = 99
            foreach ($w in $all) {
                if ($used.ContainsKey($w.Id)) { continue }
                $dist = Get-A2CEditDistance (& $norm $name) (& $norm $w.Name)
                if ($dist -lt $best) { $best = $dist; $hit = $w }
            }
            if ($best -le 2) { $status = 'Found (spelled differently)' } else { $hit = $null; $status = $notFound }
        }
        $row = [ordered]@{ ListedWorkspace = $name; Status = $status; Check = ''; Workspace = ''; WorkspaceId = ''; Type = ''; State = ''; Items = $null; HasContent = $null
            ChangedLast30Days = $null; CapacityId = ''; InScope = ''; Source = $(if ($adminSource) { 'Admin API' } else { 'Member' }) }
        if ($null -ne $hit) {
            $used[$hit.Id] = $true
            $row['Workspace'] = $hit.Name; $row['WorkspaceId'] = $hit.Id; $row['Type'] = $hit.Type; $row['State'] = $hit.State; $row['Items'] = $hit.Items; $row['HasContent'] = $hit.HasContent
            $row['ChangedLast30Days'] = $hit.ChangedRecently; $row['CapacityId'] = $hit.CapacityId; $row['InScope'] = $(if ($scopeKnown) { [bool]$scope[$hit.Id] } else { '' })
            # the reverse question: listed, but is there anything to migrate?
            if (-not $hit.Active) { $row['Check'] = 'On the list but not active (' + $hit.State + ')' }
            elseif ($hit.IsPersonal) { $row['Check'] = 'On the list but a personal workspace' }
            elseif ($hit.HasContent -eq $false) { $row['Check'] = 'On the list but empty' }
        }
        foreach ($x in $extra) { $row[$x.Column] = $l[$x.Source] }
        $rows.Add($row)
    }
    foreach ($w in @($all | Sort-Object Name)) {
        if ($used.ContainsKey($w.Id) -or $w.IsPersonal -or -not $w.Active) { continue }
        $check = $(if ($w.HasContent -eq $true) { 'Shared with content, not on the list' } elseif ($w.HasContent -eq $false) { 'Shared, empty' } else { 'Shared, content unknown' })
        $rows.Add([ordered]@{ ListedWorkspace = ''; Status = 'Not on the list'; Check = $check; Workspace = $w.Name; WorkspaceId = $w.Id; Type = $w.Type; State = $w.State; Items = $w.Items
                HasContent = $w.HasContent; ChangedLast30Days = $w.ChangedRecently; CapacityId = $w.CapacityId; InScope = $(if ($scopeKnown) { [bool]$scope[$w.Id] } else { '' })
                Source = $(if ($adminSource) { 'Admin API' } else { 'Member' }) })
    }
    Add-A2CRows 'PowerBI/workspacePlan.csv' $rows
    $missing = @($rows | Where-Object { $_['Status'] -eq $notFound }).Count
    Write-A2CLog -Level $(if ($missing -gt 0) { 'Notice' } else { 'Ok' }) ('Workspace list: {0} listed, {1} {3}, {2} not on the list' -f $listed.Count, $missing,
        @($rows | Where-Object { $_['Status'] -eq 'Not on the list' }).Count, $notFound.ToLowerInvariant())
}

function New-A2CReportFabricSql {
    $map = $script:A2C.Data['FabricSql']
    $files = @('sqlObjects', 'sqlColumns', 'sqlModules', 'sqlReferences')
    $rows = @{}
    foreach ($f in $files) { $rows[$f] = [System.Collections.Generic.List[object]]::new() }
    foreach ($w in @($script:A2C.Data['PbiWorkspaces'])) {
        $byName = @{}   # database name -> item of this workspace, for references between databases
        foreach ($x in $w.Lakehouses) { $byName[([string]$x.displayName).ToLowerInvariant()] = @('Lakehouse', $w.Name) }
        foreach ($x in $w.Warehouses) { $byName[([string]$x.displayName).ToLowerInvariant()] = @('Warehouse', $w.Name) }
        $dbs = @($w.Lakehouses | ForEach-Object { , @('Lakehouse', $_) }) + @($w.Warehouses | ForEach-Object { , @('Warehouse', $_) })
        foreach ($db in $dbs) {
            $x = $null
            if ($null -ne $map) { $x = $map[[string]$db[1].id] }
            if ($null -eq $x -or -not $x.Ok) { continue }
            $base = [ordered]@{ ItemType = $db[0]; Item = [string]$db[1].displayName; ItemId = [string]$db[1].id }
            $colCount = @{}
            foreach ($c in $x.Columns) {
                $k = ([string]$c['schema'] + '.' + [string]$c['table']).ToLowerInvariant()
                $colCount[$k] = [int]$colCount[$k] + 1
                $v = [ordered]@{}; foreach ($bk in $base.Keys) { $v[$bk] = $base[$bk] }
                $v['Schema'] = [string]$c['schema']; $v['Table'] = [string]$c['table']; $v['Column'] = [string]$c['column']; $v['Ordinal'] = $c['ordinal']; $v['DataType'] = [string]$c['type']
                $v['Length'] = $c['length']; $v['Precision'] = $c['precision']; $v['Scale'] = $c['scale']; $v['Nullable'] = [string]$c['nullable']
                $rows['sqlColumns'].Add((New-A2CPbiRow $w $v))
            }
            $refs = @{}
            foreach ($r in $x.References) {
                $k = ([string]$r['schema'] + '.' + [string]$r['name']).ToLowerInvariant()
                if (-not $refs.ContainsKey($k)) { $refs[$k] = [System.Collections.Generic.List[string]]::new() }
                if (-not $refs[$k].Contains([string]$r['ref_database'])) { $refs[$k].Add([string]$r['ref_database']) }
                $hit = $byName[([string]$r['ref_database']).ToLowerInvariant()]
                $v = [ordered]@{}; foreach ($bk in $base.Keys) { $v[$bk] = $base[$bk] }
                $v['Schema'] = [string]$r['schema']; $v['Object'] = [string]$r['name']; $v['ReferencedDatabase'] = [string]$r['ref_database']; $v['ReferencedSchema'] = [string]$r['ref_schema']
                $v['ReferencedObject'] = [string]$r['ref_entity']; $v['ReferencedItemType'] = $(if ($null -ne $hit) { $hit[0] } else { '' }); $v['ReferencedWorkspace'] = $(if ($null -ne $hit) { $hit[1] } else { '' })
                $rows['sqlReferences'].Add((New-A2CPbiRow $w $v))
            }
            $defLen = @{}
            foreach ($m in $x.Modules) {
                $k = ([string]$m['schema'] + '.' + [string]$m['name']).ToLowerInvariant()
                $def = [string]$m['definition']
                $defLen[$k] = $def.Length
                $v = [ordered]@{}; foreach ($bk in $base.Keys) { $v[$bk] = $base[$bk] }
                $v['Schema'] = [string]$m['schema']; $v['Object'] = [string]$m['name']; $v['ObjectType'] = [string]$script:A2CSqlObjectTypes[[string]$m['type']]; $v['DefinitionLength'] = $def.Length
                $v['ReferencedDatabases'] = $(if ($refs.ContainsKey($k)) { Join-A2C $refs[$k] } else { '' }); $v['Definition'] = $def
                $rows['sqlModules'].Add((New-A2CPbiRow $w $v))
            }
            foreach ($o in $x.Objects) {
                $k = ([string]$o['schema'] + '.' + [string]$o['name']).ToLowerInvariant()
                $v = [ordered]@{}; foreach ($bk in $base.Keys) { $v[$bk] = $base[$bk] }
                $v['Schema'] = [string]$o['schema']; $v['Object'] = [string]$o['name']; $v['ObjectType'] = [string]$script:A2CSqlObjectTypes[[string]$o['type']]; $v['Created'] = $o['create_date']
                $v['Modified'] = $o['modify_date']; $v['Columns'] = $colCount[$k]; $v['DefinitionLength'] = $defLen[$k]
                $rows['sqlObjects'].Add((New-A2CPbiRow $w $v))
            }
        }
    }
    foreach ($f in $files) { Add-A2CRows ('PowerBI/' + $f + '.csv') $rows[$f] }
}

function New-A2CReportFabricStorage {
    $sum = [System.Collections.Generic.List[object]]::new()
    $det = [System.Collections.Generic.List[object]]::new()
    foreach ($e in @($script:A2C.Data['FabricStorage'])) {
        if ($null -eq $e) { continue }
        $base = [ordered]@{ ItemType = $e.Type; Item = [string]$e.Item.displayName; ItemId = [string]$e.Item.id }
        $files = 0; $bytes = [long]0; $last = $null; $tables = 0
        if ($null -ne $e.Tables) {
            foreach ($k in @($e.Tables.Keys | Sort-Object)) {
                $g = $e.Tables[$k]
                $seg = $k.Split('/')        # Tables/<table> or Tables/<schema>/<table>
                $tables++; $files += $g.Files; $bytes += $g.Bytes
                if ($null -ne $g.LastModified -and ($null -eq $last -or $g.LastModified -gt $last)) { $last = $g.LastModified }
                $v = [ordered]@{}; foreach ($bk in $base.Keys) { $v[$bk] = $base[$bk] }
                $v['Schema'] = $(if ($seg.Count -ge 3) { $seg[1] } else { '' }); $v['Table'] = $seg[$seg.Count - 1]; $v['Files'] = $g.Files; $v['Bytes'] = $g.Bytes
                $v['MB'] = [math]::Round($g.Bytes / 1MB, 2); $v['LastModified'] = $g.LastModified
                $det.Add((New-A2CPbiRow $e.Ws $v))
            }
        }
        $ff = $null; $fb = $null
        if ($null -ne $e.Files) { $ff = 0; $fb = [long]0; foreach ($g in $e.Files.Values) { $ff += $g.Files; $fb += $g.Bytes } }
        $status = 'Read'
        if ($e.Error) { $status = 'Not read: ' + $e.Error } elseif ($e.Truncated) { $status = 'Read in part (more than 1,000,000 files)' }
        $v = [ordered]@{}; foreach ($bk in $base.Keys) { $v[$bk] = $base[$bk] }
        $v['Tables'] = $(if ($e.Error) { $null } else { $tables }); $v['TableFiles'] = $(if ($e.Error) { $null } else { $files }); $v['TableBytes'] = $(if ($e.Error) { $null } else { $bytes })
        $v['TableGB'] = $(if ($e.Error) { $null } else { [math]::Round($bytes / 1GB, 3) }); $v['FilesAreaFiles'] = $ff; $v['FilesAreaBytes'] = $fb
        $v['FilesAreaGB'] = $(if ($null -ne $fb) { [math]::Round($fb / 1GB, 3) } else { $null }); $v['LastModified'] = $last; $v['Status'] = $status
        $sum.Add((New-A2CPbiRow $e.Ws $v))
    }
    Add-A2CRows 'PowerBI/onelakeStorage.csv' $sum
    Add-A2CRows 'PowerBI/onelakeTableSizes.csv' $det
}

function New-A2CReportPowerAutomate {
    $pa = $script:A2C.Data['PowerAutomate']
    if ($null -eq $pa -or -not (Test-A2CDataset 'PowerAutomate')) { return }
    $count = @{}
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($x in $pa.Flows) {
        $e = $x.Environment; $f = $x.Flow; $p = $f.properties
        $count[[string]$e.name] = [int]$count[[string]$e.name] + 1
        $connectors = Get-A2CUnique @(@($p.connectionReferences.PSObject.Properties | ForEach-Object { [string](Get-A2CFirst $_.Value @('displayName', 'apiName', 'id')) }) +
            @($p.definitionSummary.actions | ForEach-Object { [string]$_.api.name }))
        $rows.Add([ordered]@{ Environment = [string]$e.properties.displayName; EnvironmentId = [string]$e.name; Flow = [string]$p.displayName; FlowId = [string]$f.name; State = [string]$p.state
                Scope = $x.Scope; CreatedTime = $p.createdTime; LastModifiedTime = $p.lastModifiedTime; Creator = [string](Get-A2CFirst $p @('creator.userId', 'creator.objectId'))
                Triggers = (Join-A2C (Get-A2CUnique @($p.definitionSummary.triggers | ForEach-Object { (([string]$_.type) + ' ' + ([string]$_.kind) + ' ' + ([string]$_.swaggerOperationId)).Trim() })))
                Connectors = (Join-A2C $connectors); UsesPowerBI = [bool]((Join-A2C $connectors) -match '(?i)power ?bi|fabric'); Properties = $p })
    }
    $envRows = [System.Collections.Generic.List[object]]::new()
    foreach ($e in $pa.Environments) {
        $envRows.Add([ordered]@{ Environment = [string]$e.properties.displayName; EnvironmentId = [string]$e.name; Location = [string]$e.location; Sku = [string]$e.properties.environmentSku
                IsDefault = $e.properties.isDefault; CreatedTime = $e.properties.createdTime; Flows = [int]$count[[string]$e.name] })
    }
    Add-A2CRows 'PowerBI/flowEnvironments.csv' $envRows
    Add-A2CRows 'PowerBI/flows.csv' $rows
}

function New-A2CReportPowerBIModels {
    $files = @('modelTables', 'modelColumns', 'modelMeasures', 'modelRelationships')
    $rows = @{}
    foreach ($f in $files) { $rows[$f] = [System.Collections.Generic.List[object]]::new() }
    foreach ($m in @($script:A2C.Data['PbiModels'])) {
        if ($null -eq $m) { continue }
        $w = $m.Ws
        $name = [string]$m.Dataset.name
        $id = [string]$m.Dataset.id
        foreach ($r in $m.Tables) {
            $rows['modelTables'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Table = Get-A2CDaxValue $r @('Name'); StorageMode = Get-A2CDaxValue $r @('StorageMode')
                            IsHidden = Get-A2CDaxValue $r @('IsHidden'); DataCategory = Get-A2CDaxValue $r @('DataCategory'); Description = Get-A2CDaxValue $r @('Description')
                            Expression = Get-A2CDaxValue $r @('Expression'); Properties = $r })))
        }
        foreach ($r in $m.Columns) {
            $rows['modelColumns'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Table = Get-A2CDaxValue $r @('Table'); Column = Get-A2CDaxValue $r @('Name')
                            DataType = Get-A2CDaxValue $r @('DataType'); Type = Get-A2CDaxValue $r @('Type'); IsHidden = Get-A2CDaxValue $r @('IsHidden'); IsKey = Get-A2CDaxValue $r @('IsKey')
                            FormatString = Get-A2CDaxValue $r @('FormatString'); DisplayFolder = Get-A2CDaxValue $r @('DisplayFolder'); Description = Get-A2CDaxValue $r @('Description')
                            Expression = Get-A2CDaxValue $r @('Expression'); Properties = $r })))
        }
        foreach ($r in $m.Measures) {
            $rows['modelMeasures'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Table = Get-A2CDaxValue $r @('Table'); Measure = Get-A2CDaxValue $r @('Name')
                            DataType = Get-A2CDaxValue $r @('DataType'); FormatString = Get-A2CDaxValue $r @('FormatString', 'FormatStringDefinition'); IsHidden = Get-A2CDaxValue $r @('IsHidden')
                            DisplayFolder = Get-A2CDaxValue $r @('DisplayFolder'); State = Get-A2CDaxValue $r @('State'); Description = Get-A2CDaxValue $r @('Description')
                            Expression = Get-A2CDaxValue $r @('Expression'); Properties = $r })))
        }
        foreach ($r in $m.Relationships) {
            $rows['modelRelationships'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Relationship = Get-A2CDaxValue $r @('Relationship', 'Name')
                            FromTable = Get-A2CDaxValue $r @('FromTable'); FromColumn = Get-A2CDaxValue $r @('FromColumn'); FromCardinality = Get-A2CDaxValue $r @('FromCardinality')
                            ToTable = Get-A2CDaxValue $r @('ToTable'); ToColumn = Get-A2CDaxValue $r @('ToColumn'); ToCardinality = Get-A2CDaxValue $r @('ToCardinality')
                            IsActive = Get-A2CDaxValue $r @('IsActive'); CrossFilteringBehavior = Get-A2CDaxValue $r @('CrossFilteringBehavior'); Properties = $r })))
        }
    }
    foreach ($f in $files) { Add-A2CRows ('PowerBI/' + $f + '.csv') $rows[$f] }
}

$script:A2CPbiScanLists = @('reports', 'datasets', 'dashboards', 'dataflows', 'datamarts')   # the classic lists of a scanned workspace

function New-A2CReportPowerBIAdmin {
    $a = $script:A2C.Data['PbiAdmin']
    if ($null -eq $a -or -not (Test-A2CDataset 'PowerBIAdmin')) { return }   # no role: no (empty) admin files
    $files = @('adminWorkspaces', 'adminWorkspaceUsers', 'adminItems', 'adminItemUsers', 'adminDatasourceInstances', 'adminModelTables', 'adminModelExpressions',
        'tenantSettings', 'adminCapacities', 'adminApps', 'widelyShared', 'adminDomains')
    $rows = @{}
    foreach ($f in $files) { $rows[$f] = [System.Collections.Generic.List[object]]::new() }

    $itemCount = @{}
    foreach ($sw in $a.Scan) {
        $w = [pscustomobject]@{ Name = [string]$sw.name; Id = [string]$sw.id }
        $n = 0
        foreach ($prop in $sw.PSObject.Properties) {
            # every array of objects with an id is a list of items: reports, datasets, ... and Fabric types (Lakehouse, Notebook, ...)
            if ($prop.Name -eq 'users' -or -not ($prop.Value -is [array])) { continue }
            $type = $prop.Name
            foreach ($x in $prop.Value) {
                if ($null -eq $x) { continue }
                if ($null -eq $x.PSObject.Properties['id'] -and $null -eq $x.PSObject.Properties['objectId']) { continue }
                $n++
                $id = [string](Get-A2CFirst $x @('id', 'objectId'))
                $name = [string](Get-A2CFirst $x @('name', 'displayName'))
                $up = [System.Collections.Generic.List[string]]::new()
                foreach ($u in @($x.upstreamDataflows)) { if ($null -ne $u) { $up.Add('dataflow ' + [string]$u.targetDataflowId) } }
                foreach ($u in @($x.upstreamDatasets)) { if ($null -ne $u) { $up.Add('model ' + [string]$u.targetDatasetId) } }
                foreach ($u in @($x.upstreamDatamarts)) { if ($null -ne $u) { $up.Add('datamart ' + [string]$u.targetDatamartId) } }
                foreach ($u in @($x.relations)) { if ($null -ne $u) { $up.Add(([string]$u.dependentOnArtifactId)) } }
                $users = @($x.users | Where-Object { $null -ne $_ })
                $rows['adminItems'].Add((New-A2CPbiRow $w ([ordered]@{ ItemType = $type; Item = $name; ItemId = $id; ConfiguredBy = [string]$x.configuredBy
                                CreatedDate = (Get-A2CFirst $x @('createdDate', 'createdDateTime')); ModifiedDate = (Get-A2CFirst $x @('modifiedDateTime', 'modifiedDate', 'lastUpdatedDate'))
                                ModifiedBy = [string]$x.modifiedBy; Endorsement = [string]$x.endorsementDetails.endorsement; CertifiedBy = [string]$x.endorsementDetails.certifiedBy
                                SensitivityLabelId = [string]$x.sensitivityLabel.labelId; ModelId = [string]$x.datasetId; TargetStorageMode = [string]$x.targetStorageMode
                                ContentProviderType = [string]$x.contentProviderType; Description = [string]$x.description; Users = $users.Count
                                DatasourceIds = (Join-A2C @($x.datasourceUsages | ForEach-Object { $_.datasourceInstanceId })); Upstream = (Join-A2C $up) })))
                foreach ($u in $users) {
                    $right = ''
                    foreach ($up2 in $u.PSObject.Properties) { if ($up2.Name -like '*AccessRight') { $right = [string]$up2.Value; break } }
                    $rows['adminItemUsers'].Add((New-A2CPbiRow $w ([ordered]@{ ItemType = $type; Item = $name; ItemId = $id; DisplayName = [string]$u.displayName
                                    Identifier = [string](Get-A2CFirst $u @('identifier', 'emailAddress', 'graphId')); PrincipalType = [string]$u.principalType; AccessRight = $right })))
                }
                if ($type -ne 'datasets') { continue }
                foreach ($e in @($x.expressions)) {
                    if ($null -eq $e) { continue }
                    $rows['adminModelExpressions'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Kind = 'Expression'; Table = ''; Name = [string]$e.name
                                    Expression = [string]$e.expression; Description = [string]$e.description })))
                }
                foreach ($tb in @($x.tables)) {
                    if ($null -eq $tb) { continue }
                    $src = Join-A2C @($tb.source | ForEach-Object { $_.expression }) ' | '
                    $measures = @($tb.measures | Where-Object { $null -ne $_ })
                    $rows['adminModelTables'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Table = [string]$tb.name; IsHidden = $tb.isHidden; StorageMode = [string]$tb.storageMode
                                    Columns = @($tb.columns | Where-Object { $null -ne $_ }).Count; Measures = $measures.Count; Source = $src })))
                    foreach ($ms in $measures) {
                        $rows['adminModelExpressions'].Add((New-A2CPbiRow $w ([ordered]@{ Model = $name; ModelId = $id; Kind = 'Measure'; Table = [string]$tb.name; Name = [string]$ms.name
                                        Expression = [string]$ms.expression; Description = [string]$ms.description })))
                    }
                }
            }
        }
        $itemCount[$w.Id.ToLowerInvariant()] = $n
    }
    $cs = [bool]$script:A2C.Options.ClientScope
    $facts = @{}
    foreach ($f in (Get-A2CPbiAdminFacts)) { $facts[$f.Id.ToLowerInvariant()] = $f }
    foreach ($g in $a.Workspaces) {
        $w = [pscustomobject]@{ Name = [string]$g.name; Id = [string]$g.id }
        $users = @($g.users | Where-Object { $null -ne $_ })
        $f = $facts[$w.Id.ToLowerInvariant()]
        $rows['adminWorkspaces'].Add((New-A2CPbiRow $w ([ordered]@{ Type = [string]$g.type; Scope = $f.Scope; IsPersonal = $f.IsPersonal; State = [string]$g.state
                        IsOnDedicatedCapacity = $g.isOnDedicatedCapacity; CapacityId = [string]$g.capacityId; Description = $(if ($cs) { '' } else { [string]$g.description })
                        DefaultStorageFormat = [string]$g.defaultDatasetStorageFormat; Users = $users.Count; Admins = $f.Admins; AdminsDeparted = $f.AdminsDeparted; NoAdmin = $f.NoAdmin; Items = $f.Items
                        HasContent = $f.HasContent; LastItemChange = $f.LastItemChange; ChangedLast30Days = $f.ChangedRecently; RecentActivities = $f.RecentActivities; ActiveUsers = $f.ActiveUsers
                        DatasourceTypes = $f.DatasourceTypes; Datasources = $f.Datasources; UsesGateway = $f.UsesGateway; CreatedDate = $f.CreatedDate; ModifiedDate = $f.ModifiedDate })))
        foreach ($u in $users) {
            $rows['adminWorkspaceUsers'].Add((New-A2CPbiRow $w ([ordered]@{ DisplayName = [string]$u.displayName; Identifier = [string]$u.identifier; EmailAddress = [string]$u.emailAddress
                            PrincipalType = [string]$u.principalType; AccessRight = [string]$u.groupUserAccessRight; UserType = [string]$u.userType; GraphId = [string]$u.graphId })))
        }
    }
    $use = $script:A2C.Data['PbiAdminInstanceUse']
    foreach ($d in $a.DatasourceInstances) {
        # -ClientScope: the type and the gateway only, never the server, database, URL or path
        $rows['adminDatasourceInstances'].Add([ordered]@{ DatasourceId = [string]$d.datasourceId; DatasourceType = [string]$d.datasourceType; GatewayId = [string]$d.gatewayId
                Workspaces = $(if ($null -ne $use) { [int]$use[([string]$d.datasourceId).ToLowerInvariant()] } else { $null }); ConnectionDetails = $(if ($cs) { $null } else { $d.connectionDetails }) })
    }
    foreach ($s in $a.TenantSettings) {
        $rows['tenantSettings'].Add([ordered]@{ SettingName = [string]$s.settingName; Title = [string]$s.title; TenantSettingGroup = [string]$s.tenantSettingGroup; Enabled = $s.enabled
                CanSpecifySecurityGroups = $s.canSpecifySecurityGroups; EnabledSecurityGroups = (Join-A2C @($s.enabledSecurityGroups | ForEach-Object { $_.name }))
                ExcludedSecurityGroups = (Join-A2C @($s.excludedSecurityGroups | ForEach-Object { $_.name })); DelegateToCapacity = $s.delegateToCapacity
                DelegateToDomain = $s.delegateToDomain; DelegateToWorkspace = $s.delegateToWorkspace; Properties = $s.properties })
    }
    foreach ($c in $a.Capacities) {
        $rows['adminCapacities'].Add([ordered]@{ Capacity = [string]$c.displayName; CapacityId = [string]$c.id; Sku = [string]$c.sku; Region = [string]$c.region; State = [string]$c.state
                Admins = (Join-A2C $c.admins); TenantKeyId = [string]$c.tenantKeyId })
    }
    foreach ($x in $a.Apps) {
        $rows['adminApps'].Add([ordered]@{ App = [string]$x.name; AppId = [string]$x.id; Description = [string]$x.description; PublishedBy = [string]$x.publishedBy; LastUpdate = $x.lastUpdate
                WorkspaceId = [string]$x.workspaceId })
    }
    foreach ($pair in @(@('Published to web', $a.PublishedToWeb), @('Link for the whole organization', $a.OrgLinks))) {
        foreach ($x in $pair[1]) {
            $rows['widelyShared'].Add([ordered]@{ Kind = $pair[0]; ArtifactType = [string]$x.artifactType; Artifact = [string]$x.displayName; ArtifactId = [string]$x.artifactId
                    Sharer = [string](Get-A2CFirst $x @('sharer.emailAddress', 'sharer.displayName', 'sharer.identifier')); SharerType = [string]$x.sharer.principalType
                    ShareType = [string]$x.shareType; AccessRight = [string]$x.accessRight })
        }
    }
    foreach ($x in $a.Domains) {
        $rows['adminDomains'].Add([ordered]@{ Domain = [string]$x.displayName; DomainId = [string]$x.id; Description = [string]$x.description; ParentDomainId = [string]$x.parentDomainId
                ContributorsScope = [string]$x.contributorsScope })
    }
    if ($cs) { $files = @('adminWorkspaces', 'adminCapacities', 'adminDatasourceInstances') }   # the share-safe set
    foreach ($f in $files) { Add-A2CRows ('PowerBI/' + $f + '.csv') $rows[$f] }
    New-A2CPbiRepointingMap
    New-A2CPbiCapacityFootprint
    New-A2CPbiWorkspacePlan
    New-A2CPbiScopeSummary
    Write-A2CLog -Level Ok ('Tenant-wide: {0} workspaces, {1} items, {2} tenant settings -> PowerBI folder' -f $rows['adminWorkspaces'].Count, $rows['adminItems'].Count, $rows['tenantSettings'].Count)
}

function Get-A2CPbiAdminFacts {
    # One object per workspace of the admin list with what scope decisions need: personal or shared, active or not,
    # content (items found by the Scanner API), last change seen, administrators, data source types, gateway use.
    # Cached; empty when the admin report did not run or the account has no administrator role.
    $cache = $script:A2C.Data['PbiAdminFacts']
    if ($null -ne $cache) { return , $cache }
    $facts = [System.Collections.Generic.List[object]]::new()
    $script:A2C.Data['PbiAdminFacts'] = $facts
    $a = $script:A2C.Data['PbiAdmin']
    if ($null -eq $a -or -not (Test-A2CDataset 'PowerBIAdmin')) { return , $facts }
    $inst = @{}
    foreach ($d in $a.DatasourceInstances) { if ($d.datasourceId) { $inst[([string]$d.datasourceId).ToLowerInvariant()] = $d } }
    $use = @{}   # data source instance -> number of workspaces that use it
    $dsUse = @{}  # workspace -> data source instance -> items that use it
    $scan = @{}
    foreach ($sw in $a.Scan) {
        $n = 0
        $last = $null
        $dsIds = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($prop in $sw.PSObject.Properties) {
            # every array of objects with an id is a list of items: reports, datasets, ... and Fabric types
            if ($prop.Name -eq 'users' -or -not ($prop.Value -is [array])) { continue }
            foreach ($x in $prop.Value) {
                if ($null -eq $x) { continue }
                if ($null -eq $x.PSObject.Properties['id'] -and $null -eq $x.PSObject.Properties['objectId']) { continue }
                $n++
                foreach ($dv in @((Get-A2CFirst $x @('modifiedDateTime', 'modifiedDate', 'lastUpdatedDate')), (Get-A2CFirst $x @('createdDate', 'createdDateTime')))) {
                    $dt = ConvertTo-A2CUtc $dv
                    if ($null -ne $dt -and ($null -eq $last -or $dt -gt $last)) { $last = $dt }
                }
                foreach ($u in @($x.datasourceUsages)) {
                    if ($null -eq $u -or -not $u.datasourceInstanceId) { continue }
                    $dk = ([string]$u.datasourceInstanceId).ToLowerInvariant()
                    [void]$dsIds.Add($dk)
                    # re-pointing map: which items of this workspace use which data source
                    $wk = ([string]$sw.id).ToLowerInvariant()
                    if (-not $dsUse.ContainsKey($wk)) { $dsUse[$wk] = @{} }
                    if (-not $dsUse[$wk].ContainsKey($dk)) { $dsUse[$wk][$dk] = [System.Collections.Generic.List[object]]::new() }
                    $dsUse[$wk][$dk].Add([pscustomobject]@{ Type = $prop.Name; Id = [string](Get-A2CFirst $x @('id', 'objectId')); Name = [string](Get-A2CFirst $x @('name', 'displayName')) })
                }
            }
        }
        $types = [System.Collections.Generic.List[string]]::new()
        $gw = $false
        foreach ($k in $dsIds) {
            $use[$k] = [int]$use[$k] + 1
            $d = $inst[$k]
            if ($null -eq $d) { continue }
            if ($d.datasourceType) { $types.Add([string]$d.datasourceType) }
            if ($d.gatewayId) { $gw = $true }
        }
        $scan[([string]$sw.id).ToLowerInvariant()] = [pscustomobject]@{ Items = $n; Last = $last; Types = (Get-A2CUnique $types); Gateway = $gw; Datasources = $dsIds.Count }
    }
    $script:A2C.Data['PbiAdminInstanceUse'] = $use
    $script:A2C.Data['PbiAdminDatasourceUse'] = $dsUse
    $recent = $script:A2C.Data['PbiAdminRecent']
    $act = $null
    $actUsers = $null
    if (Test-A2CDataset 'PowerBIActivity') {
        $act = @{}
        $actUsers = @{}
        foreach ($e in @($script:A2C.Data['PbiActivity'])) {
            if ($null -eq $e -or -not $e.WorkspaceId) { continue }
            $k = ([string]$e.WorkspaceId).ToLowerInvariant()
            $act[$k] = [int]$act[$k] + 1
            if ($e.UserId) {
                if (-not $actUsers.ContainsKey($k)) { $actUsers[$k] = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase) }
                [void]$actUsers[$k].Add([string]$e.UserId)
            }
        }
    }
    $script:A2C.Data['PbiAdminActiveUsers'] = $actUsers
    foreach ($g in $a.Workspaces) {
        $id = [string]$g.id
        $k = $id.ToLowerInvariant()
        $type = [string]$g.type
        $personal = ($type -eq 'PersonalGroup' -or $type -eq 'Personal')
        $state = [string]$g.state
        $active = ($state -ne 'Deleted' -and $state -ne 'Removing')   # an Orphaned workspace still exists and may hold content
        $users = @($g.users | Where-Object { $null -ne $_ })
        $admins = @($users | Where-Object { $_.groupUserAccessRight -eq 'Admin' })
        # an admin whose account was deleted or disabled in Entra ID (Principals report) cannot act as an owner
        $gone = @($admins | Where-Object { (Get-A2CPbiPrincipalStatus $_.graphId $(if ($_.emailAddress) { $_.emailAddress } else { $_.identifier })) -eq 'Departed' }).Count
        $s = $scan[$k]
        $items = $null; $has = $null; $gw = $null; $dsn = $null; $types = ''; $last = $null
        if ($null -ne $s) { $items = $s.Items; $has = ($s.Items -gt 0); $gw = $s.Gateway; $dsn = $s.Datasources; $types = (Join-A2C $s.Types); $last = $s.Last }
        $kind = $(if ($personal) { 'Personal' } else { 'Shared' })
        $scope = $(if (-not $active) { 'Not active (' + $state + ')' } elseif ($null -eq $has) { $kind + ', not scanned' } elseif ($has) { $kind + ', with content' } else { $kind + ', empty' })
        $facts.Add([pscustomobject]@{
                Id = $id; Name = [string]$g.name; Type = $type; IsPersonal = $personal; State = $state; Active = $active; CapacityId = [string]$g.capacityId
                OnCapacity = ([bool]$g.isOnDedicatedCapacity -or [bool]$g.capacityId); Users = $users.Count
                Admins = (Join-A2C @($admins | ForEach-Object { $(if ($_.emailAddress) { $_.emailAddress } elseif ($_.displayName) { $_.displayName } else { $_.identifier }) }))
                NoAdmin = $(if ($personal) { $null } else { ($admins.Count - $gone -le 0 -or $state -eq 'Orphaned') }); AdminsDeparted = $gone
                Items = $items; HasContent = $has; LastItemChange = $last
                ChangedRecently = $(if ($null -ne $recent) { [bool]$recent[$k] } else { $null })
                RecentActivities = $(if ($null -ne $act) { [int]$act[$k] } else { $null })
                ActiveUsers = $(if ($null -ne $actUsers) { $(if ($actUsers.ContainsKey($k)) { $actUsers[$k].Count } else { 0 }) } else { $null })
                DatasourceTypes = $types; Datasources = $dsn; UsesGateway = $gw; Scope = $scope
                CreatedDate = (Get-A2CFirst $g @('createdDate', 'createdDateTime')); ModifiedDate = (Get-A2CFirst $g @('modifiedDate', 'modifiedDateTime', 'lastUpdatedDate'))
            })
    }
    return , $facts
}

function Get-A2CPbiEndpointHost {
    # Host name of a data source from its connection details: server (without port or tcp:), the host of a URL, or
    # the server of a UNC path. Nothing else: no database, path or file.
    param($Details)
    if ($null -eq $Details) { return '' }
    foreach ($key in @('server', 'url', 'path', 'domain', 'account', 'sharePointSiteUrl', 'extensionDataSourcePath')) {
        $v = [string]$Details.$key
        if (-not $v) { continue }
        $v = $v.Trim()
        if ($v -match '^[a-zA-Z][a-zA-Z0-9+.-]*://') { try { return ([uri]$v).Host } catch { continue } }
        if ($v.StartsWith('\\')) { return ($v.TrimStart('\') -split '\\')[0] }
        $v = ($v -split ';')[0] -replace '^(tcp|np|lpc):', ''
        $v = ($v -replace ',\d+$', '') -replace ':\d+$', ''
        return (($v -split '[\\/]')[0])
    }
    return ''
}

function New-A2CPbiRepointingMap {
    # repointingMap.csv: per workspace (also the ones this account cannot open) the data sources its items use, with
    # the host name, gateway, the number of items and models and their refresh schedules - what has to be re-pointed
    # at cutover. -ClientScope keeps the type and the host name only.
    if ($script:A2C.Out.ContainsKey('PowerBI/repointingMap.csv')) { return }
    $a = $script:A2C.Data['PbiAdmin']
    $use = $script:A2C.Data['PbiAdminDatasourceUse']
    if ($null -eq $a -or $null -eq $use) { return }
    $cs = [bool]$script:A2C.Options.ClientScope
    $inst = @{}
    foreach ($d in $a.DatasourceInstances) { if ($d.datasourceId) { $inst[([string]$d.datasourceId).ToLowerInvariant()] = $d } }
    $sched = @{}
    foreach ($x in $a.Refreshables) { if ($x.id) { $sched[([string]$x.id).ToLowerInvariant()] = $x } }
    $facts = @{}
    foreach ($f in (Get-A2CPbiAdminFacts)) { $facts[$f.Id.ToLowerInvariant()] = $f }
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($wk in @($use.Keys)) {
        $f = $facts[$wk]
        if ($null -eq $f -or -not $f.Active) { continue }
        $w = [pscustomobject]@{ Name = $f.Name; Id = $f.Id }
        foreach ($dk in @($use[$wk].Keys)) {
            $d = $inst[$dk]
            $items = @($use[$wk][$dk])
            $models = @($items | Where-Object { $_.Type -eq 'datasets' })
            $known = @($models | Where-Object { $sched.ContainsKey(([string]$_.Id).ToLowerInvariant()) })
            $on = @($known | Where-Object { $sched[([string]$_.Id).ToLowerInvariant()].refreshSchedule.enabled -eq $true })
            $times = Get-A2CUnique @($on | ForEach-Object {
                    $rs = $sched[([string]$_.Id).ToLowerInvariant()].refreshSchedule
                    ((Join-A2C $rs.days ',') + ' ' + (Join-A2C $rs.times ',')).Trim() })
            $schedKnown = $(if ($models.Count -eq 0) { '' } elseif (-not $f.OnCapacity) { 'No: shared capacity (no tenant-wide schedule API)' } elseif ($known.Count -eq $models.Count) { 'Yes' } else { 'Partly' })
            $cd = $d.connectionDetails
            $host1 = Get-A2CPbiEndpointHost $cd
            $target = [string](Get-A2CFirst $cd @('url', 'path', 'server', 'domain', 'account'))
            $ks = Get-A2CPbiTargetKind ([string]$d.datasourceType) $(if ($target) { $target } else { $host1 })
            $rows.Add((New-A2CPbiRow $w ([ordered]@{ Scope = $f.Scope; DatasourceType = [string]$d.datasourceType; Kind = $ks[0]; Host = $host1
                            Target = $(if ($cs) { '' } else { $target }); Database = $(if ($cs) { '' } else { [string]$cd.database }); GatewayId = [string]$d.gatewayId
                            UsesGateway = [bool]$d.gatewayId; Items = $items.Count; ItemNames = $(if ($cs) { '' } else { Join-A2C @($items | ForEach-Object { $_.Name }) })
                            Models = $models.Count; ScheduledModels = $(if ($schedKnown -like 'No*') { $null } else { $on.Count }); RefreshSchedules = (Join-A2C $times)
                            ScheduleKnown = $schedKnown })))
        }
    }
    Add-A2CRows 'PowerBI/repointingMap.csv' @($rows | Sort-Object { $_['Workspace'] }, { $_['Kind'] }, { $_['Host'] })
}

function New-A2CPbiCapacityFootprint {
    # capacityFootprint.csv: a load estimate per capacity from counts, refresh statistics and the activity log. True
    # capacity consumption (CU) has no API, so this is a sanity check of scale, not a sizing.
    if ($script:A2C.Out.ContainsKey('PowerBI/capacityFootprint.csv')) { return }
    $a = $script:A2C.Data['PbiAdmin']
    if ($null -eq $a) { return }
    $facts = Get-A2CPbiAdminFacts
    if ($facts.Count -eq 0) { return }
    $actUsers = $script:A2C.Data['PbiAdminActiveUsers']
    $caps = [ordered]@{}
    foreach ($c in $a.Capacities) { $caps[([string]$c.id).ToLowerInvariant()] = $c }
    $groups = [ordered]@{}
    foreach ($f in $facts) {
        if (-not $f.Active) { continue }
        $k = $(if ($f.CapacityId) { $f.CapacityId.ToLowerInvariant() } else { '' })
        if (-not $groups.Contains($k)) { $groups[$k] = [System.Collections.Generic.List[object]]::new() }
        $groups[$k].Add($f)
    }
    foreach ($k in @($caps.Keys)) { if (-not $groups.Contains($k)) { $groups[$k] = [System.Collections.Generic.List[object]]::new() } }
    $ref = @{}
    foreach ($x in $a.Refreshables) {
        $ck = ([string]$x.capacity.id).ToLowerInvariant()
        if (-not $ref.ContainsKey($ck)) { $ref[$ck] = [System.Collections.Generic.List[object]]::new() }
        $ref[$ck].Add($x)
    }
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($k in @($groups.Keys)) {
        $ws = $groups[$k]
        $c = $caps[$k]
        $rf = @($(if ($ref.ContainsKey($k)) { $ref[$k] } else { @() }))
        $users = $null
        if ($null -ne $actUsers) {
            $set = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
            foreach ($f in $ws) { $wk = $f.Id.ToLowerInvariant(); if ($actUsers.ContainsKey($wk)) { foreach ($u in $actUsers[$wk]) { [void]$set.Add($u) } } }
            $users = $set.Count
        }
        $durs = @($rf | Where-Object { $null -ne $_.averageDuration } | ForEach-Object { [double]$_.averageDuration })
        $rows.Add([ordered]@{
                Capacity = $(if ($k -eq '') { '(shared capacity)' } elseif ($null -ne $c) { [string]$c.displayName } else { '(capacity not visible)' }); CapacityId = $(if ($k) { $k } else { '' })
                Sku = [string]$c.sku; Region = [string]$c.region; State = [string]$c.state; Workspaces = $ws.Count
                SharedWithContent = @($ws | Where-Object { -not $_.IsPersonal -and $_.HasContent -eq $true }).Count
                Items = (@($ws | ForEach-Object { [int]$_.Items }) | Measure-Object -Sum).Sum
                ScheduledModels = $(if ($k -eq '') { $null } else { @($rf | Where-Object { $_.refreshSchedule.enabled -eq $true }).Count })
                RefreshesPerDay = $(if ($k -eq '') { $null } else { [math]::Round((@($rf | ForEach-Object { [double]$_.refreshesPerDay }) | Measure-Object -Sum).Sum, 1) })
                AvgRefreshSec = $(if ($durs.Count -gt 0) { [math]::Round(($durs | Measure-Object -Average).Average, 0) } else { $null })
                ActiveUsers = $users
                Note = 'estimate: counts, refresh statistics and the activity log; no capacity consumption (CU) data exists in the APIs'
            })
    }
    Add-A2CRows 'PowerBI/capacityFootprint.csv' $rows
}

function New-A2CPbiScopeSummary {
    # workspaceScopeSummary.csv: how many of the tenant's workspaces are real migration scope, in one glance.
    if ($script:A2C.Out.ContainsKey('PowerBI/workspaceScopeSummary.csv')) { return }
    $f = Get-A2CPbiAdminFacts
    if ($f.Count -eq 0) { return }
    $act = @($f | Where-Object { $_.Active })
    $pers = @($act | Where-Object { $_.IsPersonal })
    $shared = @($act | Where-Object { -not $_.IsPersonal })
    $content = @($shared | Where-Object { $_.HasContent -eq $true })
    $rows = [System.Collections.Generic.List[object]]::new()
    $add = { param($Measure, $Count, $Note) $rows.Add([ordered]@{ Measure = $Measure; Count = $Count; Note = $Note }) }
    & $add 'Workspaces in the tenant' $f.Count 'all workspaces the admin API returns, also deleted and orphaned ones'
    & $add 'Not active' @($f | Where-Object { -not $_.Active }).Count 'deleted, or being removed'
    & $add 'Personal ("My workspace")' $pers.Count 'created automatically for every user'
    & $add 'Personal with content' @($pers | Where-Object { $_.HasContent -eq $true }).Count 'items in someone''s My workspace: decide per owner'
    & $add 'Personal, empty' @($pers | Where-Object { $_.HasContent -eq $false }).Count 'nothing to migrate'
    & $add 'Shared' $shared.Count ''
    & $add 'Shared, empty' @($shared | Where-Object { $_.HasContent -eq $false }).Count 'nothing to migrate'
    & $add 'Shared with content' $content.Count 'the migration scope'
    & $add 'Shared with content, on a capacity' @($content | Where-Object { $_.OnCapacity }).Count 'Premium / Fabric capacity: the target tenant needs one too'
    & $add 'Shared with content, using a gateway' @($content | Where-Object { $_.UsesGateway -eq $true }).Count 'gateways and their connections are recreated by hand'
    & $add 'Shared with content, without an admin' @($content | Where-Object { $_.NoAdmin -eq $true }).Count 'nobody to decide for them: find an owner first'
    $recentKnown = @($f | Where-Object { $null -ne $_.ChangedRecently }).Count -gt 0
    & $add 'Shared with content, changed in the last 30 days' $(if ($recentKnown) { @($content | Where-Object { $_.ChangedRecently -eq $true }).Count } else { $null }) $(if ($recentKnown) { 'in active use' } else { 'not read' })
    if ($null -ne $content[0] -and $null -ne $content[0].RecentActivities) {
        & $add 'Shared with content, with activity in the activity log' @($content | Where-Object { $_.RecentActivities -gt 0 }).Count ('last ' + $script:A2C.Options.PbiActivityDays + ' days')
    }
    & $add 'Shared, not scanned' @($shared | Where-Object { $null -eq $_.HasContent }).Count 'the Scanner API returned nothing for them: see coverage.csv'
    $git = $null; $gitNote = 'not available from the admin APIs'
    if ($script:A2C.Out.ContainsKey('PowerBI/workspaceGit.csv')) {
        $ids = @{}; foreach ($c in $content) { $ids[$c.Id.ToLowerInvariant()] = $true }
        $git = @((Get-A2CRawRows 'PowerBI/workspaceGit.csv') | Where-Object { [string]$_['State'] -like 'Connected*' -and $ids.ContainsKey(([string]$_['WorkspaceId']).ToLowerInvariant()) }).Count
        $gitNote = 'only workspaces this account is a member of'
    }
    & $add 'Shared with content, connected to Git' $git $gitNote
    $caps = $script:A2C.Data['PbiAdmin'].Capacities
    & $add 'Capacities' @($caps).Count ''
    if ($script:A2C.Out.ContainsKey('PowerBI/repointingMap.csv')) {
        $ids = @{}; foreach ($c in $content) { $ids[$c.Id.ToLowerInvariant()] = $true }
        $map = @((Get-A2CRawRows 'PowerBI/repointingMap.csv') | Where-Object { $ids.ContainsKey(([string]$_['WorkspaceId']).ToLowerInvariant()) })
        & $add 'Data source hosts used by shared workspaces with content' (Get-A2CUnique @($map | ForEach-Object { [string]$_['Kind'] + '|' + [string]$_['Host'] })).Count 'what has to be re-pointed at cutover: repointingMap.csv'
        & $add 'Data sources through a gateway (shared workspaces with content)' (Get-A2CUnique @($map | Where-Object { $_['UsesGateway'] -eq $true } | ForEach-Object { [string]$_['Host'] })).Count 'gateways are recreated in the target tenant'
    }
    $pr = $script:A2C.Data['PbiPrincipals']
    if ($null -ne $pr -and $pr.Count -gt 0) {
        $people = @($pr | Where-Object { $_.Type -eq 'User' -or $_.Type -eq 'Guest' })
        & $add 'Principals: users' @($pr | Where-Object { $_.Type -eq 'User' }).Count 'distinct people with a role on a workspace, item, capacity or gateway'
        & $add 'Principals: guests' @($pr | Where-Object { $_.Type -eq 'Guest' }).Count 'B2B guests: they need a new invitation in the target tenant'
        & $add 'Principals: groups' @($pr | Where-Object { $_.Type -eq 'Group' }).Count 'groups are recreated or synced in the target tenant'
        & $add 'Principals: service principals and apps' @($pr | Where-Object { $_.Type -eq 'Service principal' -or $_.Type -eq 'App' }).Count 'app registrations and their secrets are recreated'
        & $add 'Principals: managed identities' @($pr | Where-Object { $_.Type -eq 'Managed identity' }).Count ''
        & $add 'Principals: people whose account was deleted or disabled' @($people | Where-Object { $_.Departed -eq 'Yes' }).Count 'Entra ID lookup'
    }
    $au = $script:A2C.Data['PbiAdminActiveUsers']
    if ($null -ne $au) {
        $all = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($set in $au.Values) { foreach ($u in $set) { [void]$all.Add($u) } }
        & $add 'Active users in the activity log' $all.Count ('estimate: distinct users with an event in a workspace, last ' + $script:A2C.Options.PbiActivityDays + ' days')
    }
    if ($script:A2C.Out.ContainsKey('PowerBI/workspacePlan.csv')) {
        $plan = Get-A2CRawRows 'PowerBI/workspacePlan.csv'
        $listed = @($plan | Where-Object { $_['ListedWorkspace'] })
        & $add 'On the workspace list' $listed.Count ''
        & $add 'On the list, not found in the tenant' @($listed | Where-Object { $_['Status'] -like 'Not *' }).Count 'misspelled beyond recognition, renamed or deleted'
        & $add 'On the list, but empty, personal or not active' @($listed | Where-Object { [string]$_['Check'] -like 'On the list but*' }).Count 'may not need a migration'
        & $add 'Shared with content, not on the list' @($plan | Where-Object { [string]$_['Check'] -eq 'Shared with content, not on the list' }).Count 'overlooked?'
    }
    Add-A2CRows 'PowerBI/workspaceScopeSummary.csv' $rows
    Write-A2CLog -Level Ok ('Scope: {0} workspaces, {1} shared with content, {2} empty personal' -f $f.Count, $content.Count, @($pers | Where-Object { $_.HasContent -eq $false }).Count)
}

function New-A2CReportPowerBIActivity {
    if (-not (Test-A2CDataset 'PowerBIActivity')) { return }
    if ($script:A2C.Options.ClientScope) { return }   # counted per workspace and capacity only
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($e in @($script:A2C.Data['PbiActivity'])) {
        if ($null -eq $e) { continue }
        $rows.Add([ordered]@{ CreationTime = $e.CreationTime; UserId = [string]$e.UserId; Activity = [string]$e.Activity; Operation = [string]$e.Operation
                Workspace = [string](Get-A2CFirst $e @('WorkSpaceName', 'WorkspaceName')); WorkspaceId = [string]$e.WorkspaceId
                ItemName = [string](Get-A2CFirst $e @('ItemName', 'ArtifactName')); ItemType = [string](Get-A2CFirst $e @('ItemType', 'ArtifactKind', 'ObjectType'))
                ObjectId = [string]$e.ObjectId; ModelName = [string]$e.DatasetName; ReportName = [string]$e.ReportName; ClientIP = [string]$e.ClientIP; UserAgent = [string]$e.UserAgent
                IsSuccess = $e.IsSuccess; RequestId = [string]$e.RequestId; Id = [string]$e.Id; Properties = $e })
    }
    Add-A2CRows 'PowerBI/activityEvents.csv' $rows
}

function New-A2CReportFabricCode { New-A2CReportPowerBIExport }

function New-A2CReportPowerBIExport {
    if ($script:A2C.Out.ContainsKey('PowerBI/exportIndex.csv')) { return }   # FabricCode and PowerBIExport share the index
    $rows = [System.Collections.Generic.List[object]]::new()
    foreach ($x in @($script:A2C.Data['PbiExportIndex'])) {
        if ($null -eq $x) { continue }
        $rows.Add((New-A2CPbiRow $x.Ws ([ordered]@{ ItemType = $x.ItemType; Item = $x.Item; ItemId = $x.ItemId; Kind = $x.Kind; Status = $x.Status; Files = $x.Files; Bytes = $x.Bytes
                        Path = $x.Path; Detail = $x.Detail })))
    }
    Add-A2CRows 'PowerBI/exportIndex.csv' $rows
}
# ---- end Power BI ----

#endregion

#region ============================== 8. FINDINGS =====================================================

Register-A2CSchema 'findings.csv' 'Id Severity Category Title:note Count:num AffectedFile Recommendation:note LearnMore:url Detail:note'

$script:A2CFindingRules = [System.Collections.Generic.List[object]]::new()
$script:A2CSeverityOrder = @{ High = 0; Medium = 1; Low = 2; Info = 3 }

function Register-A2CFinding {
    # Evaluate returns $null/0 (nothing found) or @{ Count = n; Detail = '...' }.
    param([string] $Id, [ValidateSet('High', 'Medium', 'Low', 'Info')] [string] $Severity,
        [ValidateSet('Migration', 'Retirement', 'Cost', 'Security', 'Operations', 'Governance')] [string] $Category,
        [string] $Title, [string[]] $Needs, [string] $File, [string] $Recommendation, [string] $Url, [scriptblock] $Evaluate)
    $script:A2CFindingRules.Add([pscustomobject]@{ Id = $Id; Severity = $Severity; Category = $Category; Title = $Title; Needs = $Needs
            File = $File; Recommendation = $Recommendation; Url = $Url; Evaluate = $Evaluate })
}

function Invoke-A2CFindings {
    $notEvaluated = 0
    foreach ($rule in $script:A2CFindingRules) {
        $ok = $true
        foreach ($n in $rule.Needs) {
            if ($n -like 'Report:*') { if (-not (Test-A2CSelected $n.Substring(7))) { $ok = $false; break } }
            elseif (-not (Test-A2CDataset $n)) { $ok = $false; break }
        }
        if (-not $ok) { $notEvaluated++; continue }
        try {
            $res = & $rule.Evaluate
            if ($null -eq $res) { continue }
            $count = [int]$res.Count
            if ($count -le 0) { continue }
            $title = $rule.Title
            $script:A2C.Findings.Add([ordered]@{
                    Id = $rule.Id; Severity = $rule.Severity; Category = $rule.Category; Title = $title; Count = $count
                    AffectedFile = $rule.File; Recommendation = $rule.Recommendation; LearnMore = $rule.Url; Detail = [string]$res.Detail
                })
        }
        catch { Write-A2CLog -Level Warn ('Finding {0} could not be evaluated: {1}' -f $rule.Id, $_.Exception.Message) }
    }
    $sorted = @($script:A2C.Findings | Sort-Object @{ Expression = { $script:A2CSeverityOrder[$_.Severity] } }, @{ Expression = { $_.Category } }, @{ Expression = { -1 * $_.Count } })
    $script:A2C.Findings.Clear()
    foreach ($f in $sorted) { $script:A2C.Findings.Add($f) }
    $high = @($sorted | Where-Object { $_.Severity -eq 'High' }).Count
    Write-A2CLog -Level Ok ('{0} findings ({1} high); {2} rules not evaluated because their data was not selected' -f $sorted.Count, $high, $notEvaluated)
    $script:A2C.Data['FindingsNotEvaluated'] = $notEvaluated
}

function Get-A2CTopText {
    # "a (3), b (2), ..." from a list of strings
    param($Items, [int] $Top = 5)
    $h = @{}
    foreach ($i in $Items) { if ($i) { $h[[string]$i] = [int]$h[[string]$i] + 1 } }
    $parts = @($h.GetEnumerator() | Sort-Object @{ Expression = { -1 * $_.Value } }, @{ Expression = { $_.Key } } | Select-Object -First $Top | ForEach-Object { '{0} ({1})' -f $_.Key, $_.Value })
    $t = [string]::Join(', ', $parts)
    if ($h.Count -gt $Top) { $t += ', ...' }
    return $t
}

function New-A2CFindingResult {
    param([int] $Count, [string] $Detail = '')
    if ($Count -le 0) { return $null }
    return [pscustomobject]@{ Count = $Count; Detail = $Detail }
}

# ---- Power BI ----
Register-A2CFinding -Id 'PBI-01' -Severity High -Category Operations -Title 'Semantic models whose last refresh failed' -Needs @('PowerBI') -File 'PowerBI/semanticModels.csv' `
    -Recommendation 'Open the refresh history of the model (refreshHistory.csv has the error) and fix the credentials, gateway or query.' `
    -Url 'https://learn.microsoft.com/power-bi/connect-data/refresh-troubleshooting-refresh-scenarios' -Evaluate {
    $bad = @((Get-A2CRows 'PowerBI/semanticModels.csv') | Where-Object { $_['LastRefreshStatus'] -eq 'Failed' })
    New-A2CFindingResult $bad.Count (Get-A2CTopText @($bad | ForEach-Object { $_['Workspace'] + ' / ' + $_['Model'] }))
}

Register-A2CFinding -Id 'PBI-02' -Severity Low -Category Operations -Title 'Refreshable semantic models without an enabled refresh schedule' -Needs @('PowerBI') -File 'PowerBI/semanticModels.csv' `
    -Recommendation 'Check whether the data is refreshed another way (pipeline, API, Direct Lake, manually); otherwise the reports show old data.' `
    -Url 'https://learn.microsoft.com/power-bi/connect-data/refresh-scheduled-refresh' -Evaluate {
    $bad = @((Get-A2CRows 'PowerBI/semanticModels.csv') | Where-Object { $_['IsRefreshable'] -eq $true -and $_['ScheduleEnabled'] -eq $false })
    New-A2CFindingResult $bad.Count (Get-A2CTopText @($bad | ForEach-Object { $_['Workspace'] }))
}

Register-A2CFinding -Id 'PBI-03' -Severity Low -Category Governance -Title 'Semantic models that no report in the same workspace uses' -Needs @('PowerBI') -File 'PowerBI/semanticModels.csv' `
    -Recommendation 'Check for use from other workspaces or Excel (lineage view) and remove models nobody uses before a migration.' `
    -Url 'https://learn.microsoft.com/power-bi/collaborate-share/service-data-lineage' -Evaluate {
    $used = @{}
    foreach ($r in (Get-A2CRows 'PowerBI/reports.csv')) { if ($r['ModelId']) { $used[[string]$r['ModelId']] = $true } }
    $bad = @((Get-A2CRows 'PowerBI/semanticModels.csv') | Where-Object { -not $used.ContainsKey([string]$_['ModelId']) })
    New-A2CFindingResult $bad.Count (Get-A2CTopText @($bad | ForEach-Object { $_['Workspace'] }))
}

Register-A2CFinding -Id 'PBI-04' -Severity Medium -Category Governance -Title 'Workspaces with a single administrator' -Needs @('PbiWorkspaces') -File 'PowerBI/workspaces.csv' `
    -Recommendation 'Add a second administrator, preferably a security group, so the workspace is not orphaned when that person leaves.' `
    -Url 'https://learn.microsoft.com/fabric/fundamentals/roles-workspaces' -Evaluate {
    $bad = @((Get-A2CRows 'PowerBI/workspaces.csv') | Where-Object { $_['Admins'] -and @(([string]$_['Admins']) -split '; ').Count -eq 1 })
    New-A2CFindingResult $bad.Count (Get-A2CTopText @($bad | ForEach-Object { $_['Workspace'] }) 8)
}

Register-A2CFinding -Id 'PBI-05' -Severity Medium -Category Governance -Title 'Reports, models or dataflows owned by somebody who is not a member of the workspace' -Needs @('PowerBI') -File 'PowerBI/semanticModels.csv' `
    -Recommendation 'The owner probably left or changed role: take over the item with an account that stays, or refreshes and gateway connections stop working.' `
    -Url 'https://learn.microsoft.com/power-bi/connect-data/service-datasets-take-over' -Evaluate {
    $members = @{}
    foreach ($u in (Get-A2CRows 'PowerBI/workspaceUsers.csv')) {
        foreach ($v in @($u['Identifier'], $u['EmailAddress'])) { if ($v) { $members[[string]$u['WorkspaceId'] + '|' + ([string]$v).ToLowerInvariant()] = $true } }
        if ($u['PrincipalType'] -eq 'Group') { $members[[string]$u['WorkspaceId'] + '|*group'] = $true }
    }
    $known = @{}
    foreach ($k in $members.Keys) { $known[$k.Split('|')[0]] = $true }
    $bad = [System.Collections.Generic.List[string]]::new()
    foreach ($f in @('PowerBI/semanticModels.csv', 'PowerBI/dataflows.csv')) {
        foreach ($r in (Get-A2CRows $f)) {
            $ws = [string]$r['WorkspaceId']
            $owner = ([string]$r['ConfiguredBy']).ToLowerInvariant()
            # only where the members are known and none of them is a group (membership of a group is not expanded)
            if (-not $owner -or -not $known.ContainsKey($ws) -or $members.ContainsKey($ws + '|*group')) { continue }
            if (-not $members.ContainsKey($ws + '|' + $owner)) { $bad.Add($r['Workspace'] + ': ' + $r['ConfiguredBy']) }
        }
    }
    New-A2CFindingResult $bad.Count (Get-A2CTopText $bad)
}

Register-A2CFinding -Id 'PBI-06' -Severity Info -Category Migration -Title 'Workspaces that are not on a capacity (shared / Pro)' -Needs @('PbiWorkspaces') -File 'PowerBI/workspaces.csv' `
    -Recommendation 'Their content needs a Pro or PPU licence per user and cannot hold Fabric items; decide which capacity they get in the target tenant.' `
    -Url 'https://learn.microsoft.com/fabric/enterprise/licenses' -Evaluate {
    $bad = @((Get-A2CRows 'PowerBI/workspaces.csv') | Where-Object { -not $_['CapacityId'] })
    New-A2CFindingResult $bad.Count (Get-A2CTopText @($bad | ForEach-Object { $_['Workspace'] }) 8)
}

Register-A2CFinding -Id 'PBI-07' -Severity Low -Category Governance -Title 'Workspaces on a capacity with content but without a Git connection' -Needs @('Fabric') -File 'PowerBI/workspaceGit.csv' `
    -Recommendation 'Connect the workspace to a Git repository: the definitions are versioned and can be deployed to another tenant from source.' `
    -Url 'https://learn.microsoft.com/fabric/cicd/git-integration/intro-to-git-integration' -Evaluate {
    $bad = @((Get-A2CRows 'PowerBI/workspaces.csv') | Where-Object { $_['CapacityId'] -and [int]$_['FabricItems'] -gt 0 -and $_['GitState'] -notlike 'Connected*' })
    New-A2CFindingResult $bad.Count (Get-A2CTopText @($bad | ForEach-Object { $_['Workspace'] }) 8)
}

Register-A2CFinding -Id 'PBI-08' -Severity High -Category Security -Title 'Reports published to the web (public, no sign-in)' -Needs @('PowerBIAdmin') -File 'PowerBI/widelyShared.csv' `
    -Recommendation 'Anybody with the link can see the data. Remove the embed codes that are not meant to be public and limit the tenant setting to a group.' `
    -Url 'https://learn.microsoft.com/power-bi/collaborate-share/service-publish-to-web' -Evaluate {
    $bad = @((Get-A2CRows 'PowerBI/widelyShared.csv') | Where-Object { $_['Kind'] -eq 'Published to web' })
    New-A2CFindingResult $bad.Count (Get-A2CTopText @($bad | ForEach-Object { $_['Artifact'] }))
}

Register-A2CFinding -Id 'PBI-09' -Severity Medium -Category Security -Title 'Items shared with a link for the whole organization' -Needs @('PowerBIAdmin') -File 'PowerBI/widelyShared.csv' `
    -Recommendation 'Everybody in the tenant who gets the link can open the item. Replace the links with access for specific people or groups.' `
    -Url 'https://learn.microsoft.com/power-bi/collaborate-share/service-share-dashboards' -Evaluate {
    $bad = @((Get-A2CRows 'PowerBI/widelyShared.csv') | Where-Object { $_['Kind'] -ne 'Published to web' })
    New-A2CFindingResult $bad.Count (Get-A2CTopText @($bad | ForEach-Object { $_['Artifact'] }))
}

Register-A2CFinding -Id 'PBI-10' -Severity Info -Category Governance -Title 'Workspaces that this account could read only partly' -Needs @('PbiWorkspaces') -File 'PowerBI/workspaces.csv' `
    -Recommendation 'The inventory is incomplete there (see the Problems columns and coverage.csv). Run it with a Fabric administrator or a workspace Admin/Contributor account.' `
    -Url 'https://learn.microsoft.com/fabric/fundamentals/roles-workspaces' -Evaluate {
    $bad = @((Get-A2CRows 'PowerBI/workspaces.csv') | Where-Object { $_['PowerBIStatus'] -eq 'Partial' -or $_['FabricStatus'] -eq 'Partial' })
    New-A2CFindingResult $bad.Count (Get-A2CTopText @($bad | ForEach-Object { $_['Workspace'] }) 8)
}
Register-A2CFinding -Id 'PBI-11' -Severity Medium -Category Migration -Title 'Workspaces on the workspace list that were not found' -Needs @() -File 'PowerBI/workspacePlan.csv' `
    -Recommendation 'Run as a member: nothing of them is in this inventory; get added (Member or higher) or run as a Fabric administrator. Run as an administrator: the name is wrong, or the workspace was renamed or deleted.' `
    -Url 'https://learn.microsoft.com/fabric/fundamentals/give-access-workspaces' -Evaluate {
    $bad = @((Get-A2CRows 'PowerBI/workspacePlan.csv') | Where-Object { $_['Status'] -eq 'Not visible to this account' -or $_['Status'] -eq 'Not in the tenant' })
    New-A2CFindingResult $bad.Count (Get-A2CTopText @($bad | ForEach-Object { $_['ListedWorkspace'] }) 10)
}

Register-A2CFinding -Id 'PBI-12' -Severity Medium -Category Operations -Title 'Pipelines and notebooks whose last run failed' -Needs @('Fabric') -File 'PowerBI/fabricItems.csv' `
    -Recommendation 'Check the run in the monitoring hub (itemRuns.csv has the failure reason); a pipeline that fails before a migration will fail after it.' `
    -Url 'https://learn.microsoft.com/fabric/admin/monitoring-hub' -Evaluate {
    $bad = @((Get-A2CRows 'PowerBI/fabricItems.csv') | Where-Object { $_['LastRunStatus'] -eq 'Failed' })
    New-A2CFindingResult $bad.Count (Get-A2CTopText @($bad | ForEach-Object { $_['Workspace'] + ' / ' + $_['Item'] }))
}

Register-A2CFinding -Id 'PBI-13' -Severity Info -Category Migration -Title 'External targets that item definitions and connections point to' -Needs @('ItemDefinitions') -File 'PowerBI/externalDependencies.csv' `
    -Recommendation 'These addresses (SharePoint sites, SQL servers, storage accounts) and the Fabric addresses of this tenant are written into dataflows, models, shortcuts and notebooks; plan who changes them in the target tenant.' `
    -Url 'https://learn.microsoft.com/fabric/data-factory/connector-overview' -Evaluate {
    $targets = @{}
    foreach ($r in (Get-A2CRows 'PowerBI/externalDependencies.csv')) { $targets[[string]$r['Kind'] + '|' + [string]$r['Target']] = [string]$r['Kind'] }
    New-A2CFindingResult $targets.Count (Get-A2CTopText @($targets.Values) 8)
}

Register-A2CFinding -Id 'PBI-14' -Severity Low -Category Governance -Title 'Items that exist in DEV or TEST but not in PROD' -Needs @('Fabric') -File 'PowerBI/environmentMatrix.csv' `
    -Recommendation 'Work in progress, leftovers or content that was never promoted: decide per item whether it moves to the target tenant.' `
    -Url 'https://learn.microsoft.com/fabric/cicd/deployment-pipelines/intro-to-deployment-pipelines' -Evaluate {
    $bad = @((Get-A2CRows 'PowerBI/environmentMatrix.csv') | Where-Object { $_['MissingIn'] -match 'PROD' })
    New-A2CFindingResult $bad.Count (Get-A2CTopText @($bad | ForEach-Object { $_['Solution'] }) 8)
}
Register-A2CFinding -Id 'PBI-15' -Severity Medium -Category Migration -Title 'Shared workspaces with content that are not on the workspace list' -Needs @() -File 'PowerBI/workspacePlan.csv' `
    -Recommendation 'Overlooked, or deliberately left out? Decide for each one before the migration plan is final; the Items and ChangedLast30Days columns show which are in use.' `
    -Url 'https://learn.microsoft.com/fabric/admin/portal-workspaces' -Evaluate {
    $bad = @((Get-A2CRows 'PowerBI/workspacePlan.csv') | Where-Object { $_['Check'] -eq 'Shared with content, not on the list' })
    New-A2CFindingResult $bad.Count (Get-A2CTopText @($bad | ForEach-Object { $_['Workspace'] }) 10)
}

Register-A2CFinding -Id 'PBI-16' -Severity Low -Category Migration -Title 'Workspaces on the list that are empty, personal or not active' -Needs @() -File 'PowerBI/workspacePlan.csv' `
    -Recommendation 'These may need no migration at all. Confirm with the owners and take them off the list.' `
    -Url 'https://learn.microsoft.com/fabric/admin/portal-workspaces' -Evaluate {
    $bad = @((Get-A2CRows 'PowerBI/workspacePlan.csv') | Where-Object { [string]$_['Check'] -like 'On the list but*' })
    New-A2CFindingResult $bad.Count (Get-A2CTopText @($bad | ForEach-Object { $_['ListedWorkspace'] + ': ' + ([string]$_['Check']).Replace('On the list but ', '') }) 10)
}

Register-A2CFinding -Id 'PBI-17' -Severity Medium -Category Governance -Title 'Shared workspaces with content and no administrator' -Needs @('PowerBIAdmin') -File 'PowerBI/adminWorkspaces.csv' `
    -Recommendation 'Orphaned workspaces have nobody who can decide what happens to them. Assign an administrator (Fabric admin portal, Workspaces) before planning the move.' `
    -Url 'https://learn.microsoft.com/fabric/admin/portal-workspaces' -Evaluate {
    $bad = @((Get-A2CRows 'PowerBI/adminWorkspaces.csv') | Where-Object { $_['NoAdmin'] -eq $true -and $_['HasContent'] -eq $true })
    New-A2CFindingResult $bad.Count (Get-A2CTopText @($bad | ForEach-Object { $_['Workspace'] }) 10)
}

Register-A2CFinding -Id 'PBI-18' -Severity Info -Category Migration -Title 'Migration scope: shared workspaces with content' -Needs @('PowerBIAdmin') -File 'PowerBI/workspaceScopeSummary.csv' `
    -Recommendation 'The rest of the tenant is personal, empty or not active. workspaceScopeSummary.csv has the breakdown, adminWorkspaces.csv the Scope column per workspace.' `
    -Url 'https://learn.microsoft.com/fabric/admin/portal-workspaces' -Evaluate {
    $sum = @{}
    foreach ($r in (Get-A2CRows 'PowerBI/workspaceScopeSummary.csv')) { $sum[[string]$r['Measure']] = $r['Count'] }
    if (-not $sum.ContainsKey('Shared with content')) { return $null }
    New-A2CFindingResult ([int]$sum['Shared with content']) ('{0} workspaces: {1} personal ({2} empty), {3} shared ({4} empty), {5} not active' -f $sum['Workspaces in the tenant'],
        $sum['Personal ("My workspace")'], $sum['Personal, empty'], $sum['Shared'], $sum['Shared, empty'], $sum['Not active'])
}
Register-A2CFinding -Id 'PBI-19' -Severity Info -Category Migration -Title 'Why the tenant has so many personal workspaces' -Needs @('PowerBIAdmin') -File 'PowerBI/workspaceScopeSummary.csv' `
    -Recommendation 'Power BI creates a "My workspace" for every user the first time they open the service. Most stay empty; the large total is not migration scope. Move only the personal workspaces with content, after asking their owners.' `
    -Url 'https://learn.microsoft.com/fabric/admin/portal-workspaces#govern-my-workspaces' -Evaluate {
    $sum = @{}
    foreach ($r in (Get-A2CRows 'PowerBI/workspaceScopeSummary.csv')) { $sum[[string]$r['Measure']] = $r['Count'] }
    $pers = [int]$sum['Personal ("My workspace")']
    if ($pers -le 0) { return $null }
    $empty = [int]$sum['Personal, empty']
    $detail = '{0} personal workspaces, {1} of them empty ({2:N0}%), {3} with content' -f $pers, $empty, (100.0 * $empty / $pers), $sum['Personal with content']
    if ($sum.ContainsKey('Principals: users')) { $detail += ('; {0} users and {1} guests have a role somewhere' -f $sum['Principals: users'], $sum['Principals: guests']) }
    $settings = @((Get-A2CRows 'PowerBI/tenantSettings.csv') | Where-Object { [string]$_['SettingName'] + ' ' + [string]$_['Title'] -match '(?i)personal|my ?workspace' })
    if ($settings.Count -gt 0) { $detail += '; tenant settings: ' + [string]::Join(', ', @($settings | ForEach-Object { '{0} = {1}' -f $_['Title'], $_['Enabled'] })) }
    New-A2CFindingResult $pers $detail
}

Register-A2CFinding -Id 'PBI-20' -Severity Medium -Category Governance -Title 'Workspace or capacity admins whose account was deleted or disabled' -Needs @('Principals') -File 'PowerBI/principals.csv' `
    -Recommendation 'Their workspaces may have nobody left who can decide or act. Replace them with an active admin, preferably a group, before the migration.' `
    -Url 'https://learn.microsoft.com/fabric/admin/portal-workspaces' -Evaluate {
    $bad = @((Get-A2CRows 'PowerBI/principals.csv') | Where-Object { $_['Departed'] -eq 'Yes' -and [string]$_['Roles'] -match 'Workspace Admin|Capacity admin' })
    New-A2CFindingResult $bad.Count (Get-A2CTopText @($bad | ForEach-Object { $(if ($_['Identifier']) { $_['Identifier'] } else { $_['ObjectId'] }) }) 10)
}

Register-A2CFinding -Id 'PBI-21' -Severity Info -Category Migration -Title 'Data source hosts to re-point at cutover' -Needs @('PowerBIAdmin') -File 'PowerBI/repointingMap.csv' `
    -Recommendation 'Every host in repointingMap.csv is referenced by models, dataflows or reports. Plan for each one: does it move, keep its address, or change, and who updates the connections and gateways.' `
    -Url 'https://learn.microsoft.com/power-bi/connect-data/service-gateway-data-sources' -Evaluate {
    $h = @{}
    foreach ($r in (Get-A2CRows 'PowerBI/repointingMap.csv')) { if ($r['Host']) { $h[[string]$r['Kind'] + '|' + [string]$r['Host']] = [string]$r['Kind'] } }
    New-A2CFindingResult $h.Count (Get-A2CTopText @($h.Values) 8)
}
# ---- end Power BI ----

#endregion

#region ============================== 9. OUTPUT =======================================================

Register-A2CSchema 'coverage.csv' 'Area:note Status Items:num Detail:note'

function Test-A2CModuleAvailable {
    param([string] $Name)
    return [bool](Get-Module -ListAvailable -Name $Name)
}

function Get-A2COutputRows {
    # Rows as written to files: masked copies under -Mask (cached), originals otherwise.
    param([string] $File)
    if ($script:A2C.Options.Mask) {
        if ($null -eq $script:A2C.Data['MaskedOut']) { $script:A2C.Data['MaskedOut'] = @{} }
        $c = $script:A2C.Data['MaskedOut'][$File]
        if ($null -eq $c) { $c = Protect-A2CRows -File $File -Rows (Get-A2CRawRows $File); $script:A2C.Data['MaskedOut'][$File] = $c }
        return , $c
    }
    return , (Get-A2CRawRows $File)
}

function Save-A2CExtraFiles {
    Add-A2CRows 'findings.csv' $script:A2C.Findings
    Save-A2CCsv -File 'findings.csv'
    Add-A2CRows 'coverage.csv' $script:A2C.Coverage
    Save-A2CCsv -File 'coverage.csv'
    $tl = $script:A2C.Data['TagList']
    if ($null -ne $tl) {
        if ($script:A2C.Options.Mask) { $tl = Protect-A2CTagList $tl }
        [void](Save-A2CText -File 'tagList.json' -Text (ConvertTo-Json -InputObject $tl -Depth 5) -NoBom)
        $script:A2C.Files.Add([pscustomobject]@{ File = 'tagList.json'; Rows = $tl.Count; Path = [IO.Path]::Combine($script:A2C.OutDir, 'tagList.json') })
    }
}

function Save-A2CRunInfo {
    $o = $script:A2C.Options
    $fin = [DateTime]::UtcNow
    $findingCounts = [ordered]@{}
    foreach ($sev in @('High', 'Medium', 'Low', 'Info')) { $findingCounts[$sev] = @($script:A2C.Findings | Where-Object { $_.Severity -eq $sev }).Count }
    $files = [System.Collections.Generic.List[object]]::new()
    foreach ($f in $script:A2C.Files) { $files.Add([ordered]@{ file = $f.File; rows = $f.Rows }) }
    $info = [ordered]@{
        tool        = $script:A2CBrand.Tool
        version     = $script:A2CVersion
        startedUtc  = Format-A2CDate $script:A2C.StartedUtc
        finishedUtc = Format-A2CDate $fin
        durationSec = [math]::Round(($fin - $script:A2C.StartedUtc).TotalSeconds, 1)
        host        = $script:A2C.Env
        tenantId    = $script:A2C.TenantId
        account     = $script:A2C.Account
        options     = [ordered]@{
            profile = $o.Profile; reports = @($o.Reports); addReports = @($o.AddReports); skipReports = @($o.SkipReports)
            selected = @($script:A2C.Selected); datasets = @($script:A2C.Datasets); workspaceFilter = @($o.Workspace); excludeWorkspace = @($o.ExcludeWorkspace)
            personalWorkspaces = [bool]$o.PbiIncludePersonalWorkspaces; clientScope = [bool]$o.ClientScope; refreshHistory = $o.PbiRefreshHistory; runHistory = $o.PbiRunHistory; workspaceListFile = [bool]$o.WorkspaceListFile; activityDays = $o.PbiActivityDays
            exportMaxSizeMB = $o.PbiExportMaxSizeMB; zip = [bool]$o.Zip; noExcel = [bool]$o.NoExcel
        }
        workspaces  = [ordered]@{ inScope = @($script:A2C.Data['PbiWorkspaces'] | Where-Object { $null -ne $_ }).Count; visible = $script:A2C.Data['PbiWorkspacesVisible']
            tenant = $(if ($null -ne $script:A2C.Data['PbiAdmin']) { $script:A2C.Data['PbiAdmin'].Workspaces.Count } else { $null }) }
        stages      = $script:A2C.Stages.ToArray()
        apiCalls    = $script:A2C.Stats
        files       = $files.ToArray()
        findings    = $findingCounts
        coverage    = @($script:A2C.Coverage)
        warnings    = $script:A2C.Warnings.Count
        errors      = $script:A2C.Errors.ToArray()
    }
    [void](Save-A2CText -File ($script:A2C.RunStamp + '__run.json') -Text (ConvertTo-Json -InputObject $info -Depth 8) -NoBom)
}

function Open-A2CGeneralLog {
    # <yyyyMMdd-HHmmss>__general_log.txt in the run folder (same time stamp as the folder), starting with the lines
    # logged before the folder existed. With -Mask the log holds real names, so it is kept next to the shareable
    # folder instead, as <yyyyMMdd-HHmmss>__general_log.PRIVATE.txt.
    $run = $script:A2C
    $path = [IO.Path]::Combine($run.OutDir, $run.RunStamp + '__general_log.txt')
    if ($run.Options.Mask) { $path = [IO.Path]::Combine($run.Env['OutputBase'], $run.RunStamp + '__general_log.PRIVATE.txt') }
    $text = ''
    if ($run.Log.Count -gt 0) { $text = [string]::Join("`r`n", $run.Log) + "`r`n" }
    [IO.File]::WriteAllText($path, $text, $script:A2CUtf8Bom)
    $run.LogPath = $path
}

function Save-A2CEarlyLog {
    # The run stopped before its folder was created (e.g. Az.Accounts or sign-in): keep the log where the folder would be.
    $run = $script:A2C
    $name = $run.StartedUtc.ToLocalTime().ToString('yyyyMMdd-HHmmss') + '__general_log'
    if ($run.Options.Mask) { $name += '.PRIVATE.txt' } else { $name += '.txt' }
    $path = [IO.Path]::Combine((Get-A2COutputBase), $name)
    [IO.File]::WriteAllText($path, [string]::Join("`r`n", $run.Log) + "`r`n", $script:A2CUtf8Bom)
    $run.LogPath = $path
}

function New-A2CZip {
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $zipPath = [IO.Path]::Combine($script:A2C.Env['OutputBase'], $script:A2C.RunName + '.zip')
    if (Test-Path -LiteralPath $zipPath) { Remove-Item -LiteralPath $zipPath -Force }
    $root = $script:A2C.OutDir
    $zip = [IO.Compression.ZipFile]::Open($zipPath, [IO.Compression.ZipArchiveMode]::Create)
    try {
        foreach ($f in @(Get-ChildItem -LiteralPath $root -Recurse -File)) {
            $rel = $f.FullName.Substring($root.Length).TrimStart('\', '/')
            if ($script:A2C.Options.Mask -and ($rel -like '*general_log*' -or $rel -like '.cache*')) { continue }
            $entry = $script:A2C.RunName + '/' + ($rel -replace '\\', '/')
            [void][IO.Compression.ZipFileExtensions]::CreateEntryFromFile($zip, $f.FullName, $entry, [IO.Compression.CompressionLevel]::Optimal)
        }
    }
    finally { $zip.Dispose() }
    return $zipPath
}

function Export-A2CExcel {
    Import-Module ImportExcel -ErrorAction Stop
    $path = [IO.Path]::Combine($script:A2C.OutDir, $script:A2C.RunName + '.xlsx')
    $order = [System.Collections.Generic.List[string]]::new()
    $order.Add('findings.csv')
    # one sheet per report
    foreach ($f in @($script:A2C.Out.Keys | Where-Object { $_ -like '*.csv' -and $_ -ne 'findings.csv' -and $_ -ne 'coverage.csv' } |
                Sort-Object)) { $order.Add($f) }
    $order.Add('coverage.csv')
    $pkg = $null
    $max = [int]$script:A2C.Options.PropertiesMaxChars
    $used = @{}
    foreach ($file in $order) {
        $rows = Get-A2COutputRows $file
        if ($rows.Count -eq 0) { continue }
        $schema = $script:A2CSchemas[$file]
        $objs = [System.Collections.Generic.List[object]]::new()
        foreach ($row in $rows) {
            $o = [ordered]@{}
            foreach ($c in $schema) {
                $v = $null
                if ($row.Contains($c.Name)) { $v = $row[$c.Name] }
                if ($c.Kind -eq 'num' -and $null -ne $v -and "$v" -ne '') {
                    $d = 0.0
                    if ([double]::TryParse([string](Format-A2CCell $v), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { $o[$c.Name] = $d; continue }
                }
                $s = Format-A2CCell $v
                if ($s.Length -gt $max) { $s = $s.Substring(0, $max - 20) + '...[truncated]' }
                $o[$c.Name] = $s
            }
            $objs.Add([pscustomobject]$o)
        }
        $ws = ($file -replace '\.csv$', '') -replace '[\\/]', ' '   # PowerBI/reports.csv -> sheet "PowerBI reports"
        if ($ws.Length -gt 31) { $ws = $ws.Substring(0, 31) }
        if ($used.ContainsKey($ws)) { continue }
        $used[$ws] = $true
        $p = @{ WorksheetName = $ws; TableName = ('t_' + ($ws -replace '[^A-Za-z0-9_]', '_')); TableStyle = 'Medium2'; FreezeTopRow = $true
            BoldTopRow = $true; NoNumberConversion = '*'; PassThru = $true }
        if ($objs.Count -le 5000) { $p['AutoSize'] = $true; $p['MaxAutoSizeRows'] = 1000 }
        if ($null -eq $pkg) { $p['Path'] = $path } else { $p['ExcelPackage'] = $pkg }
        $pkg = $objs | Export-Excel @p
    }
    if ($null -ne $pkg) { Close-ExcelPackage $pkg }
    $script:A2C.Files.Add([pscustomobject]@{ File = [IO.Path]::GetFileName($path); Rows = $used.Count; Path = $path })
    Write-A2CLog -Level Ok ('Excel workbook: ' + $path)
}

function Write-A2CSummary {
    param([string] $ZipPath)
    $run = $script:A2C
    $dur = [DateTime]::UtcNow - $run.StartedUtc
    $high = @($run.Findings | Where-Object { $_.Severity -eq 'High' }).Count
    $calls = [int]$run.Stats['pbi'] + [int]$run.Stats['fabric'] + [int]$run.Stats['sql'] + [int]$run.Stats['onelake'] + [int]$run.Stats['flow'] + [int]$run.Stats['graph']
    $wsCount = @($run.Data['PbiWorkspaces'] | Where-Object { $null -ne $_ }).Count
    $adminWs = $run.Data['PbiAdmin']
    if ($null -ne $adminWs -and $adminWs.Workspaces.Count -gt $wsCount) { $wsCount = $adminWs.Workspaces.Count }
    $done = 'Done in {0:hh\:mm\:ss}: {1} workspace(s), {2} files, {3} findings ({4} high), {5} API calls.' -f $dur, $wsCount, $run.Files.Count, $run.Findings.Count, $high, $calls
    Add-A2CLogLine -Level Ok $done
    Write-Host ''
    Write-Host $done -ForegroundColor Green
    Write-Host ('Results: ' + $run.OutDir) -ForegroundColor Green
    if ($ZipPath) {
        Write-Host ('Zip    : ' + $ZipPath) -ForegroundColor Green
        if ($run.Env.CloudShell) { Write-Host '         In Cloud Shell use "Manage files > Download" to get it.' -ForegroundColor DarkGray }
    }
    if ($run.LogPath) { Write-Host ('Log    : ' + $run.LogPath) -ForegroundColor DarkGray }
    $partial = @($run.Coverage | Where-Object { $_['Status'] -eq 'Partial' -or $_['Status'] -eq 'Skipped' }).Count
    if ($partial -gt 0) { Write-Host ('{0} area(s) partly or not collected with this account - see coverage.csv' -f $partial) -ForegroundColor Yellow }
    if ($run.Errors.Count -gt 0) { Write-Host ('{0} error(s) - see the log and coverage.csv' -f $run.Errors.Count) -ForegroundColor Yellow }
    Write-Host ''
    Write-Host ('{0} by {1} - {2}' -f $script:A2CBrand.Tool, $script:A2CBrand.Author, $script:A2CBrand.Site) -ForegroundColor Cyan
    Write-Host ($script:A2CBrand.Tagline + '  ' + $script:A2CBrand.Email) -ForegroundColor Cyan
    Write-Host ''
}

#endregion

#region ============================== 10. MAIN ========================================================

$script:A2CListParameters = @('Reports', 'AddReports', 'SkipReports')

function Split-A2CNameList {
    # Names from arrays and/or comma, semicolon or space separated text ("RBAC, Policy")
    param($Value)
    $l = [System.Collections.Generic.List[string]]::new()
    foreach ($v in @($Value)) { if ($null -ne $v) { foreach ($p in ([string]$v -split '[,;\s]+')) { if ($p) { $l.Add($p) } } } }
    return , $l.ToArray()
}

function Resolve-A2CListArguments {
    # Started with pwsh -File, from cmd or by a scheduler, "-Reports RBAC, Policy" arrives as "RBAC," for -Reports plus
    # the loose word "Policy". The loose words are in RemainingArguments, in command-line order like the bound
    # parameters, so a list whose last value ends with a comma takes the next ones. Returns the parameters without
    # RemainingArguments; any other loose word is an error.
    param([System.Collections.IDictionary] $Params)
    $out = [ordered]@{}
    $rest = [System.Collections.Generic.Queue[string]]::new()
    foreach ($x in @($Params['RemainingArguments'])) { if ($null -ne $x -and [string]$x -ne '') { $rest.Enqueue([string]$x) } }
    foreach ($k in @($Params.Keys)) {
        if ($k -eq 'RemainingArguments') { continue }
        $v = $Params[$k]
        if ($script:A2CListParameters -contains $k) {
            $vals = [System.Collections.Generic.List[string]]::new()
            foreach ($x in @($v)) { if ($null -ne $x) { $vals.Add([string]$x) } }
            while ($rest.Count -gt 0 -and $vals.Count -gt 0 -and $vals[$vals.Count - 1].TrimEnd().EndsWith(',')) { $vals.Add($rest.Dequeue()) }
            $v = $vals.ToArray()
        }
        $out[$k] = $v
    }
    if ($rest.Count -gt 0) {
        throw ('Unexpected argument(s): {0}. Lists are written with commas, for example -Reports RBAC,Policy' -f [string]::Join(' ', $rest.ToArray()))
    }
    return $out
}

function New-A2COptions {
    param([System.Collections.IDictionary] $Params)
    $o = [ordered]@{
        TenantId = ''; Workspace = @(); ExcludeWorkspace = @(); PbiIncludePersonalWorkspaces = $false; OutputPath = ''; ClientScope = $false; ClientScopeIgnored = $false
        Profile = 'Standard'; Reports = @(); AddReports = @(); SkipReports = @(); ListReports = $false
        SelectionSpecified = $false
        PbiRefreshHistory = 20; PbiRunHistory = 20; WorkspaceListFile = ''; PbiNotebookFormat = 'Source'; PbiActivityDays = 7; PbiExportMaxSizeMB = 1024
        Mask = $false; Zip = $false; NoExcel = $false; UseDeviceAuthentication = $false; PropertiesMaxChars = 32000; Delimiter = ','
        V1 = @{}; Now = [DateTime]::UtcNow
    }
    if ($null -eq $Params) { $Params = @{} }
    $Params = Resolve-A2CListArguments $Params
    foreach ($k in @($Params.Keys)) {
        $v = $Params[$k]
        if ($v -is [System.Management.Automation.SwitchParameter]) { $v = $v.IsPresent }
        switch ($k) {
            'ReportProfile' { $o.Profile = [string]$v; $o.SelectionSpecified = $true }
            'Profile'       { $o.Profile = [string]$v; $o.SelectionSpecified = $true }
            'Reports'       { $o.Reports = @($v); $o.SelectionSpecified = $true }
            'AddReports'    { $o.AddReports = @($v); $o.SelectionSpecified = $true }
            'SkipReports'   { $o.SkipReports = @($v); $o.SelectionSpecified = $true }
            'Mask'          { }   # not in this script: the mask functions live in azure2csv.ps1
            default         { if ($o.Contains($k)) { $o[$k] = $v } }
        }
    }
    # workspace names can contain spaces and commas: they are taken as given, one per array element
    $o.Workspace = @($o.Workspace | Where-Object { $_ } | ForEach-Object { ([string]$_).Trim() })
    $o.ExcludeWorkspace = @($o.ExcludeWorkspace | Where-Object { $_ } | ForEach-Object { ([string]$_).Trim() })
    # names are checked here instead of by ValidateSet (see the param block); unknown ones stop the run before the sign-in
    foreach ($k in @('Reports', 'AddReports', 'SkipReports')) { $o[$k] = Split-A2CNameList $o[$k] }
    [void](Expand-A2CReportNames (@($o.Reports) + @($o.AddReports) + @($o.SkipReports)) -IncludeOptIn)
    if ($o.TenantId) { $o.TenantId = ([string]$o.TenantId).Trim().ToLowerInvariant() }
    if ($o.WorkspaceListFile) {
        if (-not (Test-Path -LiteralPath $o.WorkspaceListFile -PathType Leaf)) { throw ("-WorkspaceListFile not found: {0}" -f $o.WorkspaceListFile) }
        $o.WorkspaceListFile = (Resolve-Path -LiteralPath $o.WorkspaceListFile).ProviderPath
    }
    if (-not $script:A2CProfiles.Contains([string]$o.Profile)) { throw ("Unknown profile '{0}'" -f $o.Profile) }
    if ($o.ClientScope) {
        # share-safe scope run: the admin report only, whatever else was asked for
        $o.ClientScopeIgnored = ($o.SelectionSpecified -and (@($o.Reports) + @($o.AddReports) + @($o.SkipReports)).Count -gt 0)
        $o.Reports = @('PowerBIAdmin', 'PowerBIActivity', 'Principals'); $o.AddReports = @(); $o.SkipReports = @(); $o.Zip = $true; $o.SelectionSpecified = $true
    }
    return [pscustomobject]$o
}

function Expand-A2CReportNames {
    # Report and group names -> report names. Group names add opt-in reports only with -IncludeOptIn.
    param($Names, [switch] $IncludeOptIn)
    $out = [System.Collections.Generic.List[string]]::new()
    foreach ($n in @($Names)) {
        if (-not $n) { continue }
        # All (and -Profile Full) leaves out Explicit reports such as the repository export: they run only when named
        if ($n -ieq 'All') { foreach ($r in $script:A2CReportCatalog) { if (-not $r.Explicit -or $IncludeOptIn) { $out.Add($r.Name) } }; continue }
        if ($script:A2CReportGroups -contains $n) {
            foreach ($r in $script:A2CReportCatalog) { if ($r.Group -ieq $n -and ((-not $r.OptIn) -or $IncludeOptIn)) { $out.Add($r.Name) } }
            continue
        }
        $rep = $null
        foreach ($r in $script:A2CReportCatalog) { if ($r.Name -ieq $n) { $rep = $r; break } }
        if ($null -eq $rep) { throw ("Unknown report or group '{0}'. Run -ListReports to see valid names." -f $n) }
        $out.Add($rep.Name)
    }
    return , $out.ToArray()
}

function Resolve-A2CSelection {
    param($Options)
    if (@($Options.Reports).Count -gt 0) { $base = Expand-A2CReportNames $Options.Reports }
    else { $base = Expand-A2CReportNames $script:A2CProfiles[[string]$Options.Profile] }
    $add = Expand-A2CReportNames $Options.AddReports
    $skipList = [System.Collections.Generic.List[string]]::new()
    foreach ($s in (Expand-A2CReportNames $Options.SkipReports -IncludeOptIn)) { $skipList.Add($s) }
    if ($Options.V1.ContainsKey('createAllResourcesReport') -and -not $Options.V1['createAllResourcesReport']) { $skipList.Add('Resources'); $skipList.Add('Tags') }
    if ($Options.V1.ContainsKey('createRBACreports') -and -not $Options.V1['createRBACreports']) { $skipList.Add('RBAC') }
    if ($Options.V1.ContainsKey('createPrivateDNSreport') -and -not $Options.V1['createPrivateDNSreport']) { $skipList.Add('PrivateDNS') }
    $want = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($n in $base) { [void]$want.Add($n) }
    foreach ($n in $add) { [void]$want.Add($n) }
    foreach ($n in $skipList) { [void]$want.Remove($n) }
    $sel = [System.Collections.Generic.List[string]]::new()
    foreach ($r in $script:A2CReportCatalog) { if ($want.Contains($r.Name)) { $sel.Add($r.Name) } }
    return , $sel.ToArray()
}

function Resolve-A2CDatasets {
    # Selected reports -> datasets to collect (transitive dependencies), in execution order.
    param([string[]] $Reports)
    $need = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($rn in $Reports) {
        foreach ($r in $script:A2CReportCatalog) { if ($r.Name -eq $rn) { foreach ($d in $r.Needs) { [void]$need.Add($d) } } }
    }
    $changed = $true
    while ($changed) {
        $changed = $false
        foreach ($d in @($need)) { foreach ($dep in $script:A2CDatasetCatalog[$d].Needs) { if ($need.Add($dep)) { $changed = $true } } }
    }
    $ordered = [System.Collections.Generic.List[string]]::new()
    foreach ($k in $script:A2CDatasetCatalog.Keys) { if ($need.Contains($k)) { $ordered.Add($k) } }
    return , $ordered.ToArray()
}

function Show-A2CReportList {
    param($Options)
    Write-Host ''
    Write-Host ('{0} {1} - report catalog' -f $script:A2CBrand.Tool, $script:A2CVersion) -ForegroundColor Cyan
    $profiles = @($script:A2CProfiles.Keys)
    $selected = @()
    if ($Options.SelectionSpecified) { $selected = Resolve-A2CSelection $Options }
    $rows = foreach ($r in $script:A2CReportCatalog) {
        $inProfiles = @(foreach ($p in $profiles) {
                $members = Expand-A2CReportNames $script:A2CProfiles[$p]
                if ($members -contains $r.Name) { $p }
            })
        $o = [ordered]@{}
        if ($Options.SelectionSpecified) { $o['Run'] = $(if ($selected -contains $r.Name) { 'yes' } else { '' }) }
        $o['Group'] = $r.Group
        $o['Report'] = $r.Name + $(if ($r.Explicit) { ' **' } elseif ($r.OptIn) { ' *' } else { '' })
        $o['Speed'] = $r.Speed
        $o['Profiles'] = [string]::Join(',', $inProfiles)
        $o['Files'] = $r.Files
        [pscustomobject]$o
    }
    if ($Options.SelectionSpecified) {
        $desc = if (@($Options.Reports).Count -gt 0) { '-Reports ' + [string]::Join(', ', @($Options.Reports)) } else { '-Profile ' + $Options.Profile }
        if (@($Options.AddReports).Count -gt 0) { $desc += ' -AddReports ' + [string]::Join(', ', @($Options.AddReports)) }
        if (@($Options.SkipReports).Count -gt 0) { $desc += ' -SkipReports ' + [string]::Join(', ', @($Options.SkipReports)) }
        Write-Host ('Selection: {0}  ->  {1} reports' -f $desc, @($selected).Count) -ForegroundColor Cyan
        $rows = @($rows | Where-Object { $_.Run -eq 'yes' })
    }
    Write-Host ($rows | Format-Table -AutoSize -Wrap | Out-String -Width 220)
    if ($Options.SelectionSpecified) {
        $ds = Resolve-A2CDatasets $selected
        Write-Host ('Data collected: ' + [string]::Join(', ', $ds))
        $slow = @($script:A2CReportCatalog | Where-Object { $selected -contains $_.Name -and $_.Speed -eq 'Slow' } | ForEach-Object { $_.Name })
        if ($slow.Count -gt 0) {
            Write-Host ('Slow reports in this selection (one or more calls per item): ' + [string]::Join(', ', $slow)) -ForegroundColor Yellow
        }
        else { Write-Host 'No slow reports in this selection.' -ForegroundColor Green }
    }
    else {
        Write-Host '* = opt-in: only in -Profile Full or when named in -Reports/-AddReports.'
        Write-Host '** = only when named in -Reports/-AddReports (not in -Profile Full).'
        Write-Host ('Groups: ' + [string]::Join(', ', $script:A2CReportGroups) + ', All')
        Write-Host 'Examples:  -Profile Full   -Reports PowerBI,Fabric   -AddReports PowerBIModels,PowerBIExport   -SkipReports PowerBIAdmin'
    }
    Write-Host ''
}

function Invoke-A2CDataset {
    param([string] $Name)
    $info = $script:A2CDatasetCatalog[$Name]
    Write-A2CLog -Level Step ('{0}: {1}' -f $Name, $info.Desc)
    $t0 = [DateTime]::UtcNow
    $status = 'Collected'
    $note = ''
    try {
        $res = & ('Get-A2CDataset' + $Name)
        if ($null -ne $res -and $res.PSObject.Properties['Status']) { $status = [string]$res.Status; $note = [string]$res.Note }
    }
    catch {
        $status = 'Failed'
        $note = $_.Exception.Message
        Write-A2CLog -Level Error ('{0} failed: {1}' -f $Name, $note)
        Add-A2CLogLine -Level Detail ('at ' + (Get-A2CStackText $_))
        Add-A2CCoverage -Area $Name -Status Failed -Detail $note
    }
    $script:A2C.DatasetStatus[$Name] = $status
    $script:A2C.Stages.Add([ordered]@{ Stage = 'Dataset:' + $Name; Status = $status; DurationSec = [math]::Round(([DateTime]::UtcNow - $t0).TotalSeconds, 1); Note = $note })
}

function New-A2CDatasetResult {
    param([string] $Status = 'Collected', [string] $Note = '')
    return [pscustomobject]@{ Status = $Status; Note = $Note }
}

function Invoke-A2CReportBuilder {
    param([string] $Name)
    $t0 = [DateTime]::UtcNow
    $status = 'Built'
    $note = ''
    try {
        & ('New-A2CReport' + $Name)
    }
    catch {
        $status = 'Failed'
        $note = $_.Exception.Message + ' @ ' + $_.InvocationInfo.ScriptLineNumber
        Write-A2CLog -Level Error ('Report {0} failed: {1}' -f $Name, $note)
        Add-A2CLogLine -Level Detail ('at ' + (Get-A2CStackText $_))
        Add-A2CCoverage -Area ('Report ' + $Name) -Status Failed -Detail $note
    }
    $script:A2C.Stages.Add([ordered]@{ Stage = 'Report:' + $Name; Status = $status; DurationSec = [math]::Round(([DateTime]::UtcNow - $t0).TotalSeconds, 1); Note = $note })
}

function Get-A2COutputBase {
    $base = $script:A2C.Options.OutputPath
    if (-not $base) { $base = (Get-Location).ProviderPath }
    if (-not (Test-Path -LiteralPath $base)) { New-Item -ItemType Directory -Path $base -Force | Out-Null }
    return (Resolve-Path -LiteralPath $base).ProviderPath
}

function Initialize-A2COutputFolder {
    $base = Get-A2COutputBase
    $tenant8 = ''
    if ($script:A2C.TenantId) { $tenant8 = $script:A2C.TenantId.Substring(0, [math]::Min(8, $script:A2C.TenantId.Length)) }
    $ts = [DateTime]::Now.ToString('yyyyMMdd-HHmmss')
    $name = '{0}_{1}_{2}' -f $script:A2CBrand.Tool, $ts, $tenant8
    $dir = [IO.Path]::Combine($base, $name)
    $n = 1
    while (Test-Path -LiteralPath $dir) { $n++; $name = '{0}_{1}_{2}-{3}' -f $script:A2CBrand.Tool, $ts, $tenant8, $n; $dir = [IO.Path]::Combine($base, $name) }
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
    $script:A2C.OutDir = $dir
    $script:A2C.RunName = $name
    $script:A2C.RunStamp = $ts
    if ($n -gt 1) { $script:A2C.RunStamp = $ts + '-' + $n }
    $script:A2C.Env['OutputBase'] = $base
    Open-A2CGeneralLog
}

function Invoke-A2CRun {
    param($Options)
    $run = New-A2CRunContext -Options $Options
    $run.Selected = Resolve-A2CSelection $Options
    $run.Datasets = Resolve-A2CDatasets $run.Selected
    if (@($run.Selected).Count -eq 0) { throw 'No reports selected. Run -ListReports to see the report names.' }

    $banner = '{0} {1}  -  {2}  -  {3}' -f $script:A2CBrand.Tool, $script:A2CVersion, $script:A2CBrand.Site, $script:A2CBrand.Email
    $reports = 'Reports ({0}): {1}' -f @($run.Selected).Count, [string]::Join(', ', $run.Selected)
    Write-Host ''
    Write-Host $banner -ForegroundColor Cyan
    Write-Host $reports -ForegroundColor DarkGray
    Add-A2CLogLine -Level Info $banner
    Add-A2CLogLine -Level Info $reports

    $run.Env = Get-A2CHostInfo
    if ($run.Env.PSEdition -ne 'Core') {
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { Write-Verbose ('TLS 1.2 could not be enabled: ' + $_.Exception.Message) }
    }
    $run.Env['AzAccounts'] = Assert-A2CPrerequisites
    $psName = $(if ($run.Env.PSEdition -eq 'Core') { 'PowerShell' } else { 'Windows PowerShell' })
    $where = $(if ($run.Env.CloudShell) { ' in Cloud Shell' } else { '' })
    Write-A2CLog ('{0} {1}{2}, Az.Accounts {3}' -f $psName, $run.Env.PSVersion, $where, $run.Env.AzAccounts)
    $az = Connect-A2CAzure -Options $Options
    $run.TenantId = [string]$Options.TenantId
    if (-not $run.TenantId -and $az.Tenant) { $run.TenantId = ([string]$az.Tenant.Id).ToLowerInvariant() }
    $run.Account = [string]$az.Account.Id
    Initialize-A2COutputFolder
    Write-A2CLog ('Tenant {0}, signed in as {1}' -f $run.TenantId, $run.Account)
    Write-A2CLog ('Output folder: {0}' -f $run.OutDir)
    if ($Options.ClientScope) {
        Write-A2CLog -Level Notice '-ClientScope: only the scope of the tenant is collected, through the admin APIs.'
        Write-A2CLog -Level Detail 'Collected: workspaces with type, state, capacity, number of users, workspace admins, item counts, last change seen, data source types and gateway use; capacities.'
        Write-A2CLog -Level Detail 'Re-pointing map: data source type and host name per workspace (no database, path, file or credentials).'
        Write-A2CLog -Level Detail ('Counted, not written: activity log of the last {0} day(s) (active users per workspace and capacity); people, as numbers. Named: groups, apps and workspace admins.' -f $Options.PbiActivityDays)
        Write-A2CLog -Level Detail 'Looked up in Entra ID: whether the accounts still exist (deleted or disabled).'
        Write-A2CLog -Level Detail 'Not collected: item names, data source addresses, model schemas, DAX and M, code, tenant settings, sharing links, raw Scanner API answers.'
        if ($Options.ClientScopeIgnored) { Write-A2CLog -Level Notice '-Reports / -AddReports / -SkipReports are ignored with -ClientScope.' }
    }
    if ($run.Env.CloudShell -and -not $run.Env.ClouddriveExists) {
        Write-A2CLog -Level Warn 'Cloud Shell without a clouddrive: files are lost when the session ends - download the zip before closing.'
    }

    foreach ($d in $run.Datasets) {
        Invoke-A2CDataset -Name $d
        if ($d -eq 'PbiWorkspaces' -and $run.DatasetStatus['PbiWorkspaces'] -eq 'Failed') { throw 'Cannot continue without the workspace list.' }
    }
    if ($Options.ClientScope -and -not (Test-A2CDataset 'PowerBIAdmin')) {
        Write-A2CLog -Level Warn '-ClientScope needs an account with the Fabric administrator role: nothing of the tenant was collected (coverage.csv has the reason).'
    }

    Write-A2CLog -Level Step 'Building reports'
    foreach ($rep in $run.Selected) { Invoke-A2CReportBuilder -Name $rep }
    foreach ($r in $script:A2CReportCatalog) {
        if ($run.Selected -notcontains $r.Name) { Add-A2CCoverage -Area ('Report ' + $r.Name) -Status NotSelected -Detail $r.Files }
    }

    Write-A2CLog -Level Step 'Findings'
    try { Invoke-A2CFindings }
    catch { Write-A2CLog -Level Error ('Findings failed: ' + $_.Exception.Message) }

    Write-A2CLog -Level Step 'Writing files'
    foreach ($f in @($script:A2C.Out.Keys | Sort-Object)) {
        try { Save-A2CCsv -File $f } catch { Write-A2CLog -Level Error ('Cannot write {0}: {1}' -f $f, $_.Exception.Message) }
    }
    foreach ($x in @($script:A2C.Files | Sort-Object File)) { Write-A2CLog -Level Detail ('{0,-44} {1,8} rows' -f $x.File, $x.Rows) }
    $excelOk = $false
    if (-not $Options.NoExcel) {
        if (Test-A2CModuleAvailable 'ImportExcel') { $excelOk = $true; Add-A2CCoverage -Area 'Excel workbook' -Status Collected -Detail ($run.RunName + '.xlsx') }
        else {
            Add-A2CCoverage -Area 'Excel workbook' -Status Skipped -Detail 'ImportExcel module not installed - Install-Module ImportExcel -Scope CurrentUser'
            Write-A2CLog 'Excel workbook skipped (ImportExcel module not installed).'
        }
    }
    Save-A2CExtraFiles
    if ($excelOk) {
        try { Export-A2CExcel } catch { Write-A2CLog -Level Error ('Excel export failed: ' + $_.Exception.Message) }
    }
    Save-A2CRunInfo
    $zip = ''
    if ($Options.Zip) {
        try { $zip = New-A2CZip } catch { Write-A2CLog -Level Error ('Zip failed: ' + $_.Exception.Message) }
    }
    Write-A2CSummary -ZipPath $zip
}

function Invoke-A2CMain {
    param([System.Collections.IDictionary] $BoundParameters)
    try {
        $opt = New-A2COptions -Params $BoundParameters
        if ($opt.ListReports) { Show-A2CReportList -Options $opt; return 0 }
    }
    catch {
        # a mistake in the parameters: one line, nothing was started
        Write-Host ('ERROR: ' + $_.Exception.Message) -ForegroundColor Red
        return 2
    }
    try {
        Invoke-A2CRun -Options $opt
        return 0
    }
    catch {
        Write-Host ''
        Write-Host ('ERROR: ' + $_.Exception.Message) -ForegroundColor Red
        if ($null -ne $script:A2C) {
            $script:A2C.Errors.Add($_.Exception.Message)
            Add-A2CLogLine -Level Error ('Run stopped: ' + $_.Exception.Message)
            Add-A2CLogLine -Level Detail ('at ' + (Get-A2CStackText $_))
            if ($script:A2C.OutDir) { try { Save-A2CRunInfo } catch { Write-Host ('Could not save the run info (__run.json): ' + $_.Exception.Message) -ForegroundColor Yellow } }
            elseif (-not $script:A2C.LogPath) { try { Save-A2CEarlyLog } catch { Write-Host ('Could not save the log: ' + $_.Exception.Message) -ForegroundColor Yellow } }
            if ($script:A2C.LogPath) { Write-Host ('Log: ' + $script:A2C.LogPath) -ForegroundColor Yellow }
        }
        return 1
    }
}

function Test-A2CStartedFromExplorer {
    # True when Windows started PowerShell only to run this script from Explorer ("Run with PowerShell", a shortcut,
    # Win+R). That window closes as soon as the script ends, so the result or the error would vanish unread.
    # Never true in an open console, a scheduled task, CI or -NonInteractive / hidden windows.
    try {
        if ($Host.Name -ne 'ConsoleHost' -or -not [Environment]::UserInteractive -or -not $PSCommandPath) { return $false }
        if ($PSVersionTable.PSEdition -eq 'Core' -and -not $IsWindows) { return $false }
        $cmd = [Environment]::CommandLine
        if ($cmd -notmatch [regex]::Escape([IO.Path]::GetFileName($PSCommandPath)) -or $cmd -match '\s-noni|\s-w\w*\s+hid') { return $false }
        $me = Get-CimInstance -ClassName Win32_Process -Filter ('ProcessId = ' + $PID) -ErrorAction Stop
        return ((Get-Process -Id $me.ParentProcessId -ErrorAction Stop).ProcessName -eq 'explorer')
    }
    catch { return $false }
}

if ($MyInvocation.InvocationName -ne '.' -and -not $env:POWERBI2CSV_NO_MAIN) {
    $a2cBound = [ordered]@{}   # in command-line order: Resolve-A2CListArguments relies on it
    foreach ($a2cKey in $PSBoundParameters.Keys) { $a2cBound[$a2cKey] = $PSBoundParameters[$a2cKey] }
    $a2cExit = Invoke-A2CMain -BoundParameters $a2cBound
    if (Test-A2CStartedFromExplorer) { [void](Read-Host 'Press Enter to close this window') }
    if ($a2cExit -ne 0) { exit $a2cExit }
}

#endregion
