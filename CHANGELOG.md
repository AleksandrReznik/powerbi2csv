# Changelog

## 1.3.0 - 2026-11-12

What can be copied as it is and what has to be refactored: the IDs, paths and addresses written into the code, a
portability class per item and the order to create the items in. 12 reports, up to 100 CSV files, 26 findings.

### Portability (`Dependencies`)
- `hardcodedReferences.csv`: every item ID, workspace ID, connection ID, OneLake path, SQL endpoint, URL, storage
  path and mail address written into a notebook, pipeline or dataflow, with what it points to (type, name, workspace)
  and its `Scope`: same workspace, other workspace, other environment (with the direction), connection, external or
  not in scope. Comment lines of notebooks are skipped. An ID the run cannot name is typed by the JSON key it stands
  under: a connection the account cannot see stays a connection, a Teams team or channel is outside Fabric.
- `itemPortability.csv`: one row per notebook, pipeline and dataflow with the counts per scope, the number of
  parameter and variable uses and a `Class`: Portable (nothing written in), Remap IDs (only IDs of the same
  workspace), Re-point (connections, outside addresses, mail addresses), Refactor (reaches into other workspaces or
  environments, or uses IDs the run cannot name).
- `migrationOrder.csv`: a wave per item from the dependency graph. Wave 1 depends on nothing in scope, wave n only on
  earlier waves; items that depend on each other in a circle get no wave and `InCycle`.
- `notebookDependencies.csv`: a table name now resolves to a lakehouse, by the name of an attached lakehouse or one in
  the same workspace (`lakehouse.schema.table`, `lakehouse.table`), any other name to the default lakehouse. New
  column `ResolvedBy` (ID, notebook name, lakehouse name, default lakehouse). Names built at run time stay open.
  Resolved tables also become `ReadsFrom` and `WritesTo` lines in `itemDependencies.csv`.
- `itemDependencies.csv`: new column `Crossing`: empty inside a workspace, `Workspace`, or
  `Environment (DEV -> PROD)` when the two workspace names carry different environments (no suffix = PROD, as in the
  environment matrix).
- Findings PBI-25 (dependencies that cross environments) and PBI-26 (notebooks, pipelines and dataflows that reach
  into other workspaces by a written-in ID or path).

### Fixed
- A notebook called with `%run ./Name` was not matched to the notebook `Name`, so most `%run` calls had no target.

### Changed
- `notebookDependencies.csv` and `itemDependencies.csv` each have one more column (`ResolvedBy` before `Evidence`,
  `Crossing` before `Via`). Read them by column name.

## 1.2.0 - 2026-11-06

How each model and workspace moves: access mode, real gateways, refresh identity and health per model, the
connections the target tenant needs, and a migration class per workspace. 12 reports, up to 97 CSV files, 24 findings.

### Migration handling (`PowerBI`, `PowerBIAdmin`)
- `modelMigration.csv`: one row per semantic model, from the workspaces the account is a member of and, as a Fabric
  administrator, the whole tenant. `AccessMode` (Import, DirectLake, DirectQuery, Composite, LiveConnection, Push,
  Unknown) with `IsComposite`, `DirectQueryToModel` and where it came from (definition, Scanner tables, admin content
  type, data sources); gateway type; refresh identity (gateway, shared cloud or personal cloud connection) with
  credential types, `PersonalCredential` and `StoredPassword`; refresh health (last status, refreshes, failures,
  failure rate); `HandlingClass`.
- `targetConnections.csv`: the distinct connection targets per connectivity (on-premises gateway, VNet gateway, cloud,
  Fabric) and data source type, with hosts, models, items, workspaces and credential types: the checklist for the
  target tenant's admin.
- `adminWorkspaces.csv` `MigrationClass` and `MigrationReason`: A scriptable copy, B re-point, C manual, D deprioritize;
  counted in `workspaceScopeSummary.csv` with the models per access mode, models through an on-premises gateway, models
  with a personal credential and connection targets.
- `capacityFootprint.csv` `ToMigrate`, `BatchSize`, `Batches`: workspaces per migration batch, a rule of thumb from
  the capacity size.
- `semanticModels.csv` `AccessMode` and `IsComposite`; `GatewayType` in the data source files of models, reports and
  dataflows.
- As an administrator: `/admin/datasets` for every model's on-premises gateway, refreshable and content type flags,
  and the refresh history of models outside the member workspaces where the account has a role (one refused call per
  workspace without one).
- Findings PBI-22 (models that refresh with a personal credential), PBI-23 (composite, live connection and push
  models), PBI-24 (migration class per workspace).

### Fixed
- `UsesGateway` in `adminWorkspaces.csv` and `repointingMap.csv` was true for every data source with a gateway ID,
  also for cloud connections, which carry the ID of a cloud gateway cluster. It now means an on-premises or VNet
  gateway (`GatewayType` says which).
- `refreshables.csv` stayed empty for a Fabric administrator: the tenant-wide refresh statistics of every dedicated
  capacity were read but only used for the re-pointing map. They are written now (`Source` column).
- `modelPartitions.csv` `Mode` and `semanticModels.csv` `StorageModes`: a partition without a mode now gets the
  model's default mode (`defaultMode`, import when absent) instead of an empty value or its source type (`query`).

## 1.1.0 - 2026-11-05

Migration scope for administrators, a share-safe run for a customer's administrator, and the principals behind every
role. 12 reports, up to 95 CSV files, 21 findings.

### Migration scope for administrators (`PowerBIAdmin`, `-ClientScope`)
- `adminWorkspaces.csv` per workspace: scope (personal or shared, with content or empty, not active), item count, last
  change seen, changed in the last 30 days, activity count, admins and a NoAdmin flag, data source types, number of
  data sources, gateway use. `adminDatasourceInstances.csv`: the number of workspaces that use each data source.
- `workspaceScopeSummary.csv`: the tenant's migration scope on one page.
- `-WorkspaceListFile` with the administrator role compares the list with the whole tenant, in both directions.
- `repointingMap.csv`: per workspace, also those the account cannot open, every data source host, gateway use, the
  items and models that use it and their refresh schedules (dedicated capacities).
- `capacityFootprint.csv`: an estimate per capacity from counts, refresh statistics and the activity log; active
  users per workspace in `adminWorkspaces.csv`.
- `-ClientScope`: a share-safe run for a customer's administrator; only the admin APIs and the Entra ID lookup, only
  scope metadata, host names but no addresses, people only as counts, no item names, no code, no raw Scanner API
  answer, zipped.

### Principals (`Principals`)
- `principals.csv`: every user, guest, group, service principal and managed identity with a role on a workspace,
  item, capacity, gateway or deployment pipeline, or owning a model or dataflow, de-duplicated across the APIs, with
  type, app ID, roles, workspace count, and whether the account was deleted or disabled in Entra ID (Microsoft Graph,
  by object ID or UPN). Departed admins count as no admin in `adminWorkspaces.csv`.

### Findings PBI-15 to PBI-21
- Shared workspaces with content that are not on the list, listed workspaces that are empty, personal or not active,
  shared workspaces with content and no administrator, the migration scope of the tenant, why there are so many
  personal workspaces, admins whose account was deleted or disabled, and the data source hosts to re-point at cutover.
- PBI-11 now reports listed workspaces that were not found, in the whole tenant when the account is an administrator.

## 1.0.0 - 2026-10-23

First public release. 11 reports, up to 90 CSV files, 14 findings. Read-only; works with the Az sign-in of a plain
workspace member, more with the Fabric administrator role.

### Power BI (`PowerBI`)
- Workspaces and their users, reports with pages and data sources, semantic models with data sources, users,
  parameters, refresh schedules and refresh history, dashboards and tiles, dataflows with data sources and
  transactions, lineage (report to model, model to dataflow), apps and their content, capacities and workloads,
  refreshables, gateways with data sources and users, deployment pipelines with stages, items, users and operations,
  available features.

### Microsoft Fabric (`Fabric`)
- Every item of every type with folder path, tags, logical ID and last run; folders; workspace role assignments;
  OneLake security roles; Git connection with the number of uncommitted and incoming items and the per-item status;
  workspace network policy; Spark settings, pools and the latest sessions; managed private endpoints; lakehouses with
  tables and shortcuts (target connection, location, subpath); warehouses; item schedules; run history of pipelines
  and notebooks; item connections; tenant connections, gateways, capacities and domains; deployment pipelines with
  every item type; a DEV / TEST / UAT / PROD matrix of the same solution.
- Type specific properties for dataflows, environments, mirrored databases, eventhouses, KQL databases, SQL databases
  and Spark job definitions.

### SQL endpoints (`FabricSql`)
- Tables, views, stored procedures and functions of every lakehouse SQL endpoint and warehouse, their columns, the
  code of views and procedures and references between databases. Fills the counts in `warehouses.csv` and
  `lakehouses.csv` and lists the tables of lakehouses with schemas, which the Fabric API does not return.

### Dependencies (`Dependencies`)
- From the item definitions: pipeline activities (what each pipeline runs, with stored procedure parameters),
  notebook dependencies (attached lakehouses, tables read and written, OneLake paths, notebooks called, URLs),
  dataflow sources and destinations (SharePoint site and file, SQL server and database, lakehouse), semantic model
  partitions and storage modes, DAX of measures and calculated columns, M expressions, row-level security roles,
  model sources.
- `itemDependencies.csv`: every item-to-item link in one file. `externalDependencies.csv`: every SharePoint site,
  SQL server, storage account and URL that the tenant depends on, and who uses it.
- `-WorkspaceListFile`: checks a list of expected workspaces (.csv or .xlsx) against what the account sees.

### Tenant-wide, as a Fabric administrator (`PowerBIAdmin`, `PowerBIActivity`)
- All workspaces including personal and orphaned ones, users per item, Scanner API metadata (data sources, model
  tables and expressions), tenant settings, admin capacities and apps, widely shared items, domains; the activity log.

### Downloads (`FabricCode`, `PowerBIExport`)
- `FabricCode`: notebooks (Fabric source or .ipynb), pipelines, dataflows, semantic models (TMDL) and every other
  item type with a definition, dataflow gen1 definitions, and the SQL of views and procedures, as files.
- `PowerBIExport`: the same plus report definitions and report files (.pbix / .rdl), with a size budget.

### Sizes and flows (`FabricStorage`, `PowerAutomate`)
- OneLake file count and bytes per lakehouse and warehouse table and per Files area.
- Power Automate environments and flows with triggers and connectors.

### Semantic model schema (`PowerBIModels`)
- Tables, columns, measures and relationships of every semantic model through DAX `INFO.VIEW` queries.

### Findings PBI-01 to PBI-14
- Failed refreshes and runs, models without an enabled refresh schedule, unused models, single-administrator
  workspaces, owners who are not members, workspaces without a capacity or without Git, items published to the web
  or shared with the whole organization, partly readable workspaces, listed workspaces the account cannot see,
  external targets, items that exist only in DEV or TEST.

### Output
- `PowerBI\*.csv` (UTF-8 with BOM, every field quoted), `findings.csv`, `coverage.csv`, a log, `__run.json`, an
  optional Excel workbook (ImportExcel) and `-Zip`.
