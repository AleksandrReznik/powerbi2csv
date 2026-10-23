# Changelog

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
