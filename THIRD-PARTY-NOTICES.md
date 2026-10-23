# Third-party notices

powerbi2csv is released under the MIT License (see `LICENSE`). It builds on the following public work.

## azure2csv

The structure, the shared helper functions (logging, CSV writer, HTTP transport with retries, report catalog,
findings) and the offline test harness are taken from <https://github.com/AleksandrReznik/Azure2csv> (MIT License,
same author), so that the Power BI blocks can be merged back into it.

## Microsoft documentation used for rules

The findings and their recommendations are based on public Microsoft documentation for Power BI and Microsoft
Fabric (workspace roles, refresh, Git integration, sharing and publish-to-web settings, licences). Links are
included in `findings.csv`. The REST APIs used are documented at <https://learn.microsoft.com/rest/api/power-bi/>
and <https://learn.microsoft.com/rest/api/fabric/>.

## Optional components

- **ImportExcel** (Doug Finke, Apache 2.0) is used only when it is already installed, to create the Excel workbook.
  It is never downloaded or installed by powerbi2csv.
- **Az.Accounts** (Microsoft, MIT) is the only required module.
- **System.Data.SqlClient** (part of Windows PowerShell and PowerShell 7) is used for the read-only catalog queries
  of the `FabricSql` report.
