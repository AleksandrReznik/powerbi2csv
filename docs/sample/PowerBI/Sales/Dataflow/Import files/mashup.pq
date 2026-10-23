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