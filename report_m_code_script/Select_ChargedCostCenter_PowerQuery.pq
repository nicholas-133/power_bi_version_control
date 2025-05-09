/*
    This query in Power Query Editor is to pull the sourced .txt file for ChargedCostCenterCalculation, and store it as a query returing a block of .txt. 
    This allows us to reference it in the Main Fact Table [TimesheetLineActualDataSet] with the Expression.Evaluate() function.
*/
let
    Source = SharePoint.Files("https://tasb.sharepoint.com/sites/PMO/ProjectOnlineAdmin", [ApiVersion = 15]),
    #"Filtered Rows" = Table.SelectRows(Source, each ([Folder Path] = "https://tasb.sharepoint.com/sites/PMO/ProjectOnlineAdmin/Shared Documents/Report Database/Formula Expressions/")),
    #"PowerQueryFunction txt_https://tasb sharepoint com/sites/PMO/ProjectOnlineAdmin/Shared Documents/Report Database/PWA Dataflow/" = #"Filtered Rows"{[Name="ChargedCostCenterCalculation.txt",#"Folder Path"="https://tasb.sharepoint.com/sites/PMO/ProjectOnlineAdmin/Shared Documents/Report Database/Formula Expressions/"]}[Content],
    #"Imported Text" = Text.FromBinary(#"PowerQueryFunction txt_https://tasb sharepoint com/sites/PMO/ProjectOnlineAdmin/Shared Documents/Report Database/PWA Dataflow/")

in
    #"Imported Text"