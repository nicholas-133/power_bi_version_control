
```powerquery
let
    Source = OData.Feed("https://tasb.sharepoint.com/sites/pwa/_api/Projectdata", null, [Implementation="2.0"]),
    TimesheetLineActualDataSet_table = Source{[Name="TimesheetLineActualDataSet",Signature="table"]}[Data],
    // Removed Columns other than the selected
    #"Removed Other Columns" = Table.SelectColumns(TimesheetLineActualDataSet_table,{"TimesheetLineId", "TimeByDay", "ActualWorkBillable", "Comment"}),
    
    /* Filter Date greater than or equal to Report Start Date and less than or equal to ReportEndDate */
    #"Filtered Date" = Table.SelectRows(#"Removed Other Columns", each [TimeByDay] >= ReportStartDate and [TimeByDay] <= ReportEndDate),

    /* Filter out zero hour Actuals */
    #"Filtered 0 Work" = Table.SelectRows(#"Filtered Date", each ([ActualWorkBillable] <> 0)),

    /* Renamed ActualWorkBillable to Hours */
    #"Renamed Hours" = Table.RenameColumns(#"Filtered 0 Work",{{"ActualWorkBillable", "Hours"}}),

    /* Change TimeByDay (date) to Date Type instead of Date/Time Type */
    #"Changed Type to Date" = Table.TransformColumnTypes(#"Renamed Hours",{{"TimeByDay", type date}}),
    
    /* Inner Join Merge with TimesheetLines Table */
    #"Merged TimesheetLines" = Table.NestedJoin(#"Changed Type to Date", {"TimesheetLineId"}, TimesheetLines, {"TimesheetLineId"}, "TimesheetLines", JoinKind.Inner),
    
    /* Expand the TimesheetLines table column to retrieve Id fields and [TimesheetClassName] which will be used to identify overhead types */
    #"Expanded TimesheetLines" = Table.ExpandTableColumn(#"Merged TimesheetLines", "TimesheetLines", {"TimesheetOwnerId", "ProjectId", "TaskId"}, {"ResourceId", "ProjectId", "TaskId"}),
    
    /* Rmoved [TimesheetLinesId] as it is no longer needed */     
    #"Removed Columns" = Table.RemoveColumns(#"Expanded TimesheetLines",{"TimesheetLineId"}),
    
    /* Appended the AssignmentTimePhased Table to include Non-Timesheet Resources (e.g. Jayme Edwards). Since Non-Timesheet Resources do not show in the TimesheetLineActualDataSet, we have to append the current table with the AssignmentTimePhasedDataSet.
        In order for this Append to work correctly, we must match column names, column data types, and column count of the AssignmentTimePhasedDataSet with the current TimesheetLineActualDataSet */     
    #"Appended AssignmentTimePhased" = Table.Combine({#"Removed Columns", AssignmentTimephasedDataSet}),

    /* Inner Join Merge with the Resources table. (non-timesheet resources are included in this table, so long as they are enterprise, local resources on projects will show as generic local resource or unassigned resource)*/
    #"Merged Resources" = Table.NestedJoin(#"Appended AssignmentTimePhased", {"ResourceId"}, Resources, {"ResourceId"}, "Resources", JoinKind.Inner),

    /* Expand the Resources table column to retrieve Resource information */ 
    #"Expanded Resources" = Table.ExpandTableColumn(#"Merged Resources", "Resources", {"ResourceName", "ResourceCostCenter1"}, {"ResourceName", "ResourceCostCenter1"}),
    
    /* Inner Join Merge !!!on the!!! Projects table. */
    #"Merged Projects" = Table.NestedJoin(#"Expanded Resources", {"ProjectId"}, Projects, {"ProjectId"}, "Projects", JoinKind.Inner),
    
    /* Expand the Projects table column to retrieve Project information */      
    #"Expanded Projects" = Table.ExpandTableColumn(#"Merged Projects", "Projects", {"RawProjectName", "CapProject", "ITProduct", "ProjectCostCenter", "TimeType"}, {"RawProjectName", "CapProject", "ITProduct", "ProjectCostCenter", "TimeType"}),
    
    /* Inner Join Merge with the Tasks table */
    #"Merged Tasks" = Table.NestedJoin(#"Expanded Projects", {"TaskId"}, Tasks, {"TaskId"}, "Tasks", JoinKind.LeftOuter),
    
    /* Expand the Tasks table column to retrieve Task information */     
    #"Expanded Tasks" = Table.ExpandTableColumn(#"Merged Tasks", "Tasks", {"TaskName", "CapTask", "TaskCostCenter", "TaskProduct"}, {"TaskName", "CapTask", "TaskCostCenter", "TaskProduct"}),
    
    /* Inner Join Merge on the TimeSet table */    
    #"Merged TimeSet" = Table.NestedJoin(#"Expanded Tasks", {"TimeByDay"}, TimeSet, {"TimeByDay"}, "TimeSet", JoinKind.Inner),

    /* Expand the TimeSet table column to retrieve Time Periods and Fiscal Year */ 
    #"Expanded TimeSet" = Table.ExpandTableColumn(#"Merged TimeSet", "TimeSet", {"FiscalPeriod", "FiscalYear"}, {"FiscalPeriod", "FiscalYear"}),

    /* Outer Join Merge with the ResourceCCRates table in order to calculate rates and actual costs*/
    #"Merged ResourceCCRates" = Table.NestedJoin(#"Expanded TimeSet", {"ResourceName"}, ResourceCCRates, {"Resource"}, "ResourceCCRates", JoinKind.LeftOuter),

    /* Expand the ResoureCCRates table to retrieve Work CC's, Start/End Dates, and Rates */ 
    #"Expanded ResourceCCRates" = Table.ExpandTableColumn(#"Merged ResourceCCRates", "ResourceCCRates", {"StartDate", "WorkCC", "Rate", "EndDate"}, {"StartDate", "WorkCC", "Rate", "EndDate"}),

    /* Added a custom calculated column based on the [Resource], [StartDate], and [EndDate] to identify whether the row of data falls within rate calculation table -- #(lf)
        Because earlier we would duplicate rows if the resource had a change in Work CC or Rate within the Report Time Frame Parameters ([ReportStartDate] - [ReportFinishDate]). #(lf)
        This would allow us to use this custom column to pick the correct [Rate] and [WorkCC] for the given row [Date] --- The subsequent step would then filter out the row that does not fall within the rate calculation table (ResourceCCRates)*/  
    #"Added Correct CC & Rate" = Table.AddColumn(#"Expanded ResourceCCRates", "Correct CC & Rate", each if([TimeByDay] >= [StartDate]) and ([TimeByDay] < [EndDate]) and ([Rate] <> null) then 1 else if ([Rate] <> null) then 0 else null),
    
    /* Filter out the incorrect row, which does NOT fall within the rate calculation */    
    #"Filtered Correct CC & Rate" = Table.SelectRows(#"Added Correct CC & Rate", each ([#"Correct CC & Rate"] = 1)),
    
    /* Replace null [WorkCC] values with the [ResourceCostCenter1] value from the Resources table we merged -- 
        Since the rate calculation does not include Non-Timesheet Resources, Non-Timesheet Resources would reference the Resource tables for [WorkCC] and the [Rate] (which occurs in the subsequent step) */    
    #"Added Res. WorkCC if CC blank" = Table.AddColumn(#"Filtered Correct CC & Rate", "WorkCC.1", each if [WorkCC] = null then [ResourceCostCenter1] else [WorkCC]),
    
    /* Removed both the Old Work CC and Old Rate, since we don't need anymore */ 
    #"Removed Old WorkCC" = Table.RemoveColumns(#"Added Res. WorkCC if CC blank",{"WorkCC"}),

    /* Rename the new WorkCC */
    #"Renamed WorkCC" = Table.RenameColumns(#"Removed Old WorkCC",{{"WorkCC.1", "WorkCC"}}),

    /* Changed WorkCC to number for Actual Cost Calculation */
    #"Changed Rate to Decimal" = Table.TransformColumnTypes(#"Renamed WorkCC",{{"Rate", type number}}),

    /* Added a calculated column named [RawCostCenter] that retrieves the [TaskCostCenter] unless null, then it retrieves the [ProjectCostCenter] -- essentially this is the cost center that the work is directed at but not always charged to. */
    #"Added RawChargedCostCenter" = Table.AddColumn(#"Changed Rate to Decimal", "RawChargedCostCenter", each if [TaskCostCenter] is null then [ProjectCostCenter] else [TaskCostCenter]),
    /* Added a calculated column named [ChargedCostCenter], which is the cost center that the work is actually charged to. This is calculated with the conditions below for [WorkCC], [ResourceName], and [RawCostCenter] 
    IMPORTANT: This one of the most date sensitive steps and it needs to be monitor frequently as this will need changes when new timesheet users are added, relevant Cost Centers change, or when the chargeback model changes*/
    #"Added ChargedCostCenter" = Expression.Evaluate(ChargedCostCenterCalculation, 
        [
            /* Setting References for Expression.Evaluate to work within the appropriate context */ 
            #"SourceTable" = #"Added RawChargedCostCenter",  /* SourceTable is the prior step before #"Add ObjectCodeDesc" */
            #"Table.AddColumn" = Table.AddColumn,
            #"List.PositionOf" = List.PositionOf
        ]
        ),
    /* Filtered out Non-Billable Charged CC's since this is a Cost Report for Finance */
    #"Filtered Billable CC's" = Table.SelectRows(#"Added ChargedCostCenter", each ([ChargedCostCenter] <> "776" and [ChargedCostCenter] <> "777")),

    /* Outer Join Merge with the AEDCostCenter custom table to retrieve AED's and Cost Center names. Joined on the ChargedCostCenter */ 
    #"Merged AEDCostCenter" = Table.NestedJoin(#"Filtered Billable CC's", {"ChargedCostCenter"}, AEDCostCenter, {"CostCenter"}, "AEDCostCenter", JoinKind.LeftOuter),
    
    /* Expand the AEDCostCenter custom table column to retrieve ChargedAED and ChargedCostCenter Description. (need to make sure AED Custom Table is updated with the newest AED and CC Descriptions)  */    
    #"Expanded AEDCostCenter" = Table.ExpandTableColumn(#"Merged AEDCostCenter", "AEDCostCenter", {"AEDDesc", "CostCenterDesc"}, {"ChargedAEDDesc", "ChargedCostCenterDesc"}),
    
    /* Added a calculated column named [ProjectName] that would pull the [RawProjectName] if the [TimeType] = "Project" else would return the [TaskName] */
    #"Added ProjectName" = Table.AddColumn(#"Expanded AEDCostCenter", "ProjectName", each if [TimeType] <> "Project" then [TaskName] else [RawProjectName]),
    
    /* Calculated the Actual Cost based on the Rate retrieved from the ResourceCCRates Table and Hours */
    #"Added ActualCost" = Table.AddColumn(#"Added ProjectName", "ActualCost", each if [AssignmentActualCost] is null then [Hours] * [Rate] else [AssignmentActualCost], Number.Type),
    
    /* Added a calculated column named [Product] that would take the [ITProduct] if [TaskProduct] is null, otherwise retrieve the [TaskProduct] */
    #"Added Product" = Table.AddColumn(#"Added ActualCost", "Product", each if [TaskProduct] is null then [ITProduct] else [TaskProduct]),
    
    /* Removes columns no longer needed. We want to limit the columns to only the ones we need. IMPORTANT: There are necessary calculated DAX columns that are performed in the Data Model once loaded*/    
    #"Removed Columns No Longer Needed" = Table.RemoveColumns(#"Added Product",{"AssignmentActualCost", "ResourceCostCenter1", "ITProduct", "TaskProduct", "StartDate", "EndDate", "Correct CC & Rate", "ResourceId", "ProjectId", "TaskId", "RawChargedCostCenter"}),
    #"Filtered 0 ActualCost" = Table.SelectRows(#"Removed Columns No Longer Needed", each ([ActualCost] <> 0)),
    #"Added ObjectCodeDesc" = Expression.Evaluate(ObjectCodeCalculation, 
        [
            /* Setting References for Expression.Evaluate to work within the appropriate context */ 
            #"SourceTable" = #"Filtered 0 ActualCost",  /* SourceTable is the prior step before #"Add ObjectCodeDesc" */
            #"Table.AddColumn" = Table.AddColumn,
            #"List.PositionOf" = List.PositionOf
        ]
        ),
    #"Added ObjectCode" = Table.AddColumn(#"Added ObjectCodeDesc", "ObjectCode", each Text.Start([ObjectCodeDesc], 5))
in
    #"Added ObjectCode"
```
