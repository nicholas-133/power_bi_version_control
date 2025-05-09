# power_query_version_control


### Version Control:
  Centralized power query language scripts that serve as expressions for multiple Power BI and Excel Reports.

### Goal: 
  Create one stored script that all Power BI and Excel cost reports reference -- so only one modification needs to made for all similar reports. 

### Solution / Result: 
  The solution consisted of manipulating 2 lists in M code that serve as conditions and results similar to a Switch Function. If this listed condition is met, then the listed result will be selected. This is a workaround to using the DAX Switch statement for each individual report. These expressions were stored and secured as centralized .txt files in SharePoint. There are two queries in the power query editor "ObjectCodeCalculation" and "ChargedCostCenterCalculation" that retrieve the two .txt file expressions and stored it in the editor as two big blocks of text. The Expression.Evaluate() function was then used within the main fact table [TimesheetLineActualDataSet] to insert these blocks of text as PowerQuery steps. Therefor, allowing IT to modify critical query steps in multiple reports at once -- improving data integrity, consistency, and efficiency.
