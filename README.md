# power_query_version_control
Centralized mashup language script for implementing version control of multiple Power BI and Excel Reports.

Goal: Create one stored script that all Power BI and Excel cost reports reference -- so only one modification needs to made for all similar reports. 

Solution: The solution consisted of manipulating lists in M code to derive a switch function with conditions and results. If this particular condition is met, then the listed result will be selected. This is a workaround to using the DAX Switch statement for each individual report. 
