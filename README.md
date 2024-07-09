# VBA-Challenge
Module 2 Challenge - Use VBA scripting to analyze generated stock market data.

## Instructions

### New Excel File

On a new Excel file that meets the challenge criteria, open the Developer
ribbon, and click on `Visual Basic`.  In the VBA editor that pops up, select
`File` > `Import File` from the menu.  Navigate to the
`ProcessQuarterlyTickerData.bas` file in the repo, and click `Open`.  Close the
VBA Editor.

Back in Excel, in the Developer ribbon, click on `Macros`.  In the pop-up
select `ProcessAllQuarterlyWorksheets` to process all worksheets, or
`ProcessSingleQuarterlyWorksheet` to process the active worksheet, then click
the `Run` button.

### Submitted Excel File

Open the `Multiple_year_stock_data.xlsm` workbook.  In the Developer ribbon,
click on `Macros`.

If you wish to clear all workbooks to watch a before and after, run the
`DebugClearAllWorksheets` macro.  Run the `ProcessAllQuarterlyWorksheets` macro
to process all worksheets.  Select an individual worksheet and run 
`ProcessSingleQuarterlyWorksheet` instead if you want to test the macro on an
individual worksheet.

## Files

* `README.md`: this file
* `ProcessQuarterlyTickerData.bas`: VBA file with all required macros
* `Multiple_year_stock_data.xlsm`: Macro enabled Excel file with macros included
* `processed-worksheet.png`: A screenshot of a single worksheet after processing
