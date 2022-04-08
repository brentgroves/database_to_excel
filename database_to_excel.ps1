#  https://docs.microsoft.com/en-us/answers/questions/68953/generate-excel-spreadsheet-xls-in-sql-server.html
 Write-Host –NoNewLine “Processing.................Do not close"
 ## - Get SQL Server Table data:
 $SQLServer = 'mgsqlmi.public.48d444e7f69b.dat';
 $Database = 'mgdw';
 # $today = (get-date).ToString("dd-MM-yyyy")
 $ExportFile = "C:\trial_balance_multi_level_2022_03.xlsx"
 # $ExportFile = "\\zoom-nas\Shared_Documents\FC Folder\Despatch\Brexit Files\Landmark\Landmark "+$today+".xlsx"
 ##$SqlQuery = @'EXEC [zoomfs].[LandMarkGlobalExport]'@ ;
    
 ## - Connect to SQL Server using non-SMO class 'System.Data':
 $SqlConnection = New-Object System.Data.SqlClient.SqlConnection;
 $SqlConnection.ConnectionString = `
 "Server = $SQLServer; Database = $Database; Integrated Security = True";
    
 $SqlCmd = New-Object System.Data.SqlClient.SqlCommand;
 Report.trial_balance
@start_period int,
@end_period int 
$SqlCmd.CommandText = $("EXEC [Report].[trial_balance] 202203,202203");
# $SqlCmd.CommandText = $("EXEC [scyhema].[LandMarkGlobalExport]");
 $SqlCmd.Connection = $SqlConnection;
 $SqlCmd.CommandTimeout = 0;
    
 ## - Extract and build the SQL data object '$DataSetTable':
 $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter;
 $SqlAdapter.SelectCommand = $SqlCmd;
 $DataSet = New-Object System.Data.DataSet;
 $SqlAdapter.Fill($DataSet);
 $DataSetTable = $DataSet.Tables["Table"];
    
    
 ## ---------- Working with Excel ---------- ##
    
 ## - Create an Excel Application instance:
 $xlsObj = New-Object -ComObject Excel.Application;
    
 ## - Create new Workbook and Sheet (Visible = 1 / 0 not visible)
 $xlsObj.Visible = 0;
 $xlsWb = $xlsobj.Workbooks.Add();
 $xlsSh = $xlsWb.Worksheets.item(1);
 $xlsSh.columns.item('A').NumberFormat = "@"
 $xlsSh.columns.item('P').NumberFormat = "@"
 ## - Copy entire table to the clipboard as tab delimited CSV
 $DataSetTable | ConvertTo-Csv -NoType -Del "`t" | Clip
    
 ## - Paste table to Excel
 $xlsObj.ActiveCell.PasteSpecial() | Out-Null
    
 ## - Set columns to auto-fit width
 $xlsObj.ActiveSheet.UsedRange.Columns|%{$_.AutoFit()|Out-Null}
    
 ## - Saving Excel file - if the file exist do delete then save
 $xlsFile = $ExportFile;
    
 if (Test-Path $xlsFile)
 {
 Remove-Item $xlsFile
 $xlsObj.ActiveWorkbook.SaveAs($xlsFile);
 }
 else
 {
 $xlsObj.ActiveWorkbook.SaveAs($xlsFile);
 };
    
 ## Quit Excel and Terminate Excel Application process:
 $xlsObj.Quit(); (Get-Process Excel*) | foreach ($_) { $_.kill() };