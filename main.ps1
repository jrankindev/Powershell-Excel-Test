#i want to see if i can open an existing excel document and make changes to it with powershell without the ImportExcel module

#clear screen
Clear-Host

$path = "D:\Dropbox\DEV\excel\Powershell-Excel-Test\test.xlsx"

$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($path)
$page = "Sheet1"
$ws = $workbook.worksheets | Where-Object {$_.Name -eq $page}

$cells = $ws.Cells

$row = 7
$col = 3

$cells.item($row,$col) = "blah"

$cells.item($row,$col)

$workbook.Close($true)
$excel.quit()

$killID = ((get-process excel | Select-Object MainWindowTitle, ID, StartTime | Sort-Object StartTime)[-1]).Id
Stop-Process -Id $killID