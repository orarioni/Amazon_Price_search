$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Add()
$sh = $wb.Worksheets.Item(1)
$sh.Cells.Item(1,2).Value2 = 'JAN'
$sample = @('4901234567890','4547274043587','4547274044041','4547274044065','4547274044072')
$i=2
foreach($j in $sample){ $sh.Cells.Item($i,2).Value2 = $j; $i++ }
$wb.SaveAs((Join-Path $PSScriptRoot '..\data\input.xlsx'))
$wb.Close($false)
$excel.Quit()
[Runtime.InteropServices.Marshal]::ReleaseComObject($sh)|Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($wb)|Out-Null
[Runtime.InteropServices.Marshal]::ReleaseComObject($excel)|Out-Null
