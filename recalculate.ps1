$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$path = Resolve-Path "Hexalevel_Financial_Management_Dashboard.xlsx"
$wb = $excel.Workbooks.Open($path)
$excel.CalculateFull()
$wb.Save()
$wb.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
Write-Output "Excel fully calculated and saved successfully!"
