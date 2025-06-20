# Check Excel file content
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false

try {
    $workbook = $excel.Workbooks.Open("C:\temp\AzureDevOpsExport_ProjectImport_20250619_232756.xlsx")
    $worksheet = $workbook.Worksheets["Tasks"]
    
    Write-Host "Column headers:"
    for ($col = 1; $col -le 10; $col++) {
        $header = $worksheet.Cells.Item(1, $col).Value2
        Write-Host "Column ${col}: $header"
    }
    
    Write-Host "`nFirst 10 data rows:"
    for ($i = 2; $i -le 11; $i++) {
        Write-Host "Row ${i}:"
        for ($col = 1; $col -le 10; $col++) {
            $value = $worksheet.Cells.Item($i, $col).Value2
            Write-Host "  Col ${col}: $value"
        }
        Write-Host ""
    }
    
} finally {
    $workbook.Close($false)
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
