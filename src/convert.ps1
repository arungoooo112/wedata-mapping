# Function to convert XLS to XLSX
function Convert-XlsToXlsx {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({Test-Path $_ -PathType Leaf})]
        [string]$xlsFilePath
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    $workbook = $excel.Workbooks.Open($xlsFilePath)
    $xlsxFilePath = [System.IO.Path]::ChangeExtension($xlsFilePath, "xlsx")
    $workbook.SaveAs($xlsxFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook)
    $workbook.Close()

    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}

# Function to convert XLSX to XLS
function Convert-XlsxToXls {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({Test-Path $_ -PathType Leaf})]
        [string]$xlsxFilePath
    )

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false

    $workbook = $excel.Workbooks.Open($xlsxFilePath)
    $xlsFilePath = [System.IO.Path]::ChangeExtension($xlsxFilePath, "xls")
    $workbook.SaveAs($xlsFilePath, [Microsoft.Office.Interop.Excel.XlFileFormat]::xlExcel8)
    $workbook.Close()

    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
}
