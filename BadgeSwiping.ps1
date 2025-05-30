# PowerShell script to process badge swiping files

# Set file paths
$inputFile = "C:\Users\Rputhoff\Documents\Badge_Swiping\Cardholders with Active Cards ().xls"
$validFile = "C:\Users\Rputhoff\Documents\Badge_Swiping\Midmark Production Badge Swiping Group.xlsx"
$outputFile = "C:\Users\Rputhoff\Documents\Badge_Swiping\CS_badge_import.txt"

# Check if input file exists
if (!(Test-Path $inputFile)) {
    Write-Error "Input file not found: $inputFile"
    exit 1
}

# Check file extension
if ($inputFile -notmatch "\.xls$|\.xlsx$") {
    Write-Error "Invalid file format: $inputFile. Expected .xls or .xlsx"
    exit 1
}

# Function to read Excel file into a DataTable
function Import-ExcelFile {
    param (
        [string]$Path
    )
    if ($Path -like "*.xlsx") {
        if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
            try {
                Install-Module -Name ImportExcel -Force -Scope CurrentUser -ErrorAction Stop
            } catch {
                # Fallback to COM if ImportExcel is not available
                Write-Warning "ImportExcel module not available. Falling back to COM for .xlsx file."
                $excel = New-Object -ComObject Excel.Application
                $excel.Visible = $false
                $workbook = $excel.Workbooks.Open($Path)
                $worksheet = $workbook.Worksheets.Item(1)
                $usedRange = $worksheet.UsedRange
                $data = $usedRange.Value2
                $rowCount = $usedRange.Rows.Count
                $colCount = $usedRange.Columns.Count
                $headers = @()
                for ($c = 1; $c -le $colCount; $c++) {
                    $headers += $data[1, $c]
                }
                $rows = @()
                for ($r = 2; $r -le $rowCount; $r++) {
                    $row = @{
                    }
                    for ($c = 1; $c -le $colCount; $c++) {
                        $row[$headers[$c-1]] = $data[$r, $c]
                    }
                    $rows += [PSCustomObject]$row
                }
                $workbook.Close($false)
                $excel.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
                return $rows
            }
        }
        Import-Module ImportExcel
        return Import-Excel -Path $Path
    } else {
        # Use COM for .xls
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($Path)
        $worksheet = $workbook.Worksheets.Item(1)
        $usedRange = $worksheet.UsedRange
        $data = $usedRange.Value2
        $rowCount = $usedRange.Rows.Count
        $colCount = $usedRange.Columns.Count
        $headers = @()
        for ($c = 1; $c -le $colCount; $c++) {
            $headers += $data[1, $c]
        }
        $rows = @()
        for ($r = 2; $r -le $rowCount; $r++) {
            $row = @{
            }
            for ($c = 1; $c -le $colCount; $c++) {
                $row[$headers[$c-1]] = $data[$r, $c]
            }
            $rows += [PSCustomObject]$row
        }
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        return $rows
    }
}

# Read input and valid EmpID files
$data = Import-ExcelFile -Path $inputFile
$validEmpIDRows = Import-ExcelFile -Path $validFile
$validEmpIDs = $validEmpIDRows | Select-Object -ExpandProperty 'EmpID' | Where-Object { $_ } | ForEach-Object { $_.ToString() }

# Remove specified columns
$columnsToDelete = @("Last Name", "First Name", "Company Name", "Access Group", "Imprint", "Card Status", "CardholderType")
foreach ($col in $columnsToDelete) {
    if ($data[0].PSObject.Properties.Name -contains $col) {
        $data | ForEach-Object { $_.PSObject.Properties.Remove($col) }
    }
}

# Rename columns
foreach ($row in $data) {
    if ($row.PSObject.Properties.Name -contains "Middle Name") {
        $row | Add-Member -NotePropertyName "EmpID" -NotePropertyValue $row."Middle Name"
        $row.PSObject.Properties.Remove("Middle Name")
    }
    if ($row.PSObject.Properties.Name -contains "Cardnumber") {
        $row | Add-Member -NotePropertyName "Badge_num" -NotePropertyValue $row."Cardnumber"
        $row.PSObject.Properties.Remove("Cardnumber")
    }
}

# Remove rows where EmpID is null
$data = $data | Where-Object { $_.EmpID -ne $null -and $_.EmpID -ne "" }

# Keep only rows with valid EmpID
$data = $data | Where-Object { $validEmpIDs -contains $_.EmpID.ToString() }

# Add leading zeros to Badge_num if not 10 digits
foreach ($row in $data) {
    if ($row.Badge_num) {
        $row.Badge_num = $row.Badge_num.ToString().PadLeft(10, '0')
    }
}

# Export to CSV
$data | Export-Csv -Path $outputFile -NoTypeInformation -Encoding UTF8

# Remove all double quotes from the output file
(Get-Content $outputFile) -replace '"', '' | Set-Content $outputFile -Encoding UTF8

Write-Host "Processing complete. Output saved to $outputFile"