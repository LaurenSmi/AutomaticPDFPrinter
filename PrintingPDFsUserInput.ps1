# OPEN EXCEL
$objExcel = New-Object -ComObject Excel.Application

# GET WORKBOOK ADDRESS
$directory = Read-Host -prompt 'Enter the directory to your Excel file'

# REMOVE QUOTATIONS FROM FILE ADDRESS
if($directory[0] -eq '"')
{
    $directory = $directory.Substring(1,$directory.length-2)
}

# OPEN WORKBOOK
try
{
    $workBook = $objExcel.Workbooks.Open($directory)

}
catch
{
    Write-Host 'File could not be found.' -ForegroundColor Red
    Start-Sleep -Seconds 15
    exit
}

# GET WORKSHEET NAME
$workSheetName = Read-Host -prompt 'Enter the name of your Excel worksheet. If you are unsure of its name, please enter n'

# PRINT LIST OF WORKSHEET NAMES
if($workSheetName -eq 'n')
{
    $workBook.Sheets| fl Name, index
    $workSheetName = Read-Host -prompt 'Enter name of your Excel worksheet'
}

# OPEN WORKSHEET CONTAINING PARTS LIST
try
{
    $workSheet = $workBook.Sheets.Item($workSheetName)

}
catch
{
    Write-Host 'Worksheet could not be found.' -ForegroundColor Red
    Start-Sleep -Seconds 15
    exit
}

# GET FIRST ROW AND COLUMN NUMBER
$totalRows = $workSheet.UsedRange.Rows.Count
$firstLine = Read-Host -prompt 'Does your list of parts start on cell A1 (y/n)? '
if($firstLine -eq 'y')
{
    $firstRow = 1
    $firstCol = 1
}

else
{
    [uint16]$firstRow = Read-Host -prompt 'Enter row # to start'
    [uint16]$firstCol = Read-Host -prompt 'Enter col # (as a number)'
}

# GET ADDRESS FOR OUTPUT FILE
$newDirectory = Read-Host -prompt 'Enter the directory to output missing parts file'
if($newDirectory[0] -eq '"')
{
    $newDirectory = $newDirectory.Substring(1,$newDirectory.length-2)
}

# CREATE NEW OUTPUT FILE
$newDirectory = $newDirectory + "\MissingParts.txt"
New-Item $newDirectory

for($i = $firstRow; $i-le $totalRows;$i++)
{
    # GET PART FROM EXCEL WORKSHEET
    $fileName = $WorkSheet.cells.Item($i,$firstCol).text
    $length = $fileName.Length

    # CONVERT PART NUMBER TO A PDF FILE NAME
    $fileName = $fileName + ".pdf"
    $found = $false
    $here = $false

    # REMOVE DASH FROM PART NAME
    if($length -gt 8)
    {
        $i = $length-1
        while($i -ge 0 -and !$here)
        {
            if($fileName[$i] -eq '-')
            {
                $here = $true
                $fileName = $fileName.Substring(0,$i)
                $length = $fileName.Length
                $fileName = $fileName + '.pdf'
            }
            $i--
        }
    }

    while($length -le 8 -and !$found)
    {
        try
        {
            # TRY TO PRINT FILE - IF SUCCESSFUL BREAK LOOP
            Start-Process -FilePath "G:\Drawings\AutoDrawingPrinter\DupDrawingsForPrinting\$fileName" -Verb print
            if($?)
            {
                $found = $true
                Write-Host "$fileName printed." -ForegroundColor Green
            }

        }
        catch
        {
             # TRY ADDING A ZERO TO THE BEGINNING OF PART NUMBER
             $fileName = '0'+$fileName
             $length++
        }
    }
    
    # OUTPUT MISSING PARTS TO CONSOLE AND ADD THEM TO THE OUTPUT FILE
    if(!$found)
    {
        $partNumber = $WorkSheet.cells.Item($i,$firstCol).text
        Write-Host "$partNumber not found." -ForegroundColor Red
        Add-Content $newDirectory $partNumber   
    }
}

$workBook.Close()
$objExcel.Quit()

Start-Sleep -Seconds 10