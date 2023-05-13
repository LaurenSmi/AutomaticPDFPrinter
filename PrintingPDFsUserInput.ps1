# THIS SCRIPT AUTOMATICALLY PRINTS DRAWINGS FROM AN EXCEL SHEET OF PART NUMBERS
# WRITTEN BY LAUREN SMILLIE
# CREATED FEB 1, 2023
# LAST UPDATED APR 27, 2023

# REMOVES ZEROS FROM BEGINNING OF PART NUMBER
function removeZeros($name)
{
    $i=0
    while($name[$i] -eq '0')
    {
        $i++
    }

    if($i -ne 0)
    {
        $name = $name.Substring($i,$name.Length-$i)
    }
    return $name
}

# REMOVES DASHES AND ADDS ZEROS TO BEGINNING OF PART NUMBER
function modNumber($name)
{
    if($name.contains('-'))
    {
        $here = $False
        $i = $name.Length-1
        
        while($i -ge 0 -and !$here)
        {
            if($name[$i] -eq '-' -and $i -ne 2)

            {
                $here = $true
                $name = $name.Substring(0,$i) + ".pdf"
                return $name
            }
            $i--
        }
    }
    else
    {
        $name = '0'+$name
    }
    return $name
}

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
$workSheetName = Read-Host -prompt "`nEnter the name of your Excel worksheet. If you are unsure of its name, `nplease enter n"

# PRINT LIST OF WORKSHEET NAMES
if($workSheetName -eq 'n')
{
    $workBook.Sheets| fl Name, index
    $workSheetName = Read-Host -prompt "`nEnter name of your Excel worksheet"
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
$firstLine = Read-Host -prompt "`nDoes your list of parts start on cell A1 (y/n)? "

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
$newDirectory = Read-Host -prompt "`nEnter the directory to output missing parts file"

if($newDirectory[0] -eq '"')
{
    $newDirectory = $newDirectory.Substring(1,$newDirectory.length-2)
}

if($newDirectory[$newDirectory.Length-1] -eq '\')
{
    $newDirectory = $newDirectory + "MissingParts"
}

else
{
    $newDirectory = $newDirectory + "\MissingParts"
}

$baseLength = $newDirectory.Length
$directoryMade = $False
$counter =1

# MAKE DIRECTORY
while (!$directoryMade)
{  
    try
    {
        $newDirectory = $newDirectory + ".txt"
        New-Item $newDirectory -ErrorAction Stop
        if($?)
        {
            $directoryMade = $True
            $date = Get-Date
            Add-Content $newDirectory $date
        }
    }

    catch [System.Exception]
    {
        if([System.IO.File]::Exists($newDirectory))
        {
            $newDirectory = $newDirectory.Substring(0,$baseLength) + $counter
            $counter++
        }
        
        else
        {
            Write-Host "`nInvalid directory" -ForegroundColor Red
            $newDirectory = Read-Host -prompt 'Enter the directory to output missing parts file'
            
            # REMOVE QUOTATIONS AROUND DIRECTORY
            if($newDirectory[0] -eq '"')
            {
                $newDirectory = $newDirectory.Substring(1,$newDirectory.length-2)
            }
            if($newDirectory[$newDirectory.Length-1] -eq '\')
            {
                $newDirectory = $newDirectory + "MissingParts"
            }
            else
            {
                $newDirectory = $newDirectory + "\MissingParts"
            }
            
            $baseLength = $newDirectory.Length
            $counter=1
        }
    }
}

for($i = $firstRow; $i-le $totalRows;$i++)
{
    # GET PART FROM EXCEL WORKSHEET
    $fileName = $WorkSheet.cells.Item($i,$firstCol).text
    
    if(!$fileName.Equals(""))
    {
        $fileName = removeZeros($fileName)
        $length = $fileName.Length

        $drwngDirec = "G:\Drawings\MicroFilms\"

        # CONVERT PART NUMBER TO A PDF FILE NAME
        $fileName = $fileName + ".pdf"
        $found = $false

        while($length -le 7 -or $fileName[0] -gt 57 -and !$found)
        {
            # LOCATE FILE IN FOLDER
            try
            {
                $filePath = Get-ChildItem $drwngDirec -Filter $fileName -Recurse -Name
                $filePath = $drwngDirec+$filePath
                
                if($? -and !$filePath.Equals($drwngDirec))
                {
                    # PRINT FILE
                    try
                    {
                        Start-Process -WindowStyle Hidden -FilePath $filePath -Verb print
                        if($?)
                        {
                            $found = $true
                            Write-Host "$fileName printed from $filePath." -ForegroundColor Green
                            Start-Sleep -Seconds 3
                            break
                        }
                    }
                    catch
                    {
                        Write-Host "$fileName was found at $filePath, but there was an error printing" -ForegroundColor Red
                        break
                    }
                }
                else
                {
                    Write-Host "$fileName could not be found in the $drwngDirec directory." -ForegroundColor Red
                    
                    # REMOVE DASH FROM PART NAME
                    $fileName = modNumber($fileName)
                    $length = $fileName.Length - ".pdf".Length
                }
            }
            catch
            {   
                # REMOVE DASH FROM PART NAME
                $fileName = modNumber($fileName)
                $length = $fileName.Length - ".pdf".Length
            }
        }

        # OUTPUT MISSING PARTS TO CONSOLE AND ADD THEM TO THE OUTPUT FILE
        if(!$found)
        {
            $partNumber = $WorkSheet.cells.Item($i,$firstCol).text
            Add-Content $newDirectory $partNumber  
        }
    }
}
Write-Host "`nA list of parts that weren't printed can be found at $newDirectory" -ForegroundColor Magenta

$workBook.Close()
$objExcel.Quit()

Start-Sleep -Seconds 10
