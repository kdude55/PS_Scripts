####################################################################################################
## This is a work in Progress script. The perpose of this script is to take a list of computers   ##
## in an excell file test the connection twice. if the computer sends a response, it adds the     ##
## computer to an array called $online and if the computer does not give a response, it adds the  ##
## computer to an array called $offline and exports the contents of $offline to an excell file    ##
## called Offline_$today.xlsx The variable $today takes the output of Get-Date and converts it to ##
## yyyMMdd format.                                                                                ##
## Things I want to add to this script include Adding Multicore Processing,                       ##
## asking the user to input the ComputerList.XLSX file, and compatability with specific reports   ##
####################################################################################################

# declare excel variables
$user = $env:USERNAME
$excelPath = "C:\Users\$user\Documents\ComputerList.xlsx"
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Open($excelPath)
$sheet = $workbook.Sheets.Item(1)
$row = 1
$computers = @()

#Pull data from Excel dont pull in duplicate entries
while ($sheet.Cells.Item($row, 1).Text -ne "") {
    $hostname = $sheet.Cells.Item($row, 1).Text.Trim()
    if ($computers -notcontains $hostname) {
        $computers += $hostname
    }
    $row++ 
    
}

$workbook.Close($false)
$excel.Quit()

# Create the arays to sort the computers
$online = @()
$offline = @()

# Ping each computer
foreach ($computer in $computers) {
    Write-Host "Pinging " $computer "Please Wait..."
    if (Test-Connection -ComputerName $computer -Count 2 -Quiet) {
        Write-Host "ONLINE"
        $online += $computer
    } else {
        Write-Host $computer "OFFLINE"
        $offline += $computer
    }
}

# Export offline Array to Excel
$today = Get-Date -Format "yyyyMMdd"
$offlinePath = "C:\Users\$user\Documents\offline_$today.xlsx"
Write-Host "Exporting Data to" $offlinePath
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Add()
$sheet = $workbook.Sheets.Item(1)

for ($i = 0; $i -lt $offline.Count; $i++) {
    $sheet.Cells.Item($i + 1, 1) = $offline[$i]
}

$workbook.SaveAs($offlinePath)
$workbook.Close($false)
$excel.Quit()