<# This tool takes a list of computers from an excel file, puts it in an array,
then pings each computer, if a response is recieved, the computer is added to 
the online array, if no respons is recieved, the computer is added to the 
offline array. The offline array is then sent to an excel file. The online array
will be used as input to remove specific files/applications on each computer in the
array. I am going to start with bing wallpaper.
#>

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

#function to delete folders
function Remove-BingWallpaper {
    param ($session, $hostname)

    try {
        $users = Invoke-Command -Session $session -ScriptBlock {
            Get-ChildItem -Path "C:\Users" -Directory | Select-Object -ExpandProperty Name
        }

        foreach ($user in $users) {
            $localPath = "C:\Users\$user\AppData\Local\Microsoft\WindowsApps\Microsoft.BingWallpaper*"
            $roamingPath = "C:\Users\$user\AppData\Roaming\Microsoft\WindowsApps\Microsoft.BingWallpaper*" 

            Invoke-Command -Session $session -ScriptBlock {
                param ($local, $roaming)
                Remove-Item -Path $local -Recurse -Force -ErrorAction SilentlyContinue
                Remove-Item -Path $roaming -Recurse -Force -ErrorAction SilentlyContinue
            } -ArgumentList $localPath, $roamingPath
        }

        Write-Host "Bing Wallpaper succesfully removed from $hostname"
    } catch {
        Write-Host "Trying to kill bing wallpaper and retry. Please wait..."
        Invoke-Command -Session $session -ScriptBlock {
            Get-Process BingWallpaper -ErrorAction SilentlyContinue | Stop-Process -Force
        }

        try {
            foreach ($user in $users) {
                $localPath = "C:\Users\$user\AppData\Local\Microsoft\WindowsApps\Microsoft.BingWallpaper*"
                $roamingPath = "C:\Users\$user\AppData\Roaming\Microsoft\WindowsApps\Microsoft.BingWallpaper*"

                Invoke-Command -Session $session -ScriptBlock {
                    param($local, $roaming)
                    Remove-Item -Path $local -Recurse -Force -ErrorAction SilentlyContinue
                    Remove-Item -Path $roaming -Recurse -Force -ErrorAction SilentlyContinue
                } -ArgumentList $localPath, $roamingPath
            }

            Write-Host "Killed and removed Bing Wallpaper"
        } catch {
            Write-Host "Failed to kill and remove Bing Wallpaper"
        }
    }
}

foreach ($hostname in $online) {
    try {
        $session = New-PSSession -ComputerName $hostname -ErrorAction Stop
        Remove-BingWallpaper -session $session -hostname $hostname
        Remove-PSSession $session
    } catch {
        Write-Host "Failed to connect to $hostname"
    }
}