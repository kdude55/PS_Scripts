<# This tool takes a list of computers from an excel file, puts it in an array,
then pings each computer, if a response is recieved, the computer is added to 
the online array, if no respons is recieved, the computer is added to the 
offline array. The offline array is then sent to an excel file. The online array
will be used as input to remove specific files/applications on each computer in the
array. I am going to start with bing wallpaper.
#>

# declare excel variables
$excelPath = Read-Host -Prompt "Please enter the path to your excel file(MAKE SURE YOU DONT HAVE QUOTATION MARKS!)"
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

# Export offline Array to CSV
$today = Get-Date -Format "yyyyMMdd"
$folder = Read-Host -Prompt "Where would you like me to save the CSV file?"
$offlinePath = Join-Path $folder "Offline_$today.csv"

Write-Host "Exporting data to $offlinePath"

$offline |
    ForEach-Object { [PSCustomObject]@{ ComputerName = $_ } } |
    Export-Csv -Path $offlinePath -NoTypeInformation

#function to delete Bing Wallpaper
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

                if (Test-Path -Path $local) {Remove-Item -Path $local -Recurse -Force -ErrorAction SilentlyContinue}
                if (Test-Path -Path $roaming) {Remove-Item -Path $roaming -Recurse -Force -ErrorAction SilentlyContinue}
            } -ArgumentList $localPath, $roamingPath
        }

        Write-Host "Bing Wallpaper successfully removed from $hostname"
    } catch {
        Write-Host "Trying to kill bing wallpaper and retry. Please wait..."
        Invoke-Command -Session $session -ScriptBlock {
            Get-Process BingWallpaper -ErrorAction SilentlyContinue | Stop-Process -Force
        }
        #Try to delete after killing Bing Wallpaper
        try {
            foreach ($user in $users) {
                $localPath = "C:\Users\$user\AppData\Local\Microsoft\WindowsApps\Microsoft.BingWallpaper*"
                $roamingPath = "C:\Users\$user\AppData\Roaming\Microsoft\WindowsApps\Microsoft.BingWallpaper*"

                Invoke-Command -Session $session -ScriptBlock {
                    param ($local, $roaming)

                    if (Test-Path -Path $local) {Remove-Item -Path $local -Recurse -Force -ErrorAction SilentlyContinue}
                    if (Test-Path -Path $roaming) {Remove-Item -Path $roaming -Recurse -Force -ErrorAction SilentlyContinue}
                } -ArgumentList $localPath, $roamingPath
            }

            Write-Host "Killed and removed Bing Wallpaper"
        } catch {
            Write-Host $_
            Write-Host "Failed to kill and remove Bing Wallpaper"
        }
    }
}

#Function to delete Zoom
function Remove-Zoom {
    param ($session, $hostname)

    try {
        $users = Invoke-Command -Session $session -ScriptBlock {
            Get-ChildItem -Path "C:\Users" -Directory | Select-Object -ExpandProperty Name
        }

        foreach ($user in $users) {
            $localPath = "C:\Users\$user\AppData\Local\Zoom"
            $roamingPath = "C:\Users\$user\AppData\Roaming\Zoom" 

            Invoke-Command -Session $session -ScriptBlock {
                param ($local, $roaming)

                if (Test-Path -Path $local) {Remove-Item -Path $local -Recurse -Force -ErrorAction SilentlyContinue}
                if (Test-Path -Path $roaming) {Remove-Item -Path $roaming -Recurse -Force -ErrorAction SilentlyContinue}
            } -ArgumentList $localPath, $roamingPath
        }

        Write-Host "Zoom successfully removed from $hostname"
    } catch {
        Write-Host "Trying to kill Zoom and retry. Please wait..."
        Invoke-Command -Session $session -ScriptBlock {
            Get-Process Zoom -ErrorAction SilentlyContinue | Stop-Process -Force
        }
        #Try to delete after killing Zoom
        try {
            foreach ($user in $users) {
                $localPath = "C:\Users\$user\AppData\Local\Zoom"
                $roamingPath = "C:\Users\$user\AppData\Roaming\Zoom"

                Invoke-Command -Session $session -ScriptBlock {
                    param ($local, $roaming)

                    if (Test-Path -Path $local) {Remove-Item -Path $local -Recurse -Force -ErrorAction SilentlyContinue}
                    if (Test-Path -Path $roaming) {Remove-Item -Path $roaming -Recurse -Force -ErrorAction SilentlyContinue}
                } -ArgumentList $localPath, $roamingPath
            }

            Write-Host "Killed and removed Zoom"
        } catch {
            Write-Host $_
            Write-Host "Failed to kill and remove Zoom"
        }
    }
}

<#Function to delete Dell Command Update Commented out to continue working on this
function Remove-DCU {
    param ($session, $hostname)

    $dcuPath = "C:\Program Files (x86)\Dell\CommandUpdate"
    
    try {
        Invoke-Command -Session $session -ScriptBlock {
            param($path)
            Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
        } -ArgumentList $dcuPath

        Write-Host "DCU successfully removed from $hostname"
    } catch {
        Write-Host "Trying to kill Dell services and retry. Please Wait..."
        Invoke-Command -Session $session -ScriptBlock {
            Get-Process Dell* -ErrorAction SilentlyContinue | Stop-Process -Force
        }
        try {
            Invoke-Command -Session $session -ScriptBlock {
                param($path)
                Remove-Item -Path $path -Recurse -Force -ErrorAction Stop
            } -ArgumentList $dcuPath

            Write-Host "DCU Successfully removed from host after killing Dell"
        } catch {
            Write-Host $_
            Write-Host "Failed to remove DCU"
        }
    }

    
}
#>

#Function to remove Spotify

function Remove-Spotify {
    param ($session, $hostname)

    try {
        $users = Invoke-Command -Session $session -ScriptBlock {
            Get-ChildItem -Path "C:\Users" -Directory | Select-Object -ExpandProperty Name
        }

        foreach ($user in $users) {
            $localPath = "C:\Users\$user\AppData\Local\Microsoft\WindowsApps\Spotify*"
            $roamingPath = "C:\Users\$user\AppData\Roaming\Microsoft\WindowsApps\Spotify*" 
            $downloadPath = "C:\Users\$user\Downloads\Spotify*"
            $programPath = "C:\Program Files\WindowsApps\Spotify*"

            Invoke-Command -Session $session -ScriptBlock {
                param ($local, $roaming, $download, $program)

                if (Test-Path -Path $local) {Remove-Item -Path $local -Recurse -Force -ErrorAction SilentlyContinue}
                if (Test-Path -Path $roaming) {Remove-Item -Path $roaming -Recurse -Force -ErrorAction SilentlyContinue}
                if (Test-Path -Path $download) {Remove-Item -Path $download -Recurse -Force -ErrorAction SilentlyContinue}
                if (Test-Path -Path $program) {Remove-Item -Path $program -Recurse -Force -ErrorAction SilentlyContinue}
            } -ArgumentList $localPath, $roamingPath, $downloadPath, $programPath
        }

        Write-Host "Spotify successfully removed from $hostname"
    } catch {
        Write-Host "Could not remove Spotify from ${hostname}:`n$($_.Exception.Message)"
        Write-Host "Trying to kill Spotify and retry. Please wait..."

        Invoke-Command -Session $session -ScriptBlock {
            Get-Process spotify* -ErrorAction SilentlyContinue | Stop-Process -Force
        }
        #Try to delete after killing Spotify
        try {
            foreach ($user in $users) {
                $localPath = "C:\Users\$user\AppData\Local\Microsoft\WindowsApps\Spotify*"
                $roamingPath = "C:\Users\$user\AppData\Roaming\Microsoft\WindowsApps\Spotify*"
                $downloadPath = "C:\Users\$user\Downloads\Spotify*"
                $programPath = "C:\Program Files\WindowsApps\Spotify*"

                Invoke-Command -Session $session -ScriptBlock {
                    param ($local, $roaming, $download, $program)

                    if (Test-Path -Path $local) {Remove-Item -Path $local -Recurse -Force -ErrorAction SilentlyContinue}
                    if (Test-Path -Path $roaming) {Remove-Item -Path $roaming -Recurse -Force -ErrorAction SilentlyContinue}
                    if (Test-Path -Path $download) {Remove-Item -Path $download -Recurse -Force -ErrorAction SilentlyContinue}
                    if (Test-Path -Path $program) {Remove-Item -Path $program -Recurse -Force -ErrorAction SilentlyContinue}
                } -ArgumentList $localPath, $roamingPath, $downloadPath, $programPath
            }

            Write-Host "Killed and removed Spotify"
        } catch {
            Write-Host "unable to kill and remove Spotify from ${hostname}:`n$($_.Exception.Message)"
            Write-Host "Failed to kill and remove Spotify"
        }
    }

    
}

# Connect to each online computer and run the fuctions to remove the program
foreach ($hostname in $online) {
    try {
        $session = New-PSSession -ComputerName $hostname -ErrorAction Stop
        Remove-BingWallpaper -session $session -hostname $hostname
        Remove-Zoom -session $session -hostname $hostname
        Remove-Spotify -session $session -hostname $hostname
        Remove-PSSession $session
    } catch {
        Write-Host "Failed to connect to ${hostname}:`n$($_.Exception.Message)"
    }
}