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
        Write-Host "ONLINE" -ForegroundColor Green
        $online += $computer
    } else {
        Write-Host $computer "OFFLINE" -ForegroundColor Red
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

    $removedBing = $false

    try {
        $users = Invoke-Command -Session $session -ScriptBlock {
            Get-ChildItem -Path "C:\Users" -Directory | Select-Object -ExpandProperty Name
        }

        foreach ($user in $users) {
            $localPath = "C:\Users\$user\AppData\Local\Microsoft\WindowsApps\Microsoft.BingWallpaper*"
            $roamingPath = "C:\Users\$user\AppData\Roaming\Microsoft\WindowsApps\Microsoft.BingWallpaper*" 

            $result = Invoke-Command -Session $session -ScriptBlock {
                param ($local, $roaming)
                $deleted = $false

                if (Test-Path -Path $local) {Remove-Item -Path $local -Recurse -Force -ErrorAction SilentlyContinue; $deleted = $true}
                if (Test-Path -Path $roaming) {Remove-Item -Path $roaming -Recurse -Force -ErrorAction SilentlyContinue; $deleted = $true}

                return $deleted
            } -ArgumentList $localPath, $roamingPath

            if ($result) {
                $removedBing = $true
            }
        }

        if ($removedBing) {
            Write-Host "Bing Wallpaper has been removed from $hostname" -ForegroundColor Green
        } else {
            Write-Host "Bing Wallpaper was not found on $hostname" -ForegroundColor Yellow
        }

        
    } catch {
        Write-host "Error while removing Bing Wallpaper from ${hostname}:`n$($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Trying to kill bing wallpaper and retry. Please wait..."
        Invoke-Command -Session $session -ScriptBlock {
            Get-Process BingWallpaper -ErrorAction SilentlyContinue | Stop-Process -Force
        }
        #Try to delete after killing Bing Wallpaper
        try {
            foreach ($user in $users) {
                $localPath = "C:\Users\$user\AppData\Local\Microsoft\WindowsApps\Microsoft.BingWallpaper*"
                $roamingPath = "C:\Users\$user\AppData\Roaming\Microsoft\WindowsApps\Microsoft.BingWallpaper*"

                $result = Invoke-Command -Session $session -ScriptBlock {
                    param ($local, $roaming)
                    $deleted = $false

                    if (Test-Path -Path $local) {Remove-Item -Path $local -Recurse -Force -ErrorAction SilentlyContinue; $deleted = $true}
                    if (Test-Path -Path $roaming) {Remove-Item -Path $roaming -Recurse -Force -ErrorAction SilentlyContinue; $deleted = $true}

                    return $deleted
                } -ArgumentList $localPath, $roamingPath

                if ($result) {
                    $removedBing = $true
                }
            }

            if ($removedBing) {
                Write-Host "Bing Wallpaper has been removed from $hostname after killing the process" -ForegroundColor Green
            } else {
                Write-Host "Bing wallpaper was not found on $hostname" -ForegroundColor Yellow
            }

        } catch {
            Write-host "Error while removing Bing Wallpaper from ${hostname}:`n$($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Bing Wallpaper could not be removed" -ForegroundColor Red
        }
    }
}

#Function to delete Zoom
function Remove-Zoom {
    param ($session, $hostname)

    $removedZoom = $false

    try {
        $users = Invoke-Command -Session $session -ScriptBlock {
            Get-ChildItem -Path "C:\Users" -Directory | Select-Object -ExpandProperty Name
        }

        foreach ($user in $users) {
            $localPath = "C:\Users\$user\AppData\Local\Zoom"
            $roamingPath = "C:\Users\$user\AppData\Roaming\Zoom" 

            $result = Invoke-Command -Session $session -ScriptBlock {
                param ($local, $roaming)

                $deleted = $false

                if (Test-Path -Path $local) {Remove-Item -Path $local -Recurse -Force -ErrorAction SilentlyContinue; $deleted = $true}
                if (Test-Path -Path $roaming) {Remove-Item -Path $roaming -Recurse -Force -ErrorAction SilentlyContinue; $deleted = $true}

                return $deleted
            } -ArgumentList $localPath, $roamingPath

            if ($result) {
                $removedZoom = $true
            }
        }

        if ($removedZoom) {
            Write-Host "Zoom has been removed from $hostname" -ForegroundColor Green
        } else {
            Write-Host "Zoom was not found on $hostname" -ForegroundColor Yellow
        }

    } catch {
        Write-Host "Error while removing Zoom from ${hostname}:`n$($_.Exception.Message)" -ForegroundColor Red
        Write-Host "Trying to kill Zoom and retry. Please wait..."
        Invoke-Command -Session $session -ScriptBlock {
            Get-Process Zoom -ErrorAction SilentlyContinue | Stop-Process -Force
        }
        #Try to delete after killing Zoom
        try {
            foreach ($user in $users) {
                $localPath = "C:\Users\$user\AppData\Local\Zoom"
                $roamingPath = "C:\Users\$user\AppData\Roaming\Zoom"

                $result = Invoke-Command -Session $session -ScriptBlock {
                    param ($local, $roaming)

                    $deleted = $false

                    if (Test-Path -Path $local) {Remove-Item -Path $local -Recurse -Force -ErrorAction SilentlyContinue; $deleted = $true}
                    if (Test-Path -Path $roaming) {Remove-Item -Path $roaming -Recurse -Force -ErrorAction SilentlyContinue; $deleted = $true}

                    return $deleted
                } -ArgumentList $localPath, $roamingPath

                if ($result) {
                    $removedZoom = $true
                }
            }

            if ($removedZoom) {
                Write-Host "Zoom has been removed after killing the process on $hostname" -ForegroundColor Green
            } else {
                Write-Host "Zoom could not be found on $hostname" -ForegroundColor Yellow
            }

        } catch {
            Write-Host "Error while removing Zoom from ${hostname}:`n$($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Failed to kill and remove Zoom" -ForegroundColor Red
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

    $removedSpotify = $false

    try {
        $users = Invoke-Command -Session $session -ScriptBlock {
            Get-ChildItem -Path "C:\Users" -Directory | Select-Object -ExpandProperty Name
        }

        foreach ($user in $users) {
            $localPath = "C:\Users\$user\AppData\Local\Microsoft\WindowsApps\Spotify*"
            $roamingPath = "C:\Users\$user\AppData\Roaming\Microsoft\WindowsApps\Spotify*" 
            $downloadPath = "C:\Users\$user\Downloads\Spotify*"
            $programPath = "C:\Program Files\WindowsApps\Spotify*"

            $result = Invoke-Command -Session $session -ScriptBlock {
                param ($local, $roaming, $download, $program)

                $deleted = $false

                if (Test-Path -Path $local) {Remove-Item -Path $local -Recurse -Force -ErrorAction SilentlyContinue; $deleted = $true}
                if (Test-Path -Path $roaming) {Remove-Item -Path $roaming -Recurse -Force -ErrorAction SilentlyContinue; $deleted = $true}
                if (Test-Path -Path $download) {Remove-Item -Path $download -Recurse -Force -ErrorAction SilentlyContinue; $deleted = $true}
                if (Test-Path -Path $program) {Remove-Item -Path $program -Recurse -Force -ErrorAction SilentlyContinue; $deleted = $true}

                return $deleted
            } -ArgumentList $localPath, $roamingPath, $downloadPath, $programPath

            if ($result) {
                $removedSpotify = $true
            }
        }

        if ($removedSpotify) {
            Write-Host "Spotify has been removed from $hostname" -ForegroundColor Green
        } else {
            Write-Host "Spotify could not be found on $hostname" -ForegroundColor Yellow
        }

    } catch {
        Write-Host "Error while removing Spotify from ${hostname}:`n$($_.Exception.Message)" -ForegroundColor Red
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

                $result = Invoke-Command -Session $session -ScriptBlock {
                    param ($local, $roaming, $download, $program)

                    $deleted = $false

                    if (Test-Path -Path $local) {Remove-Item -Path $local -Recurse -Force -ErrorAction SilentlyContinue; $deleted = $true}
                    if (Test-Path -Path $roaming) {Remove-Item -Path $roaming -Recurse -Force -ErrorAction SilentlyContinue; $deleted = $true}
                    if (Test-Path -Path $download) {Remove-Item -Path $download -Recurse -Force -ErrorAction SilentlyContinue; $deleted = $true}
                    if (Test-Path -Path $program) {Remove-Item -Path $program -Recurse -Force -ErrorAction SilentlyContinue; $deleted = $true}

                    return $deleted

                } -ArgumentList $localPath, $roamingPath, $downloadPath, $programPath

                if ($result) {
                    $removedSpotify = $true
                }
            }

            if ($removedSpotify) {
                Write-Host "Spotify has been removed after killing process on $hostname" -ForegroundColor Green
            } else {
                Write-Host "Spotify could not be found on $hostname" -ForegroundColor Yellow
            }

        } catch {
            Write-Host "Error while removing Spotify from ${hostname}:`n$($_.Exception.Message)" -ForegroundColor Red
            Write-Host "Failed to kill and remove Spotify" -ForegroundColor Red
        }
    }
}

function Invoke-SCCMSoftwareInventory {
    param($session, $hostname)

    try {
        Invoke-Command -Session $session -ScriptBlock {
            Invoke-WmiMethod -Namespace 'root\ccm' -Class 'SMS_Client' -Name TriggerSchedule -ArgumentList '{00000000-0000-0000-0000-000000000002}'
        }

        Write-Host "Triggered SCCM software inventory on $hostname"
    }
    catch {
        Write-Host "Failed to trigger SCCM software inventory on ${hostname}:`n$($_.Exception.Message)" -ForegroundColor Red
    }
}

# Connect to each online computer and run the fuctions to remove the program
foreach ($hostname in $online) {
    try {
        $session = New-PSSession -ComputerName $hostname -ErrorAction Stop
        Remove-BingWallpaper -session $session -hostname $hostname
        Remove-Zoom -session $session -hostname $hostname
        Remove-Spotify -session $session -hostname $hostname
        Invoke-SCCMSoftwareInventory -session $session -hostname $hostname
        Remove-PSSession $session
    } catch {
        Write-Host "Failed to connect to $hostname"
        Write-Host ""
        Write-Host $($_.Exception.Message) -ForegroundColor Red
    }
}