$nodeVersion = "20.18.0"
$winswVersion = "2.12.0"
$logrotateVersion = "2.7.0"

function refreshenv {
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User") 
}

# using this instead of Invoke-WebRequest avoids restrictions on downloads due to the "Internet Explorer Enhanced Security Configuration"
$webClient = New-Object System.Net.WebClient

$offlineFolder = (Get-Item '.\offline').FullName
Remove-Item -Path $offlineFolder -Recurse -Force
New-Item -Path $offlineFolder -ItemType Directory

Write-Host "Download Node: $nodeVersion"
$webClient.DownloadFile("https://nodejs.org/dist/v$nodeVersion/node-v$nodeVersion-x64.msi", "$offlineFolder/nodejs.msi")
# Download portable also, to ensure this script uses the same node version. (couldn't get the msi extracted easily into a folder...)
$webClient.DownloadFile("https://nodejs.org/dist/v$nodeVersion/node-v$nodeVersion-win-x64.zip", "$offlineFolder/nodejs.zip")
Expand-Archive -Path "$offlineFolder/nodejs.zip" -DestinationPath "$offlineFolder/node"

Write-Host "Add the portable node to the Path"
$npm = "$offlineFolder\node\node-v$nodeVersion-win-x64"
$origPath = $env:Path
$env:PATH = $npm + ";" + $origPath

Write-Host "Download WinSW: $winswVersion"
$webClient.DownloadFile("https://github.com/winsw/winsw/releases/download/v$winswVersion/WinSW.NET4.exe",  "$offlineFolder/WinSW.NET4.exe")

Write-Host "Download PM2"
npm install pm2 --prefix $offlineFolder/pm2

Write-Host "Download pm2-logrotate"
$webClient.DownloadFile("https://github.com/keymetrics/pm2-logrotate/archive/refs/tags/$logrotateVersion.tar.gz", "$offlineFolder/logrotate.tar.gz")
New-Item -Path "$offlineFolder\pm2-logrotate" -ItemType Directory
tar -xzvf "$offlineFolder/logrotate.tar.gz" --strip-components=1 -C "$offlineFolder\pm2-logrotate"
Remove-Item -Path "$offlineFolder/logrotate.tar.gz"
Set-Location "$offlineFolder\pm2-logrotate\"
npm install
Set-Location -Path $offlineFolder
tar -czvf "$offlineFolder\logrotate.tar.gz" "pm2-logrotate"
Remove-Item -Path "$offlineFolder\pm2-logrotate\" -Recurse -Force

Remove-Item -Path "$offlineFolder\nodejs.zip"
Remove-Item -Path "$offlineFolder\node" -Recurse -Force

Write-Host "Restore environment"
Set-Location -Path $PSScriptRoot
$env:Path = $origPath
Write-Host "Done"