function refreshenv {
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User") 
}

# check Administrator privileges
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$isAdmin=$currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if($isAdmin -eq $true){
    Write-host "Script is running with Administrator privileges!"
}
else {
    Write-host "Script is not running with Administrator privileges,cannot be process."
    return
}

$offlineFolder = $offlineFolder = (Get-Item '.\offline').FullName
$isOffline = Test-Path -Path $offlineFolder
if ($isOffline) {
    "Running in offline mode since the folder '$offlineFolder' exists."
}

# Allow Execution of Foreign Scripts
Set-ExecutionPolicy Bypass -Scope Process -Force;

if (-not $isOffline) {
    # Use TLS 1.2
    [System.Net.ServicePointManager]::SecurityProtocol = 3072;

    $wingetversion=$null
    try { $wingetversion=winget -v } catch {}
    if($null -eq $wingetversion){
        # Install VCLibs
        Add-AppxPackage 'https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx'

        # Install Microsoft.UI.Xaml (latest) from NuGet
        Invoke-WebRequest -Uri https://www.nuget.org/api/v2/package/Microsoft.UI.Xaml/ -OutFile .\microsoft.ui.xaml.zip
        Expand-Archive .\microsoft.ui.xaml.zip
        Add-AppxPackage .\microsoft.ui.xaml\tools\AppX\x64\Release\Microsoft.UI.Xaml.2.8.appx

        # Install the latest release of Microsoft.DesktopInstaller from GitHub
        Invoke-WebRequest -Uri https://github.com/microsoft/winget-cli/releases/latest/download/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle -OutFile .\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle
        Add-AppxPackage .\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle
    }
    refreshenv
}

if ($isOffline) {
    $nodeversion=$null
    try { $nodeversion=node -v } catch {}
    Write-host "Install Nodejs from offline package"
    if($null -ne $nodeversion){
        Write-host "Uninstall old Nodejs first"
        $app = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -match "Node.js" }
        $app.Uninstall()
    }

    Start-Process "msiexec.exe" -ArgumentList "/I `"$offlineFolder\nodejs.msi`" /qn" -Wait -NoNewWindow 
} else {
    winget install OpenJS.NodeJS
}
refreshenv

try { $nodeversion=node -v } catch {}
if($null -eq $nodeversion){
    Write-host "Node.js must be installed. "
    return
}
refreshenv


Set-Location -Path $PSScriptRoot
$winVersion = [System.Environment]::OSVersion.Version.Major
Write-host "Windws Version: $winVersion"

# install pm2
$pm2version=$null
$pm2Path="$env:ProgramData\pm2-etc"
try { $pm2version=pm2 -v } catch {}

if($null -eq $pm2version){

    if(-not (Test-Path $pm2Path) ){
        mkdir -Path "$pm2Path"
        # Grant Full Control permissions to the folder. 
        $newAcl = Get-Acl -Path "$pm2Path"
        $aclRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Users", "FullControl", "ContainerInherit, ObjectInherit", "None", "Allow")
        $newAcl.AddAccessRule($aclRule)
        Set-Acl -Path "$pm2Path" -AclObject $newAcl        
    }
    Set-Location -Path $pm2Path

    if ($isOffline) {
        Write-host "offline install pm2 ..."
        Copy-Item -Recurse -Path "$offlineFolder\pm2\*" -Destination $pm2Path
    } else {
        Write-host "online install pm2 ..."

        Write-host "npm config get registry"
        & npm config get registry

        Write-host "npm install pm2"
        & npm install pm2
    }   

    refreshenv    

    $sysPath = [Environment]::GetEnvironmentVariable('Path','Machine')
    $newPath="$pm2Path\node_modules\.bin"
    if ($Paths -notcontains $newPath) {
        $sysPath += ";$newPath"
        [Environment]::SetEnvironmentVariable('Path', $sysPath, 'Machine')
    }
    refreshenv

    & pm2
    Copy-Item "$($HOME)\.pm2" -Destination "$pm2Path\.pm2" -Force
    [Environment]::SetEnvironmentVariable('PM2_HOME', "$pm2Path\.pm2", 'Machine')
    refreshenv

    if(-not (Test-Path "$pm2Path\npm") ){
        mkdir -Path "$pm2Path\npm"
        # todo auth
    }  
    if(-not (Test-Path "$pm2Path\npm-cache") ){
        mkdir -Path "$pm2Path\npm-cache"
        # todo auth
    }      
    refreshenv
    & npm config --global set prefix ("$pm2Path\npm" -replace '\\','/')
    & npm config --global set cache  ("$pm2Path\npm-cache" -replace '\\','/')

    # https://stackoverflow.com/questions/61970999/pm2-logrotate-install-on-offline-linux-machine
    if($isOffline) {
        Write-host "install pm2 logrotate offline"

        mkdir -Path "$pm2Path\.pm2\modules"
        Set-Location -Path "$pm2Path\.pm2\modules"
        tar -xzvf "$offlineFolder\logrotate.tar.gz"
        & pm2 module:generate pm2-logrotate
        Set-Location -Path "$pm2Path\.pm2\modules\pm2-logrotate"
        
        Write-host "pm2 install ."
            & pm2 install .
        
        Write-host "pm2 save"
            & pm2 save
    } else {
        Write-host "pm2 install @jessety/pm2-logrotate"
            & pm2 install @jessety/pm2-logrotate
        Write-host "pm2 save"
            & pm2 save
    }
}
else {
    Write-host "PM2 $pm2version is already installed. You must uninstall PM2 to proceed."
}

Set-Location -Path $PSScriptRoot

# create windows service
if(-not (Test-Path "$pm2Path\service") ){
    mkdir -Path "$pm2Path\service"
}

# pm2 service code
$serviceCode=@'
pm2 kill
pm2 resurrect
while ($true) {
    Start-Sleep -Seconds 60
    # do nothing
}
'@
Add-Content -Path "$pm2Path\service\pm2service.ps1" -Value $serviceCode


# download WinSW (Windows Service Wrapper)
if ($isOffline) {
    Write-Host "install WinSW offline ..."
    Copy-Item -Path "$offlineFolder\WinSW.NET4.exe" -Destination "$pm2Path\service\pm2service.exe"
} else {
    try
    {
        Write-Host "downloading WinSW ..."
        Invoke-WebRequest "https://github.com/winsw/winsw/releases/download/v2.12.0/WinSW.NET4.exe" -OutFile "$pm2Path\service\pm2service.exe"
    } catch {
        $StatusCode = $_.Exception.Response.StatusCode.value__
        Write-Host "download StatusCode code: $StatusCode"
    }
}


if(-not(Test-Path "$pm2Path\service\pm2service.exe")){
    Write-Host "installation failed"
    return
}


# config WinSW for PM2 (Windows Service Wrapper)
$serviceConfig=@"
<service>
    <id>PM2</id>
    <name>PM2</name>
    <description>PM2 Admin Service</description>
    <logmode>roll</logmode>
    <depend></depend>
    <executable>pwsh.exe</executable>
    <arguments>-File "%BASE%\pm2service.ps1"</arguments>
</service>
"@
Set-Content -Path "$pm2Path\service\pm2service.xml" -Value $serviceConfig

Set-Location "$pm2Path\service"
& ./pm2service.exe Install
& ./pm2service.exe Start

Set-Location -Path $PSScriptRoot

refreshenv
