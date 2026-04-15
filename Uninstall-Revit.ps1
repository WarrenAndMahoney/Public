[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    # in theory this should support all Autodesk software, but better to be sure
    [ValidateSet("Revit", "3ds Max")]
    [String]$Software,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [String]$Year,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullorEmpty()]
    [string]$LogFolder = "C:\Temp",

    [Parameter(Mandatory = $false)]
    [switch]$EnableLogging = $true,

    [Parameter(Mandatory = $false)]
    [switch]$Force,

    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

if( [Environment]::UserName -ne "SYSTEM" )
{
    Write-Warning "Please run this as SYSTEM!"
    Write-Host "You can run ``psexec64.exe -si powershell`` to launch a terminal with SYSTEM privileges"
    return
}

if( $WhatIf )
{
    $EnableLogging = $false
}

if( ! (Test-Path -Path $LogFolder) )
{
    $LogFolder = $env:TEMP
}


# https://janikvonrotz.ch/2017/10/26/powershell-logging-in-cmtrace-format/
function Write-Log
{
    [CmdletBinding()]
    param(
        [parameter(Mandatory = $false)]
        [String]$Path = $LogPath,

        [parameter(Mandatory = $true, Position = 0)]
        [String]$Message,

        [parameter(Mandatory = $true)]
        [String]$Component,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Info", "Warning", "Error")]
        [String]$Type = "Info"
    )

    $timeStamp = Get-Date
    $timeString = Get-Date $timeStamp -Format "HH:mm:ss.ffffff"
    # CMTrace will convert it to a non-american format
    $dateString = Get-Date $timeStamp -Format "M-d-yyyy"
    $dateStringISO = Get-Date $timeStamp -Format "yyyy-MM-dd"

    # in case we enable logging for WhatIf
    if( $WhatIf )
    {
        $Component = "$Component-WhatIf"
    }

    Write-Host "$dateStringISO $timeString $Message"

    if( ! $EnableLogging )
    {
        return
    }

    switch ($Type)
    {
        "Info"
        {
            [int]$Type = 1
        }
        "Warning"
        {
            [int]$Type = 2
        }
        "Error"
        {
            [int]$Type = 3
        }
    }

    # Create a log entry
    $Content = "<![LOG[$Message]LOG]!>" + `
        "<time=`"$timeString`" " + `
        "date=`"$dateString`" " + `
        "component=`"$Component`" " + `
        "context=`"$([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)`" " + `
        "type=`"$Type`" " + `
        "thread=`"$([Threading.Thread]::CurrentThread.ManagedThreadId)`" " + `
        "file=`"`">"

    # Write the line to the log file
    #Add-Content -Path $Path -Value $Content

    # https://2pintsoftware.com/news/details/why-is-add-content-bad-in-powershell-51
    # Add-Content has to wait for it to release the handles and shit, this is waaaaay quicker
    $stream = [System.IO.StreamWriter]::new( $Path, $true, ([System.Text.Utf8Encoding]::new()) )
    $stream.WriteLine($Content)
    $stream.Close()
}

function Run-ApplicationDeploymentEvaluationCycle
{
    $invokeScheduleParam = @{
        Namespace  = 'root/ccm'
        Class      = 'SMS_CLIENT'
        MethodName = 'TriggerSchedule'
    }

    Invoke-CimMethod @invokeScheduleParam -Arguments @{ sScheduleID = "{00000000-0000-0000-0000-000000000121}" } | Out-Null
}

function Get-Uninstall
{
    param(
        [string]$uninstallString
    )

    [xml]$xml = $uninstallString
    $uninstall = $xml.SoftwareIdentity.Meta.UninstallString
    $uninstall -match 'C:.*?\.[a-z]{3}(\s|$)' | Out-Null
    $uninstallPath = $Matches[0]
    $uninstallParameters = $uninstall.Replace($uninstallPath, "")
    # quote all the paths in the uninstall string, since they're not quoted by default
    $uninstallParameters = $uninstallParameters -replace '(C:.*?\.[a-z]{3})( |$)', '"$1"$2'
    $uninstallParameters = "$uninstallParameters -q"

    $uninstall = @{
        Name       = $xml.SoftwareIdentity.Name
        Version    = $xml.SoftwareIdentity.Version
        Path       = $uninstallPath
        Parameters = $uninstallParameters
    }
    return $uninstall
}

function Get-InstalledSoftwareYears
{
    $installedVersions = Get-Package | Where-Object -Property name -Match "^(?:$Software 20\d\d$|Autodesk $Software 20\d\d)$" | Select-Object @{ name = "Name" ; expression = { $_.Name.Replace('Autodesk ', '') } }, Version | Sort-Object -Property @{ Expression = { [version]$_.Version } } | Group-Object -Property Name | ForEach-Object { $_.Group[-1] }
    #$installedVersions = Get-Package | Where-Object -Property name -match "^(?:$Software 20\d\d$|Autodesk $Software 20(?:2[5-9]|[3-9]\d))$" | Select-Object @{ name = "Name" ; expression = { $_.Name.Replace('Autodesk ','') } }, Version | Sort-Object -Property @{ Expression = { [version]$_.Version } } | Group-Object -Property Name | ForEach-Object { $_.Group[-1] }
    $installedSoftwareYears = $installedVersions.Name | ForEach-Object { $_.Split(" ")[-1] } | Sort-Object
    return $installedSoftwareYears
}

Write-Host "Installed $Software years: $((Get-InstalledSoftwareYears) -join ', ')" #-Component "Init"

while( 1 )
{
    if( ! $Year )
    {
        $yearToUninstall = Read-Host "$Software year to uninstall"
    }
    else
    {
        $yearToUninstall = $Year
    }
    # make sure it's a year from 2010-2039, by then i'm sure it'll be formit or whatever the fuck
    if( $yearToUninstall -match '^20[1-3][0-9]$' )
    {
        if( [int]$yearToUninstall -le 2020 -and ! $Force )
        {
            Write-Host "$Software 2020 and older use a different install mechanism so this may not work for uninstalling, please use -Force to try anyway" -ForegroundColor Red
            return
        }
        else
        {
            break
        }
    }
}

$currentlyInstalledRevit = Get-CimInstance -Namespace "ROOT\ccm\ClientSDK" -Class CCM_Application | Where-Object { $_.Name -like "$Software $Year*" -and $_.Name -notlike "* Update" }
# an array in case there's multiple that match
$currentlyInstalledRevitStates = @($currentlyInstalledRevit.InstallState)

$LogPath = "$LogFolder\Uninstall $Software $yearToUninstall.log"

Write-Log "$('=' * 80)" -Component "Visibility"
Write-Log "[ 1 / 3 ] Uninstalling $Software" -Component "$Software"

$softwarePackage = Get-Package | Where-Object { $_.name -eq "Autodesk $Software $yearToUninstall" -or $_.name -eq "$Software $yearToUninstall" } | Where-Object -Property ProviderName -EQ "Programs"
# uninstall the main program
if( $softwarePackage )
{
    $softwareUninstallProgram = Get-Uninstall $softwarePackage.SwidTagText
    Write-Log "  1 - [ 1 / 2 ] Uninstalling $($softwareUninstallProgram.Name) [$($softwareUninstallProgram.Version)]..." -Component "$Software"
    if( $WhatIf )
    {
        Write-Log "Start-Process `"$($softwareUninstallProgram.Path)`" -ArgumentList `"$($softwareUninstallProgram.Parameters)`" -Wait" -Component "$Software"
    }
    else
    {
        Start-Process $softwareUninstallProgram.Path -ArgumentList $softwareUninstallProgram.Parameters -Wait
    }
}
else
{
    Write-Log "  1 - [ 1 / 2 ] Already uninstalled: Autodesk $Software $yearToUninstall [Program]" -Component "$Software"
}

# uninstall the main MSI
$softwareUninstallMSI = Get-Package | Where-Object { $_.name -eq "Autodesk $Software $yearToUninstall" -or $_.name -eq "$Software $yearToUninstall" } | Where-Object -Property ProviderName -EQ "msi"
if( $softwareUninstallMSI )
{
    Write-Log "  1 - [ 2 / 2 ] Uninstalling $($softwareUninstallMSI.Name) [$($softwareUninstallMSI.Version)]..." -Component "$Software"
    $logName = "$($softwareUninstallMSI.Name) $($softwareUninstallMSI.Version)".Replace(" ", "_")
    if( $WhatIf )
    {
        Write-Log "Start-Process msiexec -ArgumentList `"/x ```"$($softwareUninstallMSI.FastPackageReference)```" /qn /l*v ```"$LogFolder\${logName}_uninstall.log```"`" -Wait" -Component "$Software"
    }
    else
    {
        Start-Process msiexec -ArgumentList "/x `"$($softwareUninstallMSI.FastPackageReference)`" /qn /l*v `"$LogFolder\${logName}_uninstall.log`"" -Wait
    }
}
else
{
    Write-Log "  1 - [ 2 / 2 ] Already uninstalled: $Software $yearToUninstall [MSI]" -Component "$Software"
}

Write-Log "[ 2 / 3 ] Uninstalling components" -Component "Components"

# the last two are plugins we deploy, so might as well keep them
$packages = Get-Package | Where-Object { $_.name -like "*$Software*$yearToUninstall*" -and $_.ProviderName -eq "Programs" -and $_.Name -notlike "*Content Catalog*" -and $_.Name -notlike "*Interoperability Tools*" }
$packageCounter = 1
$packageCount = $packages.Count
$packages | Select-Object -ExpandProperty SwidTagText | ForEach-Object {
    $uninstall = Get-Uninstall $_
    Write-Log ("  2 - [ {0,$($packageCount.Length)} / $packageCount ] Uninstalling $($uninstall.Name) [$($uninstall.Version)]..." -f $packageCounter) -Component "Components"
    if( $WhatIf )
    {
        Write-Log "Start-Process `"$($uninstall.Path)`" -ArgumentList `"$($uninstall.Parameters)`" -Wait" -Component "Components"
    }
    else
    {
        Start-Process $uninstall.Path -ArgumentList $uninstall.Parameters -Wait
    }
    $packageCounter += 1
}

# get all of the packages again because a lot of components have a 'program' and an 'msi', and you need to uninstall the program first
# which will uninstall the MSI as well, and we don't want to try uninstall already uninstalled MSIs
# this is just in case there are leftover components, or ones which are MSI only

Write-Log "[ 3 / 3 ] Uninstalling leftover components" -Component "Components"

$packages = Get-Package | Where-Object { $_.name -like "*$Software*$yearToUninstall*" -and $_.ProviderName -eq "msi" -and $_.Name -notlike "*Content Catalog*" -and $_.Name -notlike "*Interoperability Tools*" }
$packageCounter = 1
$packageCount = $packages.Count
$packages | ForEach-Object {
    Write-Log ("  3 - [ {0,$($packageCount.Length)} / $packageCount ] Uninstalling $($_.Name) [$($_.Version)]..." -f $packageCounter) -Component "Components"
    $logName = "$($_.Name) $($_.Version)".Replace(" ", "_")
    if( $WhatIf )
    {
        Write-Log "Start-Process msiexec -ArgumentList `"/x ```"$($_.FastPackageReference)```" /qn /l*v ```"$LogFolder\${logName}_uninstall.log```"`" -Wait" -Component "Components"
    }
    else
    {
        Start-Process msiexec -ArgumentList "/x `"$($_.FastPackageReference)`" /qn /l*v `"$LogFolder\${logName}_uninstall.log`"" -Wait
    }
    $packageCounter += 1
}

Run-ApplicationDeploymentEvaluationCycle

Write-Log "Running Application Deployment Evaluation..." -Component "Application Evaluation"

$currentWait = 0
$maxWait = 60
while( $currentWait -lt $maxWait )
{
    $revitInstallState = Get-CimInstance -Namespace "ROOT\ccm\ClientSDK" -Class CCM_Application | Where-Object { $_.Name -like "$Software $Year*" -and $_.Name -notlike "* Update" }
    if( "NotInstalled" -in @($revitInstallState.InstallState) )
    {
        break
    }
    $currentWait += 15
    Start-Sleep -Seconds 15
}


Start-Sleep -Seconds 60

Write-Log "Uninstallation complete, please restart before reinstalling $Software from Software Center!" -Component "$Software" -Type "Warning"
Write-Log "$('=' * 80)" -Component "Visibility"
