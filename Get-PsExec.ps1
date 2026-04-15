function Get-PsExec
{
    # it'd be cool to add a command param that gets run in psexec, but i don't think it's (easily) possible
    [CmdletBinding()]
    param(
        [switch]$StartPsExec
    )

    $pstoolsUri = "https://download.sysinternals.com/files/PSTools.zip"
    $outputPath = "C:\Temp\PSTools"
    if( ! (Test-Path -Path "$outputPath\psexec64.exe") )
    {
        $currentProgressPreference = $ProgressPreference
        $ProgressPreference = "SilentlyContinue"

        Write-Host "Downloading from [$pstoolsUri] to [$outputPath.zip]"
        Invoke-WebRequest -Uri $pstoolsUri -OutFile "$outputPath.zip"

        Write-Host "Extracting [$outputPath.zip] to [$outputPath]"
        Expand-Archive -Path "$outputPath.zip" -DestinationPath $outputPath

        Write-Host "Deleting [$outputPath.zip]"
        Remove-Item -Path "$outputPath.zip" -Force -Confirm:$false

        $ProgressPreference = $currentProgressPreference
    }

    if( $StartPsExec )
    {
        Write-Host "Launching PsExec"
        # -w sets the working dir, so set it so we immediately go to where we currently are
        Start-Process -FilePath "$outputPath\PsExec64.exe" -ArgumentList "-accepteula -w $((Get-Location).ProviderPath) -si powershell" -NoNewWindow
    }
}
