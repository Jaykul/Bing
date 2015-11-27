[CmdletBinding()]
param(
    # The last digit of the build version number (by default comes from AppVeyor)
    [Nullable[int]]$RevisionNumber = ${Env:APPVEYOR_BUILD_NUMBER}
)

foreach($Path in $(Get-ChildItem $PSScriptRoot\*\src | Split-Path | Convert-Path)) {
    # The Module name is used to identify Manifest and scrape version (defaults to the folder name)
    [string]$ModuleName = Split-Path $Path -Leaf

    $OutputPath = Join-Path $Path output
    $null = mkdir $OutputPath -Force

    $ErrorActionPreference = "Stop"
    Set-StrictMode -Version Latest
    Write-Host "Package: $ModuleName to $OutputPath"

    $Version = &"${PSScriptRoot}\Get-Version.ps1" -Module (Join-Path $Path\src "${ModuleName}.psd1") -DevBuild:$RevisionNumber -RevisionNumber:$RevisionNumber
    $ReleasePath = Join-Path $Path $Version

    Write-Verbose "COPY   $ReleasePath\"
    $null = robocopy $ReleasePath "${OutputPath}\${ModuleName}" /MIR /NP /LOG+:"$OutputPath\build.log"

    $zipFile = Join-Path $OutputPath "${ModuleName}-${Version}.zip"
    Add-Type -assemblyname System.IO.Compression.FileSystem
    Remove-Item $zipFile -ErrorAction SilentlyContinue
    Write-Verbose "ZIP    $zipFile"
    [System.IO.Compression.ZipFile]::CreateFromDirectory((Join-Path $OutputPath $ModuleName), $zipFile)

    # You can add other artifacts here
    ls $OutputPath -File

}
