[CmdletBinding()]
param(
    # Path of the folder to build from (defaults to the folder the Build script is in)
    [Alias("PSPath")]
    [string]$Path = $PSScriptRoot,
    
    # The Module name is used to identify Manifest and scrape version (defaults to the folder name)
    [string]$ModuleName = $(Split-Path $Path -Leaf),
    
    # The target framework for .net (for packages), with fallback versions
    # The default supports PS3+:  "net40","net35","net20","net45","net451","net452","net46","net461","net462"
    # To only support PS4, use:  "net45","net40","net35","net20"
    # To support PS2, you use:   "net35","net20"
    [string[]]$TargetFramework = @("net40","net35","net20","net45","net451","net452","net46","net461","net462"),
        
    # The last digit of the build version number (by default comes from AppVeyor)
    [Nullable[int]]$RevisionNumber = ${Env:APPVEYOR_BUILD_NUMBER},
    
    # MSBuild Target (defaults to "Build")
    $Target="Build",
    
    # MSBuild Configuration (defaults to "Release")
    $Configuration="Release"
)
$Path = Convert-Path $Path
$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest
Write-Host "BUILDING: $ModuleName from $Path"

$OutputPath = Join-Path $Path output
$null = mkdir $OutputPath -Force

# If the RevisionNumber is specified as ZERO, this is a release build
$Version = &"${PSScriptRoot}\Get-Version.ps1" -Module (Join-Path $Path\src "${ModuleName}.psd1") -DevBuild:$RevisionNumber -RevisionNumber:$RevisionNumber
$ReleasePath = Join-Path $Path $Version

Write-Verbose "OUTPUT Release Path: $ReleasePath"
if(Test-Path $ReleasePath) {
    Write-Verbose "       Clean up old build"
    Write-Verbose "DELETE $ReleasePath\"
    Remove-Item $ReleasePath -Recurse -Force -ErrorAction SilentlyContinue
}

## Find dependency Package Files
Write-Verbose "       Copying Packages"
$Packages = Join-Path $Path packages.config
if(Test-Path $Packages) {
    foreach($Package in ([xml](Get-Content $Packages)).packages.package) {
        $folder = Join-Path $Path "packages\$($Package.id)*"
        # Check for each TargetFramework, in order of preference, fall back to using the lib folder
        $targets = ($TargetFramework -replace '^','lib\') + 'lib' | ForEach-Object { Join-Path $folder $_ }
        $PackageSource = Get-Item $targets -ErrorAction SilentlyContinue | Select -First 1 -Expand FullName
        if(!$PackageSource) {
            throw "Could not find a lib folder for $($Package.id) from package. You may need to run Setup.ps1"
        }

        Write-Verbose "COPY   $PackageSource\"
        $null = robocopy $PackageSource $ReleasePath\lib /MIR /NP /LOG:"$OutputPath\build.log" /R:2 /W:15
        if($LASTEXITCODE -gt 1) {
            throw "Failed to copy Package $($Package.id) (${LASTEXITCODE}), see build.log for details"
        }
    }
}

## If there's a solution file, build it.
foreach($solution in Get-Item $Path\*.sln) {
    msbuild $solution.FullName /t:$Target /p:Configuration=$Configuration /p:OutputPath=$ReleasePath\lib
}

## Copy PowerShell source Files
Write-Verbose "       Copying Module Source"
Write-Verbose "COPY   '$Path\src\' '$ReleasePath'"
$null = robocopy $Path\src\  $ReleasePath /E /NP /LOG+:"$OutputPath\build.log" /R:2 /W:15
if($LASTEXITCODE -gt 3) {
    throw "Failed to copy Module (${LASTEXITCODE}), see build.log for details"
}
## Touch the PSD1 Version:
Write-Verbose "       Update Module Version"
$ReleaseManifest = Join-Path $ReleasePath "${ModuleName}.psd1"
Set-Content $ReleaseManifest ((Get-Content $ReleaseManifest) -Replace "^(\s*)ModuleVersion\s*=\s*'?[\d\.]+'?\s*`$", "`$1ModuleVersion = '$Version'")
Get-Module $ReleaseManifest -ListAvailable