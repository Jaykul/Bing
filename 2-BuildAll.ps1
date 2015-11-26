[CmdletBinding()]
param(
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

$ParameterValues = @{}
foreach($parameter in $MyInvocation.MyCommand.Parameters.GetEnumerator()) {
    try {
        $key = $parameter.Key
        if($value = Get-Variable -Name $key -ValueOnly -ErrorAction Ignore) {
            $ParameterValues.$key = $value
        }
    } finally {}
}


$Env:PSModulePath += ";$PSScriptRoot"

# Path of the folder to build from (defaults to the folder the Build script is in)
[string[]]$Path = $(Get-ChildItem $PSScriptRoot\*\src | Split-Path)

foreach($module in $Path) {
    &"$PSScriptRoot\BuildModule.ps1" -Path $module @ParameterValues
}
