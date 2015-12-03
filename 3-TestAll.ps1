[CmdletBinding()]
param(
    [Switch]$SkipBuild,

    [Switch]$Quiet,

    [Switch]$ShowWip,

    $FailLimit=0,
    
    [ValidateNotNullOrEmpty()]    
    [String]$JobID = ${Env:APPVEYOR_JOB_ID},

    [Nullable[int]]$RevisionNumber = ${Env:APPVEYOR_BUILD_NUMBER},

    [ValidateNotNullOrEmpty()]
    [String]$CodeCovToken = ${ENV:CODECOV_TOKEN}
)

$ParameterValues = @{}
foreach($parameter in $MyInvocation.MyCommand.Parameters.GetEnumerator()) {
    try {
        $key = $parameter.Key
        $value = Get-Variable -Name $key -ValueOnly -ErrorAction SilentlyContinue -ErrorVariable NoVar
        if(!$NoVar) {
            $ParameterValues.$key = $value
        }
    } finally {}
}

# Path of the folder to build from (defaults to the folder the Build script is in)
[string[]]$Path = $(Get-ChildItem $PSScriptRoot\*\src | Split-Path)


foreach($module in $Path) {
    &"$PSScriptRoot\TestModule.ps1" -Path $module @ParameterValues
}
