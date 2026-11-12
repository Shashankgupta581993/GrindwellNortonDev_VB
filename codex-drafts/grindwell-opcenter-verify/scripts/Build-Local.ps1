#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$RepositoryRoot,
    [string]$SdkReferencePath,
    [string]$MSBuildPath
)

$ErrorActionPreference = 'Stop'

$root = (Resolve-Path -LiteralPath $RepositoryRoot).ProviderPath
$project = Join-Path $root 'GrindwellNortonDev_VB.vbproj'
if (-not (Test-Path -LiteralPath $project -PathType Leaf)) {
    throw "Expected the Grindwell Norton VB.NET project at '$project'."
}

if (-not $SdkReferencePath) {
    $SdkReferencePath = Join-Path $root 'bin/Debug'
}
if (-not (Test-Path -LiteralPath $SdkReferencePath -PathType Container)) {
    throw "SDK directory '$SdkReferencePath' is missing. Supply -SdkReferencePath from the matching Opcenter installation."
}
$references = (Resolve-Path -LiteralPath $SdkReferencePath).ProviderPath

if (-not $MSBuildPath) {
    $vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio/Installer/vswhere.exe'
    if (Test-Path -LiteralPath $vswhere -PathType Leaf) {
        $MSBuildPath = & $vswhere -latest -products '*' -requires Microsoft.Component.MSBuild -find 'MSBuild\Current\Bin\MSBuild.exe' |
            Select-Object -First 1
    }
}
if (-not $MSBuildPath -or -not (Test-Path -LiteralPath $MSBuildPath -PathType Leaf)) {
    throw 'Visual Studio MSBuild was not found. Supply -MSBuildPath to MSBuild.exe.'
}

$buildId = [guid]::NewGuid().ToString('N')
$output = "bin\CodexLocal\$buildId\"
$intermediate = "obj\CodexLocal\$buildId\"
$buildArgs = @(
    $project
    '/t:Build'
    '/p:Configuration=Debug'
    '/p:Platform=x64'
    '/p:RegisterForComInterop=false'
    '/p:SignAssembly=false'
    "/p:ReferencePath=$references"
    "/p:BaseIntermediateOutputPath=$intermediate"
    "/p:IntermediateOutputPath=$intermediate"
    "/p:OutputPath=$output"
    '/v:minimal'
)

& $MSBuildPath @buildArgs | Out-Host
if ($LASTEXITCODE -ne 0) {
    throw "MSBuild failed with exit code $LASTEXITCODE."
}

$dll = Join-Path (Join-Path $root $output) 'GrindwellNortonDev_VB.dll'
if (-not (Test-Path -LiteralPath $dll -PathType Leaf)) {
    throw "Expected the built assembly at '$dll'."
}
$dll
