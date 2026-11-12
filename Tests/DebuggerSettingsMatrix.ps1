param([string]$MSBuild='C:/Program Files/Microsoft Visual Studio/2022/Community/MSBuild/Current/Bin/MSBuild.exe', [string]$CaseName='')
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$root=Split-Path -Parent $PSScriptRoot
$sourcePath=Join-Path $root 'Debugging/SchedulerDebugCollector.vb'
$source=Get-Content -Raw -LiteralPath $sourcePath
$channels=@([regex]::Matches($source,'Friend ReadOnly (\w+)Enabled As Boolean') | ForEach-Object { $_.Groups[1].Value } | Where-Object { $_ -ne 'Master' })
$cases=@(@{Name='MasterOff';Enabled=@();Master=$false},@{Name='AllOn';Enabled=$channels;Master=$true},@{Name='AllChannelsOff';Enabled=@();Master=$true},@{Name='AttemptsAndReasons';Enabled=@('ScheduleAttemptTrace','ReasonSummary');Master=$true})
$cases+=@($channels | ForEach-Object { @{Name=$_;Enabled=@($_);Master=$true} })
if ($CaseName) { $cases=@($cases | Where-Object Name -EQ $CaseName) }
foreach($case in $cases) {
    $caseRoot=Join-Path $PSScriptRoot ('obj/DebuggerMatrix/'+$case.Name)
    [IO.Directory]::CreateDirectory($caseRoot) | Out-Null
    $variant=$source
    foreach($name in @('Master')+$channels) {
        $value=if($name -eq 'Master') { $case.Master } else { $case.Enabled -contains $name }
        $literal=if($value -or (-not $case.Master -and $name -ne 'Master')) { 'True' } else { 'False' }
        $variant=[regex]::Replace($variant,"(Friend ReadOnly ${name}Enabled As Boolean = )(True|False)",('${1}'+$literal))
    }
    $variantPath=Join-Path $caseRoot 'SchedulerDebugCollector.vb'
    [IO.File]::WriteAllText($variantPath,$variant)
    [xml]$project=Get-Content -Raw -LiteralPath (Join-Path $root 'GrindwellNortonDev_VB.vbproj')
    foreach($item in $project.SelectNodes('//*[@Include]')) {
        if($item.LocalName -in @('Compile','EmbeddedResource','None')) {
            $absolute=[IO.Path]::GetFullPath((Join-Path $root $item.GetAttribute('Include')))
            if($absolute -eq [IO.Path]::GetFullPath($sourcePath)) { $absolute=$variantPath }
            $item.SetAttribute('Include',$absolute)
        }
    }
    $projectPath=Join-Path $caseRoot 'SettingsVariant.vbproj'
    $project.Save($projectPath)
    & $MSBuild $projectPath /t:Build /p:Configuration=Release /p:Platform=x64 /p:RegisterForComInterop=false "/p:ReferencePath=$root/bin/Debug" /p:BaseIntermediateOutputPath=obj/ /p:IntermediateOutputPath=obj/ /p:OutputPath=bin/ /v:quiet /nologo
    if($LASTEXITCODE -ne 0) { throw "Build failed: $($case.Name)" }
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'SchedulerDiagnosticsHarness.ps1') -AssemblyPath (Join-Path $caseRoot 'bin/GrindwellNortonDev_VB.dll') -EnabledChannels (($case.Enabled -join ',') + ',__none__')
    if($LASTEXITCODE -ne 0) { throw "Assertions failed: $($case.Name)" }
}
Write-Output "PASS: $($cases.Count) isolated configurations; production settings unchanged."
