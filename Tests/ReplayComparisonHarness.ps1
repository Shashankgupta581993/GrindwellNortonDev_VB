Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$folder = Join-Path ([IO.Path]::GetTempPath()) ('gn_replay_comparison_' + [Guid]::NewGuid().ToString('N'))
$baseline = Join-Path $folder 'baseline'
$optimized = Join-Path $folder 'optimized'
$comparator = Join-Path $PSScriptRoot 'Compare-OpcenterReplay.ps1'
try {
    [IO.Directory]::CreateDirectory($baseline) | Out-Null
    [IO.Directory]::CreateDirectory($optimized) | Out-Null
    $oldPath = Join-Path $baseline 'fixture.csv'
    $newPath = Join-Path $optimized 'fixture.csv'
    $oldText = "RunId,ExportedAt,RecordNo,Detail`r`nold,first,1,`"comma,quote`"`"`r`nline`"`r`nold,first,2,unchanged`r`n"
    $newText = $oldText.Replace('old,first,','new,second,')
    [IO.File]::WriteAllText($oldPath, $oldText)
    [IO.File]::WriteAllText($newPath, $newText)
    & $comparator -BaselineFolder $baseline -OptimizedFolder $optimized -Files fixture.csv | Out-Null
    foreach ($badText in @($newText.Replace('unchanged','changed'),
                          $newText.Replace('RecordNo','DifferentHeader'),
                          $newText.Replace("new,second,2,unchanged`r`n",''))) {
        [IO.File]::WriteAllText($newPath, $badText)
        $rejected = $false
        try { & $comparator -BaselineFolder $baseline -OptimizedFolder $optimized -Files fixture.csv | Out-Null } catch { $rejected = $true }
        if (-not $rejected) { throw 'Comparator accepted differing captures.' }
    }
    $rejected = $false
    try { & $comparator -BaselineFolder $baseline -OptimizedFolder $optimized -Files absent.csv | Out-Null } catch { $rejected = $true }
    if (-not $rejected) { throw 'Comparator accepted absent evidence.' }
    Write-Output 'PASS: replay comparator handles multiline/quoted CSV, volatile fields, and rejects value/header/count/missing-file differences.'
} finally {
    $resolved = [IO.Path]::GetFullPath($folder)
    if ($resolved.StartsWith([IO.Path]::GetFullPath([IO.Path]::GetTempPath()), [StringComparison]::OrdinalIgnoreCase)) {
        Remove-Item -LiteralPath $resolved -Recurse -Force
    }
}
