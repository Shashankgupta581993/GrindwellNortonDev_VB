param(
    [Parameter(Mandatory=$true)][string]$BaselineFolder,
    [Parameter(Mandatory=$true)][string]$OptimizedFolder,
    [string[]]$Files = @('03_OrderOperationSnapshot.csv','05_WipChainDiagnostics.csv',
        '06_StageEligibilityDiagnostics.csv','07_OptimizerCandidateTrace.csv',
        '08_BatchTunnelSwkPlanTrace.csv','09_ScheduleAttemptTrace.csv',
        '10_ResourceValidation.csv','11_ReasonCodeSummary.csv')
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName Microsoft.VisualBasic

function Open-Csv([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Missing required capture: $Path" }
    $parser = New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($Path)
    $parser.SetDelimiters(',')
    $parser.HasFieldsEnclosedInQuotes = $true
    $parser.TrimWhiteSpace = $false
    return $parser
}

foreach ($file in $Files) {
    $old = $null
    $new = $null
    try {
        $old = Open-Csv (Join-Path $BaselineFolder $file)
        $new = Open-Csv (Join-Path $OptimizedFolder $file)
        $oldHeader = $old.ReadFields()
        $newHeader = $new.ReadFields()
        if ($null -eq $oldHeader -or $null -eq $newHeader) { throw "Missing CSV header: $file" }
        if ($oldHeader.Length -ne $newHeader.Length) { throw "Header width differs: $file" }
        $compareColumns = New-Object 'System.Collections.Generic.List[int]'
        for ($column = 0; $column -lt $oldHeader.Length; $column++) {
            if (-not [string]::Equals($oldHeader[$column], $newHeader[$column], [StringComparison]::Ordinal)) {
                throw "Header differs: $file, column $column"
            }
            if ($oldHeader[$column] -notin @('RunId','ExportedAt')) { $compareColumns.Add($column) }
        }
        $row = 0
        while (-not $old.EndOfData -and -not $new.EndOfData) {
            $oldFields = $old.ReadFields()
            $newFields = $new.ReadFields()
            $row++
            if ($oldFields.Length -ne $oldHeader.Length -or $newFields.Length -ne $newHeader.Length) {
                throw "Malformed CSV width: $file, data row $row"
            }
            foreach ($column in $compareColumns) {
                if (-not [string]::Equals($oldFields[$column], $newFields[$column], [StringComparison]::Ordinal)) {
                    throw "Value/order differs: $file, data row $row, column $($oldHeader[$column])"
                }
            }
        }
        if ($old.EndOfData -ne $new.EndOfData) { throw "Row count differs: $file" }
        Write-Output "PASS: $file ($row ordered rows)."
    } finally {
        if ($null -ne $old) { $old.Dispose() }
        if ($null -ne $new) { $new.Dispose() }
    }
}
Write-Output 'Capture parity passed. This compares supplied files; it does not execute Opcenter or prove uncaptured business side effects.'
