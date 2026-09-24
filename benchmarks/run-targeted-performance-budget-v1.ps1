[CmdletBinding()]
param(
    [string]$BenchmarkDll = (Join-Path $PSScriptRoot 'Bing.Offices.Benchmarks/bin/Release/net8.0/Bing.Offices.Benchmarks.dll'),
    [string]$ArtifactPath = (Join-Path $PSScriptRoot '../artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/targeted-performance-budget-v1-5rep.json'),
    [string]$BeforeProviderArtifact = (Join-Path $PSScriptRoot '../artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/provider-comparison-100k-before-final2.json'),
    [string]$BeforeRealIoArtifact = (Join-Path $PSScriptRoot '../artifacts/benchmarks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/real-io-before-final.json'),
    [int]$RowCount = 100000,
    [int]$Repetitions = 5
)

$ErrorActionPreference = 'Stop'

if ($RowCount -ne 100000) {
    throw 'Performance Budget v1 targeted confirmation is fixed to 100000 rows.'
}
if ($Repetitions -ne 5) {
    throw 'Performance Budget v1 targeted confirmation is fixed to exactly 5 repetitions.'
}

$BenchmarkDll = [IO.Path]::GetFullPath($BenchmarkDll)
$ArtifactPath = [IO.Path]::GetFullPath($ArtifactPath)
$BeforeProviderArtifact = [IO.Path]::GetFullPath($BeforeProviderArtifact)
$BeforeRealIoArtifact = [IO.Path]::GetFullPath($BeforeRealIoArtifact)
$artifactDirectory = [IO.Path]::GetDirectoryName($ArtifactPath)
$assemblyDirectory = [IO.Path]::GetDirectoryName($BenchmarkDll)

if (-not (Test-Path -LiteralPath $BenchmarkDll -PathType Leaf)) {
    throw "Benchmark DLL not found: $BenchmarkDll"
}
foreach ($path in @($BeforeProviderArtifact, $BeforeRealIoArtifact)) {
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        throw "Baseline artifact not found: $path"
    }
}
[IO.Directory]::CreateDirectory($artifactDirectory) | Out-Null

function Read-JsonLines {
    param([Parameter(Mandatory)][string]$Path)

    $encoding = [Text.UTF8Encoding]::new($false, $true)
    $text = [IO.File]::ReadAllText($Path, $encoding)
    foreach ($line in ($text -split "`r?`n")) {
        $trimmed = $line.Trim().TrimStart([char]0xFEFF)
        if ($trimmed.Length -eq 0) {
            continue
        }
        try {
            $trimmed | ConvertFrom-Json -Depth 32
        }
        catch {
            continue
        }
    }
}

function Get-Median {
    param([Parameter(Mandatory)][object[]]$Values)

    $ordered = @($Values | ForEach-Object { [double]$_ } | Sort-Object)
    if ($ordered.Count -eq 0) {
        return $null
    }
    $middle = [int]($ordered.Count / 2)
    if (($ordered.Count % 2) -eq 1) {
        return $ordered[$middle]
    }
    return ($ordered[$middle - 1] + $ordered[$middle]) / 2
}

function Get-HashMap {
    param([Parameter(Mandatory)][string]$Directory)

    $names = @(
        'Bing.Offices.Benchmarks.dll',
        'Bing.Offices.Abstractions.dll',
        'Bing.Offices.Core.dll',
        'Bing.Offices.Npoi.dll',
        'Bing.Offices.MiniExcel.dll'
    )
    $result = [ordered]@{}
    foreach ($name in $names) {
        $path = Join-Path $Directory $name
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            throw "Candidate assembly not found: $path"
        }
        $result[$name] = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    }
    return $result
}

function Assert-CandidateStable {
    param([Parameter(Mandatory)][hashtable]$Expected)

    $actual = Get-HashMap -Directory $assemblyDirectory
    foreach ($name in $Expected.Keys) {
        if ($actual[$name] -ne $Expected[$name]) {
            throw "Candidate assembly changed during targeted confirmation: $name"
        }
    }
}

function Invoke-BenchmarkChild {
    param([Parameter(Mandatory)][string[]]$Arguments)

    $startInfo = [Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName = 'dotnet'
    $startInfo.UseShellExecute = $false
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true
    $startInfo.CreateNoWindow = $true
    foreach ($argument in $Arguments) {
        [void]$startInfo.ArgumentList.Add($argument)
    }

    $process = [Diagnostics.Process]::new()
    $process.StartInfo = $startInfo
    try {
        if (-not $process.Start()) {
            throw 'dotnet child process did not start.'
        }
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        return [pscustomobject]@{
            exitCode = $process.ExitCode
            stdout = $stdout
            stderr = $stderr
        }
    }
    finally {
        $process.Dispose()
    }
}

function Get-LastJsonObject {
    param([Parameter(Mandatory)][string]$Text)

    $objects = @()
    foreach ($line in ($Text -split "`r?`n")) {
        $trimmed = $line.Trim().TrimStart([char]0xFEFF)
        if ($trimmed.Length -eq 0) {
            continue
        }
        try {
            $objects += $trimmed | ConvertFrom-Json -Depth 32
        }
        catch {
            continue
        }
    }
    if ($objects.Count -eq 0) {
        return $null
    }
    return $objects[$objects.Count - 1]
}

function Get-SampleMetrics {
    param([Parameter(Mandatory)][object[]]$Samples)

    $elapsed = @($Samples | ForEach-Object { $_.elapsedMilliseconds })
    $allocated = @($Samples | ForEach-Object { $_.allocatedBytes })
    $peak = @($Samples | ForEach-Object { $_.peakWorkingSetBytes })
    $outputs = @($Samples | ForEach-Object { $_.outputBytes })
    $asyncWrites = @($Samples | ForEach-Object { $_.asyncWriteCount })
    $asyncBytes = @($Samples | ForEach-Object { $_.asyncBytesWritten })
    return [ordered]@{
        sampleCount = $Samples.Count
        medianElapsedMilliseconds = Get-Median $elapsed
        medianAllocatedBytes = Get-Median $allocated
        medianPeakWorkingSetBytes = Get-Median $peak
        outputBytesMin = ($outputs | Measure-Object -Minimum).Minimum
        outputBytesMax = ($outputs | Measure-Object -Maximum).Maximum
        asyncWriteCountMin = ($asyncWrites | Measure-Object -Minimum).Minimum
        asyncBytesWrittenMin = ($asyncBytes | Measure-Object -Minimum).Minimum
    }
}

function Get-PercentDelta {
    param([object]$Before, [object]$After)

    if ($null -eq $Before -or $null -eq $After -or [double]$Before -eq 0) {
        return $null
    }
    return (([double]$After - [double]$Before) / [double]$Before) * 100
}

function Get-BeforeProvider {
    param([Parameter(Mandatory)][string]$Provider, [Parameter(Mandatory)][string]$Mode)

    $objects = @(Read-JsonLines -Path $BeforeProviderArtifact)
    $header = $objects | Where-Object { $_.kind -eq 'provider-comparison-probe' } | Select-Object -First 1
    $samples = @($objects | Where-Object {
            $_.kind -eq 'provider-comparison-sample' -and
            $_.provider -eq $Provider -and $_.mode -eq $Mode -and $_.rowCount -eq $RowCount
        })
    $summary = $objects | Where-Object {
        $_.kind -eq 'provider-comparison-summary' -and
        $_.provider -eq $Provider -and $_.mode -eq $Mode -and $_.rowCount -eq $RowCount
    } | Select-Object -First 1
    if ($null -eq $header -or $null -eq $summary -or $samples.Count -ne 3) {
        throw "Expected three provider baseline samples for $Provider/$Mode."
    }
    return [ordered]@{
        sourceArtifact = $BeforeProviderArtifact
        sourcePhase = $header.phase
        sourceGitCommit = $header.candidateIdentity.gitCommit
        candidateIdentity = $header.candidateIdentity
        sampleCount = $samples.Count
        samples = $samples
        metrics = [ordered]@{
            medianElapsedMilliseconds = $summary.medianElapsedMilliseconds
            medianAllocatedBytes = $summary.medianAllocatedBytes
            medianPeakWorkingSetBytes = $summary.medianPeakWorkingSetBytes
            medianRowsPerSecond = $summary.medianRowsPerSecond
        }
        note = 'Historical base-commit evidence; not rerun by this current-candidate-only confirmation.'
    }
}

function Get-BeforeRealIo {
    param([Parameter(Mandatory)][string]$Scenario)

    $objects = @(Read-JsonLines -Path $BeforeRealIoArtifact)
    $header = $objects | Where-Object { $_.kind -eq 'real-io-probe' } | Select-Object -First 1
    $child = $objects | Where-Object {
        $_.kind -eq 'real-io-child' -and $_.scenario -eq $Scenario -and $_.rowCount -eq $RowCount -and
        $null -ne $_.result -and $_.result.sampleCount -eq 3
    } | Select-Object -First 1
    if ($null -eq $header -or $null -eq $child) {
        throw "Expected three Real IO baseline samples for $Scenario."
    }
    return [ordered]@{
        sourceArtifact = $BeforeRealIoArtifact
        sourcePhase = 'before'
        sourceGitCommit = $header.gitHead
        candidateIdentity = [ordered]@{
            sourceRole = $header.sourceRole
            gitHead = $header.gitHead
            diffIdentity = $header.diffIdentity
            assemblies = $header.candidateIdentity
        }
        sampleCount = $child.result.sampleCount
        samples = $child.result.samples
        metrics = [ordered]@{
            medianElapsedMilliseconds = $child.result.medianMilliseconds
            medianAllocatedBytes = $child.result.allocatedBytesMean
            medianPeakWorkingSetBytes = $child.result.peakWorkingSetBytesMax
            outputBytesMin = $child.result.outputBytesMin
            outputBytesMax = $child.result.outputBytesMax
            asyncWriteCountMin = $child.result.asyncWriteCountMin
            asyncBytesWrittenMin = $child.result.asyncBytesWrittenMin
        }
        note = 'Historical base-commit evidence; not rerun by this current-candidate-only confirmation.'
    }
}

function Invoke-Target {
    param([Parameter(Mandatory)][hashtable]$Target)

    $temporaryPath = Join-Path $artifactDirectory ('.targeted-' + $Target.id + '-' + [guid]::NewGuid().ToString('N') + '.jsonl')
    try {
        if ($Target.kind -eq 'provider') {
            $run = Invoke-BenchmarkChild @(
                $BenchmarkDll,
                '--provider-comparison-worker',
                $temporaryPath,
                $RowCount.ToString([Globalization.CultureInfo]::InvariantCulture),
                $Repetitions.ToString([Globalization.CultureInfo]::InvariantCulture),
                $Target.provider,
                $Target.mode
            )
            $samples = @()
            if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
                $samples = @(Read-JsonLines -Path $temporaryPath)
            }
            $valid = $run.exitCode -eq 0 -and $samples.Count -eq $Repetitions
            return [ordered]@{
                id = $Target.id
                kind = $Target.kind
                provider = $Target.provider
                mode = $Target.mode
                scenario = $null
                rowCount = $RowCount
                requestedRepetitions = $Repetitions
                status = if ($valid) { 'passed' } else { 'failed' }
                exitCode = $run.exitCode
                stdout = $run.stdout
                stderr = $run.stderr
                sampleCount = $samples.Count
                samples = $samples
                metrics = if ($valid) { Get-SampleMetrics $samples } else { $null }
                exception = if ($valid) { $null } else { 'Provider worker did not return exactly five successful samples.' }
            }
        }

        $run = Invoke-BenchmarkChild @(
            $BenchmarkDll,
            '--real-io-scenario',
            $temporaryPath,
            $Target.scenario,
            $RowCount.ToString([Globalization.CultureInfo]::InvariantCulture),
            $Repetitions.ToString([Globalization.CultureInfo]::InvariantCulture)
        )
        $result = Get-LastJsonObject -Text $run.stdout
        $samples = if ($null -ne $result -and $null -ne $result.samples) { @($result.samples) } else { @() }
        $valid = $run.exitCode -eq 0 -and $null -ne $result -and $result.status -eq 'passed' -and $samples.Count -eq $Repetitions
        return [ordered]@{
            id = $Target.id
            kind = $Target.kind
            provider = $null
            mode = $null
            scenario = $Target.scenario
            rowCount = $RowCount
            requestedRepetitions = $Repetitions
            status = if ($valid) { 'passed' } else { 'failed' }
            exitCode = $run.exitCode
            stdout = $run.stdout
            stderr = $run.stderr
            sampleCount = $samples.Count
            samples = $samples
            metrics = if ($valid) { Get-SampleMetrics $samples } else { $null }
            exception = if ($valid) { $null } elseif ($null -ne $result) { $result.exception } else { 'Real IO child did not return a JSON result.' }
        }
    }
    catch {
        return [ordered]@{
            id = $Target.id
            kind = $Target.kind
            provider = $Target.provider
            mode = $Target.mode
            scenario = $Target.scenario
            rowCount = $RowCount
            requestedRepetitions = $Repetitions
            status = 'failed'
            exitCode = $null
            stdout = ''
            stderr = ''
            sampleCount = 0
            samples = @()
            metrics = $null
            exception = $_.Exception.ToString()
        }
    }
    finally {
        if (Test-Path -LiteralPath $temporaryPath -PathType Leaf) {
            Remove-Item -LiteralPath $temporaryPath -Force
        }
    }
}

$candidateIdentity = [ordered]@{
    source = 'current-benchmark-output; captured before and checked after every child'
    assemblyDirectory = $assemblyDirectory
    assemblies = Get-HashMap -Directory $assemblyDirectory
    gitHead = (& git rev-parse HEAD).Trim()
    worktreeState = 'dirty'
    legacyManifest = 'artifacts/candidates/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/manifest-current-20260921.md'
    legacyManifestNote = 'The legacy manifest hashes are not reused because the current output DLL identity differs; this artifact binds to the hashes above.'
}
$expectedHashes = @{}
foreach ($name in $candidateIdentity.assemblies.Keys) {
    $expectedHashes[$name] = $candidateIdentity.assemblies[$name]
}

$targets = @(
    @{ id = 'npoi-async'; kind = 'provider'; provider = 'npoi'; mode = 'async'; before = Get-BeforeProvider -Provider 'npoi' -Mode 'async' },
    @{ id = 'miniexcel-sync'; kind = 'provider'; provider = 'miniexcel'; mode = 'sync'; before = Get-BeforeProvider -Provider 'miniexcel' -Mode 'sync' },
    @{ id = 'excel-file-sync'; kind = 'real-io'; scenario = 'excel-file-sync'; before = Get-BeforeRealIo -Scenario 'excel-file-sync' },
    @{ id = 'excel-file-async'; kind = 'real-io'; scenario = 'excel-file-async'; before = Get-BeforeRealIo -Scenario 'excel-file-async' },
    @{ id = 'excel-throttled-async'; kind = 'real-io'; scenario = 'excel-throttled-async'; before = Get-BeforeRealIo -Scenario 'excel-throttled-async' }
)

$results = @()
foreach ($target in $targets) {
    Assert-CandidateStable -Expected $expectedHashes
    $after = Invoke-Target -Target $target
    Assert-CandidateStable -Expected $expectedHashes
    $beforeMetrics = $target.before.metrics
    $afterMetrics = $after.metrics
    $delta = if ($null -ne $afterMetrics) {
        [ordered]@{
            medianElapsedMilliseconds = Get-PercentDelta $beforeMetrics.medianElapsedMilliseconds $afterMetrics.medianElapsedMilliseconds
            medianAllocatedBytes = Get-PercentDelta $beforeMetrics.medianAllocatedBytes $afterMetrics.medianAllocatedBytes
            medianPeakWorkingSetBytes = Get-PercentDelta $beforeMetrics.medianPeakWorkingSetBytes $afterMetrics.medianPeakWorkingSetBytes
        }
    } else {
        $null
    }
    $budgetConclusion = if ($after.status -ne 'passed') {
        'INCONCLUSIVE_FAILED'
    } elseif ($null -eq $delta.medianElapsedMilliseconds) {
        'INCONCLUSIVE_NO_BASELINE'
    } elseif ($delta.medianElapsedMilliseconds -gt 25) {
        'CONFIRMED_OVER_25_PERCENT'
    } else {
        'NOT_CONFIRMED_OVER_25_PERCENT'
    }
    $results += [ordered]@{
        id = $target.id
        workload = if ($target.kind -eq 'provider') { "provider-comparison:$($target.provider):$($target.mode)" } else { "real-io:$($target.scenario)" }
        rowCount = $RowCount
        requestedRepetitions = $Repetitions
        before = $target.before
        after = $after
        deltaPercent = $delta
        budgetV1Conclusion = $budgetConclusion
    }
}

$confirmed = @($results | Where-Object { $_.budgetV1Conclusion -eq 'CONFIRMED_OVER_25_PERCENT' }).Count
$failed = @($results | Where-Object { $_.after.status -ne 'passed' }).Count
$overall = if ($failed -gt 0) { 'INCONCLUSIVE_FAILED' } elseif ($confirmed -gt 0) { 'CONFIRMED_TARGETED_REGRESSIONS' } else { 'NO_TARGETED_REGRESSION_CONFIRMED' }
Assert-CandidateStable -Expected $expectedHashes

$document = [ordered]@{
    kind = 'targeted-performance-budget-v1'
    schema = 1
    generatedUtc = [DateTimeOffset]::UtcNow
    scope = 'Only the five approved 100K E2E regressions; no A-I, Relation, RawDate, tail-latency, c16 or c64 workloads.'
    rowCount = $RowCount
    requestedRepetitions = $Repetitions
    beforeSamplePolicy = 'Existing base-commit evidence is reused without rerun; each before series is exactly the three samples present in its cited artifact.'
    afterSamplePolicy = 'Current candidate only; every targeted after series must contain exactly five samples.'
    candidateIdentity = $candidateIdentity
    budgetV1 = [ordered]@{
        thresholdPercent = 25
        metric = 'medianElapsedMilliseconds'
        conclusion = $overall
        confirmedScenarioCount = $confirmed
        failedScenarioCount = $failed
        approvalStatus = 'APPROVED_TARGETED_CONFIRMATION_ONLY'
        note = 'This confirms only whether the selected historical >25% observations remain over 25%; it is not a full performance gate or release approval.'
    }
    scenarios = $results
}
$json = $document | ConvertTo-Json -Depth 64
[IO.File]::WriteAllText($ArtifactPath, $json + [Environment]::NewLine, [Text.UTF8Encoding]::new($false))
Write-Output "TARGETED_PERFORMANCE_BUDGET artifact=$ArtifactPath scenarios=$($results.Count) repetitions=$Repetitions conclusion=$overall"
if ($failed -gt 0) {
    exit 1
}
