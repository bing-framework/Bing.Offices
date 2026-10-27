#requires -Version 7.0
[CmdletBinding()]
param(
    [string] $OutputDirectory = 'artifacts/package-candidate'
)

[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false)
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false
Set-StrictMode -Version Latest

$repository = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..'))
$evidenceRoot = [IO.Path]::GetFullPath($OutputDirectory, $repository)
if (Test-Path -LiteralPath $evidenceRoot) {
    throw "Use a new output directory to avoid mixing candidates: $evidenceRoot"
}
$null = New-Item -ItemType Directory -Path $evidenceRoot
$logRoot = Join-Path $evidenceRoot 'logs'
$packageRoot = Join-Path $evidenceRoot 'packages'
$consumerEvidence = Join-Path $evidenceRoot 'consumers'
$null = New-Item -ItemType Directory -Path $logRoot, $packageRoot, $consumerEvidence
$utf8 = [Text.UTF8Encoding]::new($false)
$temporaryParent = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$temporaryRoot = Join-Path $temporaryParent ('bing-offices-consumer-' + [Guid]::NewGuid().ToString('N'))
$previousPackageVersion = $env:BING_OFFICES_PACKAGE_VERSION
$outcome = [ordered]@{ status = 'IN_PROGRESS'; consumers = @(); restoredSources = @(); negativeChecks = @(); failure = $null }

function Invoke-DotNet {
    param([string] $Name, [string[]] $ArgumentList, [string] $ExpectedFailure = '')
    $log = [Text.StringBuilder]::new()
    Write-Host "[$Name] dotnet $($ArgumentList -join ' ')"
    & dotnet @ArgumentList 2>&1 | ForEach-Object {
        $line = $_.ToString()
        $null = $log.AppendLine($line)
        Write-Host $line
    }
    $exitCode = $LASTEXITCODE
    [IO.File]::WriteAllText((Join-Path $logRoot "$Name.log"), $log.ToString(), $utf8)
    if ($ExpectedFailure) {
        if ($exitCode -eq 0 -or -not $log.ToString().Contains($ExpectedFailure)) {
            throw "$Name did not reject the fixture with the expected diagnostic: $ExpectedFailure (exit=$exitCode)"
        }
        return
    }
    if ($exitCode -ne 0) { throw "$Name failed with exit code $exitCode. See $Name.log." }
}

Push-Location $repository
try {
    # 全部包来自同一轮构建，版本取真实 nuspec，不覆盖项目版本或批准基线。
    foreach ($id in @('Bing.Offices.Abstractions', 'Bing.Offices.Core', 'Bing.Offices.Npoi')) {
        Invoke-DotNet "pack-$id" @('pack', "src/$id/$id.csproj", '-c', 'Release', '-o', $packageRoot)
    }
    $packages = @(Get-ChildItem -LiteralPath $packageRoot -Filter '*.nupkg')
    $symbols = @(Get-ChildItem -LiteralPath $packageRoot -Filter '*.snupkg')
    if ($packages.Count -ne 3 -or $symbols.Count -ne 3) { throw 'Expected three nupkg and three snupkg files.' }
    $versions = @()
    foreach ($package in $packages) {
        $archive = [IO.Compression.ZipFile]::OpenRead($package.FullName)
        try {
            $names = @($archive.Entries | ForEach-Object FullName)
            if (-not ($names -match '^lib/.+\.dll$') -or -not ($names -match '^lib/.+\.xml$')) {
                throw "Package is missing its library or XML documentation: $($package.Name)"
            }
            $nuspec = @($archive.Entries | Where-Object { $_.FullName.EndsWith('.nuspec') })
            if ($nuspec.Count -ne 1) { throw "Expected one nuspec: $($package.Name)" }
            $reader = [IO.StreamReader]::new($nuspec[0].Open(), $utf8, $true)
            try { [xml] $metadata = $reader.ReadToEnd() } finally { $reader.Dispose() }
            $versions += $metadata.SelectSingleNode("/*[local-name()='package']/*[local-name()='metadata']/*[local-name()='version']").InnerText
        }
        finally { $archive.Dispose() }
    }
    $versions = @($versions | Select-Object -Unique)
    if ($versions.Count -ne 1) { throw 'Production packages have different versions.' }
    $packageVersion = $versions[0]
    $outcome['packageVersion'] = $packageVersion
    $env:BING_OFFICES_PACKAGE_VERSION = $packageVersion

    Invoke-DotNet 'build-api-tool' @('build', 'build/ApiSnapshot/ApiSnapshot.csproj', '-c', 'Release')
    $tool = Join-Path $repository 'build/ApiSnapshot/bin/Release/net8.0/ApiSnapshot.dll'
    Invoke-DotNet 'identity-contract' @($tool, '--identity-self-test', 'true')
    $captureRoot = Join-Path $evidenceRoot 'capture'
    $baseline = Join-Path $repository 'build/api-snapshot-baseline.json'
    $apiArguments = @('--root', (Join-Path $repository 'output/release'), '--packages', $packageRoot, '--repository', $repository)
    Invoke-DotNet 'capture' (@($tool) + $apiArguments + @('--capture', 'true', '--output', $captureRoot))
    $candidate = Join-Path $captureRoot 'api-candidate.json'

    # 在仓库外创建消费者，阻止源码引用和既有同版本包缓存掩盖打包缺陷。
    $null = New-Item -ItemType Directory -Path $temporaryRoot
    [IO.File]::WriteAllText((Join-Path $temporaryRoot 'Directory.Build.props'), '<Project />', $utf8)
    [IO.File]::WriteAllText((Join-Path $temporaryRoot 'Directory.Build.targets'), '<Project />', $utf8)
    foreach ($name in @('Bing.Offices.Consumer.Net6', 'Bing.Offices.Consumer.Net8', 'Bing.Offices.Consumer.Core', 'Bing.Offices.Consumer.Shared')) {
        $source = Join-Path $repository "tests/$name"
        $target = Join-Path $temporaryRoot $name
        $null = New-Item -ItemType Directory -Path $target
        Get-ChildItem -LiteralPath $source -File | Where-Object { $_.Extension -in @('.cs', '.csproj') } |
            Copy-Item -Destination $target
    }
    $feed = [Security.SecurityElement]::Escape($packageRoot)
    $config = Join-Path $temporaryRoot 'NuGet.Config'
    [IO.File]::WriteAllText($config, @"
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources><clear /><add key="candidate" value="$feed" /><add key="nuget.org" value="https://api.nuget.org/v3/index.json" protocolVersion="3" /></packageSources>
  <packageSourceMapping>
    <packageSource key="candidate"><package pattern="Bing.Offices.*" /></packageSource>
    <packageSource key="nuget.org"><package pattern="*" /></packageSource>
  </packageSourceMapping>
</configuration>
"@, $utf8)
    $targets = @(
        @{ Name = 'Bing.Offices.Consumer.Net6'; Frameworks = @('net6.0') },
        @{ Name = 'Bing.Offices.Consumer.Net8'; Frameworks = @('net8.0') },
        @{ Name = 'Bing.Offices.Consumer.Core'; Frameworks = @('net6.0', 'net8.0') }
    )
    foreach ($target in $targets) {
        $name = $target.Name
        $project = Join-Path $temporaryRoot "$name/$name.csproj"
        $cache = Join-Path $temporaryRoot "cache/$name"
        Invoke-DotNet "restore-$name" @('restore', $project, '--configfile', $config, '--packages', $cache, "-p:BingOfficesPackageVersion=$packageVersion")
        $assetsPath = Join-Path $temporaryRoot "$name/obj/project.assets.json"
        $assets = Get-Content -LiteralPath $assetsPath -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable
        if (@($assets.libraries.Values | Where-Object { $_.type -eq 'project' }).Count -ne 0) {
            throw "Consumer resolved a ProjectReference: $name"
        }
        if ($name -eq 'Bing.Offices.Consumer.Core' -and @($assets.libraries.Keys | Where-Object { $_ -match '(?i)npoi' }).Count -ne 0) {
            throw 'Core-only consumer unexpectedly depends on NPOI.'
        }
        # 包源映射之外再检查实际恢复来源，不能只信成功的构建输出。
        foreach ($library in @($assets.libraries.Keys | Where-Object { $_ -like 'Bing.Offices.*/*' })) {
            $relative = $assets.libraries[$library].path
            $packageMetadata = Get-Content -LiteralPath (Join-Path $cache "$relative/.nupkg.metadata") -Raw -Encoding utf8 | ConvertFrom-Json
            if ([IO.Path]::GetFullPath($packageMetadata.source).TrimEnd('/', '\') -ne $packageRoot.TrimEnd('/', '\')) {
                throw "Office package was not restored from this candidate: $library"
            }
            $outcome.restoredSources += @{ consumer = $name; package = $library; source = $packageMetadata.source }
        }
        Copy-Item -LiteralPath $assetsPath -Destination (Join-Path $consumerEvidence "$name.assets.json")
        foreach ($tfm in $target.Frameworks) {
            Invoke-DotNet "build-$name-$tfm" @('build', $project, '-c', 'Release', '-f', $tfm, '--no-restore', "-p:BingOfficesPackageVersion=$packageVersion")
            Invoke-DotNet "run-$name-$tfm" @('run', '--project', $project, '-c', 'Release', '-f', $tfm, '--no-build', '--no-restore', "-p:BingOfficesPackageVersion=$packageVersion")
            $outcome.consumers += "$name/$tfm"
        }
    }

    $comparison = Join-Path $evidenceRoot 'comparison'
    Invoke-DotNet 'compare' (@($tool) + $apiArguments + @('--baseline', $baseline, '--candidate', $candidate, '--output', $comparison))

    # 反例验证：候选清单不得绕过成员基线，错误身份也不得静默退回基线。
    $negativeRoot = Join-Path $evidenceRoot 'negative-cases'
    $null = New-Item -ItemType Directory -Path $negativeRoot
    Invoke-DotNet 'reject-missing-candidate' (@($tool) + $apiArguments + @('--baseline', $baseline, '--candidate', (Join-Path $negativeRoot 'missing.json'))) 'API candidate manifest is missing'
    $outcome.negativeChecks += 'missing-candidate'
    $badCandidate = Get-Content -LiteralPath $candidate -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable
    $badCandidate.candidateIdentity.assemblyFiles['net8.0/Bing.Offices.Npoi.dll'] = '0' * 64
    $badCandidatePath = Join-Path $negativeRoot 'bad-candidate.json'
    [IO.File]::WriteAllText($badCandidatePath, ($badCandidate | ConvertTo-Json -Depth 30), $utf8)
    Invoke-DotNet 'reject-bad-identity' (@($tool) + $apiArguments + @('--baseline', $baseline, '--candidate', $badCandidatePath)) 'assembly net8.0/Bing.Offices.Npoi.dll SHA-256 mismatch'
    $outcome.negativeChecks += 'bad-identity'
    $badBaseline = Get-Content -LiteralPath $baseline -Raw -Encoding utf8 | ConvertFrom-Json -AsHashtable
    $badBaseline.assemblies['net8.0']['Bing.Offices.Core'].hash = '0' * 64
    $badBaseline.assemblies['net8.0']['Bing.Offices.Core'].lines += 'UNAPPROVED_MEMBER'
    $badBaselinePath = Join-Path $negativeRoot 'bad-baseline.json'
    [IO.File]::WriteAllText($badBaselinePath, ($badBaseline | ConvertTo-Json -Depth 30), $utf8)
    Invoke-DotNet 'reject-api-mismatch' (@($tool) + $apiArguments + @('--baseline', $badBaselinePath, '--candidate', $candidate)) 'net8.0/Bing.Offices.Core: expected='
    $outcome.negativeChecks += 'api-mismatch'

    $identity = (Get-Content -LiteralPath $candidate -Raw -Encoding utf8 | ConvertFrom-Json).candidateIdentity
    $outcome['candidateIdentity'] = $identity
    $outcome.status = 'PASS'
    Write-Host "package-candidate-ok consumers=$($outcome.consumers.Count) negativeChecks=$($outcome.negativeChecks.Count)"
}
catch {
    $outcome.status = 'FAIL'
    $outcome.failure = $_.Exception.Message
    throw
}
finally {
    $env:BING_OFFICES_PACKAGE_VERSION = $previousPackageVersion
    [IO.File]::WriteAllText((Join-Path $evidenceRoot 'summary.json'), ($outcome | ConvertTo-Json -Depth 30), $utf8)
    # 仅清理本脚本创建、且仍位于系统临时目录中的专属消费者目录。
    $resolvedTemporary = [IO.Path]::GetFullPath($temporaryRoot)
    $expectedParent = [IO.Path]::TrimEndingDirectorySeparator($temporaryParent)
    if ([IO.Path]::GetDirectoryName($resolvedTemporary) -eq $expectedParent -and
        [IO.Path]::GetFileName($resolvedTemporary).StartsWith('bing-offices-consumer-') -and
        (Test-Path -LiteralPath $resolvedTemporary)) {
        Remove-Item -LiteralPath $resolvedTemporary -Recurse -Force
    }
    Pop-Location
}
