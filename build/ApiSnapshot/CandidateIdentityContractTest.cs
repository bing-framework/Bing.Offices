using System.Diagnostics;
using System.Text;
using System.IO.Compression;
using Bing.Offices.ApiSnapshot;

/// <summary>
/// 验证候选程序集、NuGet 包和源清单身份合同。
/// </summary>
internal static class CandidateIdentityContractTest
{
    /// <summary>
    /// 候选身份合同使用的生成器版本。
    /// </summary>
    private const string GeneratorVersion = "2.7.0";
    /// <summary>
    /// 本测试覆盖的自定义 API 审批文件路径。
    /// </summary>
    private const string CustomApprovalPath =
        "ai_docs/tasks/BING-OFFICES-RC-API-PERF-ENTITY-HARDENING-20260920-001/api-approval-request.md";

    /// <summary>
    /// 运行候选身份合同测试。
    /// </summary>
    /// <returns>成功返回 0，失败返回 1。</returns>
    public static int Run()
    {
        var temporaryRoot = Path.Combine(Path.GetTempPath(), $"bing-offices-api-identity-{Guid.NewGuid():N}");
        var repository = Path.Combine(temporaryRoot, "source");
        var linuxCheckout = Path.Combine(temporaryRoot, "linux-checkout");
        var windowsCheckout = Path.Combine(temporaryRoot, "windows-checkout");

        try
        {
            Directory.CreateDirectory(repository);
            InitializeRepository(repository);

            var trackedSourcePath = Path.Combine(repository, "src", "Tracked.cs");
            WriteBytes(trackedSourcePath, Utf8("class TrackedCandidate { }\r\n"));
            WriteText(Path.Combine(repository, "common.props"), "<Project />\r\n");
            WriteText(Path.Combine(repository, "framework.props"), "<Project />\r\n");
            WriteText(Path.Combine(repository, "asset", "props", "package.props"), "<Project />\r\n");
            WriteText(Path.Combine(repository, "README.md"), "package readme\r\n");
            WriteText(Path.Combine(repository, "LICENSE"), "license\r\n");
            WriteBytes(Path.Combine(repository, "tests", "Base.lock"), Utf8("base-lock\r\n"));
            RunGit(repository, "add", "--all");
            RunGit(repository, "commit", "--quiet", "--no-gpg-sign", "-m", "base");
            var baseCommit = RunGit(repository, "rev-parse", "HEAD").Trim();

            WriteBytes(trackedSourcePath, Utf8("class TrackedCandidate { public int Value => 1; }\r\n"));
            WriteBytes(Path.Combine(repository, "tests", "Candidate.cs"),
                Utf8("class CandidateSource { }\r\n"));
            WriteBytes(Path.Combine(repository, "tests", "Candidate.lock"),
                Utf8("candidate-lock\r\n"));

            var assemblyRoot = Path.Combine(repository, "output");
            var assemblyPaths = new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["netstandard2.0/Bing.Offices.Abstractions.dll"] =
                    Path.Combine(assemblyRoot, "netstandard2.0", "Bing.Offices.Abstractions.dll"),
                ["netstandard2.0/Bing.Offices.Core.dll"] =
                    Path.Combine(assemblyRoot, "netstandard2.0", "Bing.Offices.Core.dll"),
                ["net6.0/Bing.Offices.Npoi.dll"] =
                    Path.Combine(assemblyRoot, "net6.0", "Bing.Offices.Npoi.dll"),
                ["net8.0/Bing.Offices.Npoi.dll"] =
                    Path.Combine(assemblyRoot, "net8.0", "Bing.Offices.Npoi.dll"),
                ["net6.0/Bing.Offices.MiniExcel.dll"] =
                    Path.Combine(assemblyRoot, "net6.0", "Bing.Offices.MiniExcel.dll"),
                ["net8.0/Bing.Offices.MiniExcel.dll"] =
                    Path.Combine(assemblyRoot, "net8.0", "Bing.Offices.MiniExcel.dll"),
                ["net6.0/Bing.Offices.ClosedXml.dll"] =
                    Path.Combine(assemblyRoot, "net6.0", "Bing.Offices.ClosedXml.dll"),
                ["net8.0/Bing.Offices.ClosedXml.dll"] =
                    Path.Combine(assemblyRoot, "net8.0", "Bing.Offices.ClosedXml.dll")
            };
            var managedAssemblyPath = typeof(CandidateIdentityContractTest).Assembly.Location;
            foreach (var path in assemblyPaths.Values)
                CopyFile(managedAssemblyPath, path);

            var packageRoot = Path.Combine(repository, "packages");
            var packageNames = new[]
            {
                "Bing.Offices.Abstractions.2.0.0.nupkg",
                "Bing.Offices.Core.2.0.0.nupkg",
                "Bing.Offices.Npoi.2.0.0.nupkg",
                "Bing.Offices.MiniExcel.2.0.0.nupkg",
                "Bing.Offices.ClosedXml.2.0.0.nupkg"
            };
            foreach (var packageName in packageNames)
                CreatePackage(Path.Combine(packageRoot, packageName), packageName, assemblyPaths);

            var approvalPath = Path.Combine(repository,
                CandidateIdentityVerifier.BreakingApprovalPath.Replace('/', Path.DirectorySeparatorChar));
            WriteText(approvalPath, "approved\r\n");
            var customApprovalPath = Path.Combine(repository,
                CustomApprovalPath.Replace('/', Path.DirectorySeparatorChar));
            WriteText(customApprovalPath, "custom approved\r\n");

            var ignoredArtifactPath = Path.Combine(repository, "ai_docs", "tasks",
                "BO-RC-20260908-002", "artifacts", "ignored.json");
            var ignoredLockfilePath = Path.Combine(repository, "tests", "packages.lock.json");

            var captureFailures = new List<string>();
            var identity = CandidateIdentityVerifier.Capture(
                assemblyPaths, packageRoot, repository, baseCommit, captureFailures, CustomApprovalPath);
            AssertNoFailures("dirty capture", captureFailures);
            if (!string.Equals(identity.BreakingApprovalArtifact, CustomApprovalPath, StringComparison.Ordinal))
                throw new InvalidOperationException("custom approval path was not normalized into candidate identity.");

            var defaultCaptureFailures = new List<string>();
            var defaultIdentity = CandidateIdentityVerifier.Capture(
                assemblyPaths, packageRoot, repository, baseCommit, defaultCaptureFailures);
            AssertNoFailures("default approval path capture", defaultCaptureFailures);
            if (!string.Equals(defaultIdentity.BreakingApprovalArtifact,
                    CandidateIdentityVerifier.BreakingApprovalPath, StringComparison.Ordinal))
                throw new InvalidOperationException("default approval path compatibility was lost.");

            var sourceManifestHash = identity.CandidateSourceManifestSha256;
            WriteBytes(ignoredArtifactPath, Utf8("ignored artifact\r\n"));
            WriteBytes(ignoredLockfilePath, Utf8("ignored lockfile\r\n"));
            var ignoredCaptureFailures = new List<string>();
            var ignoredCapture = CandidateIdentityVerifier.Capture(
                assemblyPaths, packageRoot, repository, baseCommit, ignoredCaptureFailures, CustomApprovalPath);
            AssertNoFailures("ignored file capture", ignoredCaptureFailures);
            if (!string.Equals(sourceManifestHash, ignoredCapture.CandidateSourceManifestSha256,
                    StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "ignored task artifacts and packages.lock.json changed the candidate source manifest.");

            WriteText(Path.Combine(repository, "docs", "release-note.md"), "documentation only\r\n");
            WriteBytes(Path.Combine(repository, "tests", "Candidate.lock"), Utf8("candidate-lock-2\r\n"));
            WriteText(Path.Combine(repository, ".github", "workflows", "ci.yml"), "name: CI\r\n");
            var releaseUnrelatedCaptureFailures = new List<string>();
            var releaseUnrelatedCapture = CandidateIdentityVerifier.Capture(
                assemblyPaths, packageRoot, repository, baseCommit, releaseUnrelatedCaptureFailures,
                CustomApprovalPath);
            AssertNoFailures("release-unrelated file capture", releaseUnrelatedCaptureFailures);
            if (!string.Equals(sourceManifestHash, releaseUnrelatedCapture.CandidateSourceManifestSha256,
                    StringComparison.Ordinal))
                throw new InvalidOperationException(
                    "documentation, tests or CI configuration changed the candidate source manifest.");

            var baseline = new ApiBaselineDocument
            {
                Schema = "bing.offices.public-api.v2",
                GeneratorVersion = GeneratorVersion,
                BaselineCommit = baseCommit,
                ApprovedBy = "api-identity-test",
                ApprovedAt = "2026-09-11T00:00:00Z",
                CandidateIdentity = identity
            };

            AssertNoFailures("dirty validation",
                Validate(baseline, assemblyPaths, packageRoot, repository));
            AssertValidationFails("approval path mismatch",
                baseline, assemblyPaths, packageRoot, repository,
                CandidateIdentityVerifier.BreakingApprovalPath);
            AssertValidationFails("approval path traversal",
                baseline, assemblyPaths, packageRoot, repository, "../outside.md");

            var repositoryMetadataPackageRoot = Path.Combine(repository, "repository-metadata-packages");
            foreach (var packageName in packageNames)
                CreatePackage(Path.Combine(repositoryMetadataPackageRoot, packageName), packageName,
                    assemblyPaths, repositoryCommit: "candidate");
            AssertNoFailures("repository metadata package validation",
                Validate(baseline, assemblyPaths, repositoryMetadataPackageRoot, repository));

            var equivalentXmlPackageRoot = Path.Combine(repository, "equivalent-xml-packages");
            foreach (var packageName in packageNames)
                CreatePackage(Path.Combine(equivalentXmlPackageRoot, packageName), packageName,
                    assemblyPaths, usePlatformXmlFormatting: true);
            AssertNoFailures("equivalent XML package validation",
                Validate(baseline, assemblyPaths, equivalentXmlPackageRoot, repository));

            var dllScenarios = new (string Name, byte[]? Content, string? ReplacementPath)[]
            {
                ("corrupt", Utf8("not a managed DLL"), null),
                ("missing", null, null),
                ("wrong managed assembly", File.ReadAllBytes(typeof(ZipArchive).Assembly.Location), null),
                ("wrong TFM", File.ReadAllBytes(managedAssemblyPath), "lib/net9.0/Bing.Offices.MiniExcel.dll")
            };
            foreach (var scenario in dllScenarios)
            {
                var scenarioRoot = Path.Combine(temporaryRoot, scenario.Name + "-dll-packages");
                foreach (var packageName in packageNames)
                    CopyFile(Path.Combine(packageRoot, packageName), Path.Combine(scenarioRoot, packageName));
                var changedPackage = Path.Combine(scenarioRoot, packageNames[3]);
                ReplacePackageDll(changedPackage, "lib/net8.0/Bing.Offices.MiniExcel.dll", scenario.Content,
                    scenario.ReplacementPath);
                var scenarioFailures = Validate(baseline, assemblyPaths, scenarioRoot, repository);
                if (!scenarioFailures.Any(failure => failure.Contains("nupkg " + packageNames[3], StringComparison.Ordinal)))
                    throw new InvalidOperationException($"{scenario.Name} package DLL did not fail package identity verification.");
                Console.WriteLine($"Package DLL identity scenario rejected as expected: {scenario.Name}");
            }

            var tamperedXmlPackageRoot = Path.Combine(repository, "tampered-xml-packages");
            foreach (var packageName in packageNames)
                CreatePackage(Path.Combine(tamperedXmlPackageRoot, packageName), packageName,
                    assemblyPaths, documentationSummary: "tampered");
            AssertValidationFails("package XML semantic tamper",
                baseline, assemblyPaths, tamperedXmlPackageRoot, repository);

            var tamperedRepositoryPackageRoot = Path.Combine(repository, "tampered-repository-packages");
            foreach (var packageName in packageNames)
                CreatePackage(Path.Combine(tamperedRepositoryPackageRoot, packageName), packageName,
                    assemblyPaths, repositoryCommit: "candidate", repositoryUrl: "https://example.invalid/tampered");
            AssertValidationFails("repository URL package tamper",
                baseline, assemblyPaths, tamperedRepositoryPackageRoot, repository);

            RunGit(repository, "add", "--all");
            RunGit(repository, "commit", "--quiet", "--no-gpg-sign", "-m", "candidate");
            var candidateCommit = RunGit(repository, "rev-parse", "HEAD").Trim();
            if (string.Equals(baseCommit, candidateCommit, StringComparison.Ordinal))
                throw new InvalidOperationException("candidate commit did not advance beyond the base commit.");

            var linuxPaths = ValidateCleanCheckout(
                repository, linuxCheckout, "false", baseline, assemblyPaths, packageRoot);
            var windowsPaths = ValidateCleanCheckout(
                repository, windowsCheckout, "true", baseline, assemblyPaths, packageRoot);
            if (!string.Equals(
                    RunGit(linuxCheckout, "rev-parse", "HEAD").Trim(), candidateCommit,
                    StringComparison.Ordinal)
                || !string.Equals(
                    RunGit(windowsCheckout, "rev-parse", "HEAD").Trim(), candidateCommit,
                    StringComparison.Ordinal))
                throw new InvalidOperationException("fresh checkout did not resolve the candidate commit.");

            var originalSource = File.ReadAllBytes(Path.Combine(linuxCheckout, "src", "Tracked.cs"));
            WriteBytes(Path.Combine(linuxCheckout, "src", "Tracked.cs"),
                Utf8("class TrackedCandidate { public int Value => 2; }\n"));
            AssertValidationFails("tracked source tamper",
                baseline, linuxPaths.Assemblies, linuxPaths.PackageRoot, linuxCheckout);
            WriteBytes(Path.Combine(linuxCheckout, "src", "Tracked.cs"), originalSource);

            var rootPropsPath = Path.Combine(linuxCheckout, "common.props");
            var originalRootProps = File.ReadAllBytes(rootPropsPath);
            WriteText(rootPropsPath, "<Project><PropertyGroup><Tampered>true</Tampered></PropertyGroup></Project>\n");
            AssertValidationFails("root props tamper",
                baseline, linuxPaths.Assemblies, linuxPaths.PackageRoot, linuxCheckout);
            WriteBytes(rootPropsPath, originalRootProps);

            var packagePropsPath = Path.Combine(linuxCheckout, "asset", "props", "package.props");
            var originalPackageProps = File.ReadAllBytes(packagePropsPath);
            WriteText(packagePropsPath, "<Project><PropertyGroup><Tampered>true</Tampered></PropertyGroup></Project>\n");
            AssertValidationFails("package asset tamper",
                baseline, linuxPaths.Assemblies, linuxPaths.PackageRoot, linuxCheckout);
            WriteBytes(packagePropsPath, originalPackageProps);

            var untrackedSourcePath = Path.Combine(linuxCheckout, "src", "UntrackedTamper.cs");
            WriteText(untrackedSourcePath, "class UntrackedTamper { }\n");
            AssertValidationFails("untracked source tamper",
                baseline, linuxPaths.Assemblies, linuxPaths.PackageRoot, linuxCheckout);
            File.Delete(untrackedSourcePath);

            var missingApprovalPath = Path.Combine(linuxCheckout,
                CustomApprovalPath.Replace('/', Path.DirectorySeparatorChar));
            var originalApproval = File.ReadAllBytes(missingApprovalPath);
            File.Delete(missingApprovalPath);
            AssertValidationFails("missing approval file",
                baseline, linuxPaths.Assemblies, linuxPaths.PackageRoot, linuxCheckout);
            WriteBytes(missingApprovalPath, originalApproval);

            var assemblyPath = linuxPaths.Assemblies["net8.0/Bing.Offices.Npoi.dll"];
            var originalAssembly = File.ReadAllBytes(assemblyPath);
            WriteBytes(assemblyPath, Utf8("tampered assembly"));
            AssertValidationFails("assembly tamper",
                baseline, linuxPaths.Assemblies, linuxPaths.PackageRoot, linuxCheckout);
            WriteBytes(assemblyPath, originalAssembly);

            var packagePath = Path.Combine(linuxPaths.PackageRoot, packageNames[2]);
            var originalPackage = File.ReadAllBytes(packagePath);
            WriteBytes(packagePath, Utf8("tampered package"));
            AssertValidationFails("nupkg tamper",
                baseline, linuxPaths.Assemblies, linuxPaths.PackageRoot, linuxCheckout);
            WriteBytes(packagePath, originalPackage);

            var checkoutApprovalPath = Path.Combine(linuxCheckout,
                CustomApprovalPath.Replace('/', Path.DirectorySeparatorChar));
            originalApproval = File.ReadAllBytes(checkoutApprovalPath);
            WriteText(checkoutApprovalPath, "tampered approval\n");
            AssertValidationFails("approval tamper",
                baseline, linuxPaths.Assemblies, linuxPaths.PackageRoot, linuxCheckout);
            WriteBytes(checkoutApprovalPath, originalApproval);

            Console.WriteLine(
                "API identity clean-checkout contract passed for LF and CRLF checkouts; "
                + "source, assembly, nupkg and approval tamper checks failed as expected; "
                + "ignored artifacts, lockfiles and release-unrelated files were excluded.");
            return 0;
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine($"API identity clean-checkout contract failed: {exception.Message}");
            return 1;
        }
        finally
        {
            DeleteTree(temporaryRoot);
        }
    }

    /// <summary>
    /// 创建并校验指定换行配置下的干净检出副本。
    /// </summary>
    /// <param name="sourceRepository">源仓库根目录。</param>
    /// <param name="checkout">检出副本目录。</param>
    /// <param name="autocrlf">Git 的自动换行配置值。</param>
    /// <param name="baseline">候选身份基线文档。</param>
    /// <param name="sourceAssemblies">源仓库构建程序集路径。</param>
    /// <param name="sourcePackageRoot">源仓库 NuGet 包目录。</param>
    /// <returns>检出副本中的程序集和包路径。</returns>
    private static CheckoutPaths ValidateCleanCheckout(
        string sourceRepository,
        string checkout,
        string autocrlf,
        ApiBaselineDocument baseline,
        IReadOnlyDictionary<string, string> sourceAssemblies,
        string sourcePackageRoot)
    {
        RunGit(sourceRepository, "clone", "--quiet", "--no-local", sourceRepository, checkout);
        RunGit(checkout, "config", "core.autocrlf", autocrlf);
        RunGit(checkout, "reset", "--hard", "--quiet", "HEAD");
        if (!string.IsNullOrWhiteSpace(RunGit(checkout, "status", "--porcelain=v1", "--untracked-files=all").Trim()))
            throw new InvalidOperationException($"fresh checkout is not clean: {checkout}");

        var assemblies = sourceAssemblies.ToDictionary(
            pair => pair.Key,
            pair => Path.Combine(checkout, Path.GetRelativePath(sourceRepository, pair.Value)),
            StringComparer.Ordinal);
        var packageRoot = Path.Combine(checkout, Path.GetRelativePath(sourceRepository, sourcePackageRoot));
        CopyGeneratedFiles(sourceRepository, checkout, "output");
        CopyGeneratedFiles(sourceRepository, checkout, "packages");

        AssertNoFailures($"{autocrlf} clean checkout validation",
            Validate(baseline, assemblies, packageRoot, checkout));
        return new CheckoutPaths(assemblies, packageRoot);
    }

    /// <summary>
    /// 初始化候选身份合同使用的临时 Git 仓库。
    /// </summary>
    /// <param name="repository">待初始化的仓库根目录。</param>
    private static void InitializeRepository(string repository)
    {
        RunGit(repository, "init", "--quiet");
        RunGit(repository, "config", "user.email", "api-identity-test@example.invalid");
        RunGit(repository, "config", "user.name", "API Identity Contract Test");
        WriteText(Path.Combine(repository, ".gitattributes"), "* text=auto\n");
        WriteText(Path.Combine(repository, ".gitignore"),
            "output/\n"
            + "packages/\n"
            + "ai_docs/tasks/**/artifacts/**\n"
            + "packages.lock.json\n");
    }

    /// <summary>
    /// 使用候选身份校验临时仓库和构建资产。
    /// </summary>
    /// <param name="baseline">候选身份基线文档。</param>
    /// <param name="assemblyPaths">待校验的程序集路径。</param>
    /// <param name="packageRoot">NuGet 包目录。</param>
    /// <param name="repository">候选仓库根目录。</param>
    /// <param name="approvalPath">审批文件相对路径；省略时使用默认路径。</param>
    /// <returns>校验失败信息集合；没有失败时为空。</returns>
    private static List<string> Validate(
        ApiBaselineDocument baseline,
        IReadOnlyDictionary<string, string> assemblyPaths,
        string packageRoot,
        string repository,
        string? approvalPath = CustomApprovalPath)
    {
        var failures = new List<string>();
        CandidateIdentityVerifier.Validate(
            baseline, assemblyPaths, packageRoot, repository, failures, approvalPath);
        return failures;
    }

    /// <summary>
    /// 断言指定候选身份校验必须失败。
    /// </summary>
    /// <param name="label">断言场景名称。</param>
    /// <param name="baseline">候选身份基线文档。</param>
    /// <param name="assemblyPaths">待校验的程序集路径。</param>
    /// <param name="packageRoot">NuGet 包目录。</param>
    /// <param name="repository">候选仓库根目录。</param>
    /// <param name="approvalPath">审批文件相对路径；省略时使用默认路径。</param>
    private static void AssertValidationFails(
        string label,
        ApiBaselineDocument baseline,
        IReadOnlyDictionary<string, string> assemblyPaths,
        string packageRoot,
        string repository,
        string? approvalPath = CustomApprovalPath)
    {
        var failures = Validate(baseline, assemblyPaths, packageRoot, repository, approvalPath);
        if (failures.Count == 0)
            throw new InvalidOperationException($"{label} did not fail validation.");
    }

    /// <summary>
    /// 断言指定场景没有候选身份校验失败。
    /// </summary>
    /// <param name="label">断言场景名称。</param>
    /// <param name="failures">候选身份校验失败信息集合。</param>
    private static void AssertNoFailures(string label, IReadOnlyCollection<string> failures)
    {
        if (failures.Count > 0)
            throw new InvalidOperationException($"{label}: {string.Join(" | ", failures)}");
    }

    /// <summary>
    /// 将源仓库中的生成资产复制到检出副本。
    /// </summary>
    /// <param name="sourceRepository">源仓库根目录。</param>
    /// <param name="checkout">检出副本目录。</param>
    /// <param name="directoryName">待复制的相对目录名。</param>
    private static void CopyGeneratedFiles(string sourceRepository, string checkout, string directoryName)
    {
        var sourceDirectory = Path.Combine(sourceRepository, directoryName);
        if (!Directory.Exists(sourceDirectory))
            throw new DirectoryNotFoundException(sourceDirectory);

        foreach (var sourcePath in Directory.EnumerateFiles(sourceDirectory, "*", SearchOption.AllDirectories))
        {
            var relativePath = Path.GetRelativePath(sourceRepository, sourcePath);
            var targetPath = Path.Combine(checkout, relativePath);
            var targetDirectory = Path.GetDirectoryName(targetPath);
            if (targetDirectory is not null)
                Directory.CreateDirectory(targetDirectory);
            File.Copy(sourcePath, targetPath, overwrite: true);
        }
    }

    /// <summary>
    /// 创建用于候选身份校验的最小 NuGet 包。
    /// </summary>
    /// <param name="path">生成的包文件路径。</param>
    /// <param name="packageName">NuGet 包文件名。</param>
    /// <param name="assemblyPaths">待写入包内资产的程序集路径。</param>
    /// <param name="repositoryCommit">包元数据中的仓库提交标识。</param>
    /// <param name="repositoryUrl">包元数据中的仓库地址。</param>
    /// <param name="usePlatformXmlFormatting">是否使用平台换行和 XML 属性顺序。</param>
    /// <param name="documentationSummary">程序集 XML 文档中的摘要文本。</param>
    private static void CreatePackage(string path, string packageName,
        IReadOnlyDictionary<string, string> assemblyPaths, string repositoryCommit = "base",
        string repositoryUrl = "https://example.invalid/repository", bool usePlatformXmlFormatting = false,
        string documentationSummary = "package documentation")
    {
        var directory = Path.GetDirectoryName(path);
        if (directory is not null)
            Directory.CreateDirectory(directory);

        using var file = File.Create(path);
        using var archive = new ZipArchive(file, ZipArchiveMode.Create);
        var lineEnding = usePlatformXmlFormatting ? "\n" : "\r\n";
        var repositoryAttributes = usePlatformXmlFormatting
            ? $"url=\"{repositoryUrl}\" type=\"git\""
            : $"type=\"git\" url=\"{repositoryUrl}\"";
        WritePackageText(archive, "README.md", "package readme\r\n");
        WritePackageText(archive, "package.nuspec",
            $"<?xml version=\"1.0\" encoding=\"utf-8\"?>{lineEnding}<package>{lineEnding}  <metadata>{lineEnding}    <repository {repositoryAttributes} branch=\"refs/heads/main\" commit=\"{repositoryCommit}\" />{lineEnding}  </metadata>{lineEnding}</package>{lineEnding}",
            usePlatformXmlFormatting);
        WritePackageText(archive, "_rels/.rels",
            $"<?xml version=\"1.0\" encoding=\"utf-8\"?>{lineEnding}<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">{lineEnding}  <Relationship Target=\"/{packageName}\" Type=\"http://schemas.microsoft.com/packaging/2010/07/manifest\" Id=\"R1\" />{lineEnding}</Relationships>{lineEnding}",
            usePlatformXmlFormatting);
        WritePackageText(archive, "[Content_Types].xml",
            $"<?xml version=\"1.0\" encoding=\"utf-8\"?>{lineEnding}<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">{lineEnding}  <Default ContentType=\"application/xml\" Extension=\"xml\" />{lineEnding}</Types>{lineEnding}",
            usePlatformXmlFormatting);
        WritePackageText(archive, "package/services/metadata/core-properties/nuget.psmdcp",
            $"created={DateTimeOffset.UtcNow:O}\r\n");
        foreach (var asset in GetPackageAssets(packageName, assemblyPaths))
        {
            var entry = archive.CreateEntry("lib/" + asset.Key, CompressionLevel.Optimal);
            using (var input = File.OpenRead(asset.Value))
            using (var output = entry.Open())
                input.CopyTo(output);

            WritePackageText(archive, "lib/" + Path.ChangeExtension(asset.Key, ".xml"),
                $"<?xml version=\"1.0\" encoding=\"utf-8\"?>{lineEnding}<doc>{lineEnding}  <members>{lineEnding}    <member name=\"T:Package\"><summary>{documentationSummary}</summary></member>{lineEnding}  </members>{lineEnding}</doc>{lineEnding}",
                usePlatformXmlFormatting);
        }
    }

    /// <summary>
    /// 替换测试包中的程序集资产或删除该资产。
    /// </summary>
    /// <param name="packagePath">待修改的包文件路径。</param>
    /// <param name="assetPath">包内原程序集资产路径。</param>
    /// <param name="content">替换内容；为 <see langword="null" /> 时删除资产。</param>
    /// <param name="replacementPath">替换资产路径；省略时沿用原路径。</param>
    private static void ReplacePackageDll(string packagePath, string assetPath, byte[]? content,
        string? replacementPath = null)
    {
        using var archive = ZipFile.Open(packagePath, ZipArchiveMode.Update);
        var existing = archive.GetEntry(assetPath)
            ?? throw new InvalidOperationException($"Fixture DLL is missing: {assetPath}");
        existing.Delete();
        if (content is null)
            return;
        var replacement = archive.CreateEntry(replacementPath ?? assetPath);
        using var output = replacement.Open();
        output.Write(content, 0, content.Length);
    }

    /// <summary>
    /// 根据包名返回包内程序集资产映射。
    /// </summary>
    /// <param name="packageName">NuGet 包文件名。</param>
    /// <param name="assemblyPaths">候选程序集路径。</param>
    /// <returns>包内资产路径到程序集路径的映射。</returns>
    private static IReadOnlyDictionary<string, string> GetPackageAssets(string packageName,
        IReadOnlyDictionary<string, string> assemblyPaths)
    {
        if (packageName.StartsWith("Bing.Offices.Abstractions.", StringComparison.Ordinal))
            return new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["netstandard2.0/Bing.Offices.Abstractions.dll"] =
                    assemblyPaths["netstandard2.0/Bing.Offices.Abstractions.dll"]
            };
        if (packageName.StartsWith("Bing.Offices.Core.", StringComparison.Ordinal))
            return new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["netstandard2.0/Bing.Offices.Core.dll"] =
                    assemblyPaths["netstandard2.0/Bing.Offices.Core.dll"]
            };
        if (packageName.StartsWith("Bing.Offices.MiniExcel.", StringComparison.Ordinal))
            return new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["net6.0/Bing.Offices.MiniExcel.dll"] =
                    assemblyPaths["net6.0/Bing.Offices.MiniExcel.dll"],
                ["net8.0/Bing.Offices.MiniExcel.dll"] =
                    assemblyPaths["net8.0/Bing.Offices.MiniExcel.dll"]
            };
        if (packageName.StartsWith("Bing.Offices.ClosedXml.", StringComparison.Ordinal))
            return new Dictionary<string, string>(StringComparer.Ordinal)
            {
                ["net6.0/Bing.Offices.ClosedXml.dll"] =
                    assemblyPaths["net6.0/Bing.Offices.ClosedXml.dll"],
                ["net8.0/Bing.Offices.ClosedXml.dll"] =
                    assemblyPaths["net8.0/Bing.Offices.ClosedXml.dll"]
            };

        return new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["net6.0/Bing.Offices.Npoi.dll"] = assemblyPaths["net6.0/Bing.Offices.Npoi.dll"],
            ["net8.0/Bing.Offices.Npoi.dll"] = assemblyPaths["net8.0/Bing.Offices.Npoi.dll"]
        };
    }

    /// <summary>
    /// 向测试包写入 UTF-8 文本资产。
    /// </summary>
    /// <param name="archive">目标 ZIP 包。</param>
    /// <param name="path">包内资产路径。</param>
    /// <param name="value">待写入的文本内容。</param>
    /// <param name="emitUtf8Identifier">是否写入 UTF-8 BOM。</param>
    private static void WritePackageText(ZipArchive archive, string path, string value,
        bool emitUtf8Identifier = false)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(emitUtf8Identifier));
        writer.Write(value);
    }

    /// <summary>
    /// 创建目标目录并复制测试文件。
    /// </summary>
    /// <param name="sourcePath">源文件路径。</param>
    /// <param name="targetPath">目标文件路径。</param>
    private static void CopyFile(string sourcePath, string targetPath)
    {
        var directory = Path.GetDirectoryName(targetPath);
        if (directory is not null)
            Directory.CreateDirectory(directory);
        File.Copy(sourcePath, targetPath, overwrite: true);
    }

    /// <summary>
    /// 将文本编码为不带 BOM 的 UTF-8 字节序列。
    /// </summary>
    /// <param name="value">待编码的文本。</param>
    /// <returns>UTF-8 编码后的字节序列。</returns>
    private static byte[] Utf8(string value) => Encoding.UTF8.GetBytes(value);

    /// <summary>
    /// 以不带 BOM 的 UTF-8 写入测试文本文件。
    /// </summary>
    /// <param name="path">目标文件路径。</param>
    /// <param name="value">待写入的文本内容。</param>
    private static void WriteText(string path, string value)
    {
        var directory = Path.GetDirectoryName(path);
        if (directory is not null)
            Directory.CreateDirectory(directory);
        File.WriteAllText(path, value, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }

    /// <summary>
    /// 写入测试字节文件并创建目标目录。
    /// </summary>
    /// <param name="path">目标文件路径。</param>
    /// <param name="value">待写入的字节内容。</param>
    private static void WriteBytes(string path, byte[] value)
    {
        var directory = Path.GetDirectoryName(path);
        if (directory is not null)
            Directory.CreateDirectory(directory);
        File.WriteAllBytes(path, value);
    }

    /// <summary>
    /// 删除临时目录及其全部内容。
    /// </summary>
    /// <param name="path">待删除的目录路径。</param>
    private static void DeleteTree(string path)
    {
        if (!Directory.Exists(path))
            return;

        foreach (var file in Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories))
            File.SetAttributes(file, FileAttributes.Normal);
        Directory.Delete(path, recursive: true);
    }

    /// <summary>
    /// 在指定仓库中运行 Git 命令并返回标准输出。
    /// </summary>
    /// <param name="repository">Git 命令的工作目录。</param>
    /// <param name="arguments">Git 命令参数。</param>
    /// <returns>Git 命令的标准输出文本。</returns>
    private static string RunGit(string repository, params string[] arguments)
    {
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = "git",
                WorkingDirectory = repository,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardOutputEncoding = Encoding.UTF8,
                StandardErrorEncoding = Encoding.UTF8
            }
        };
        foreach (var argument in arguments)
            process.StartInfo.ArgumentList.Add(argument);

        if (!process.Start())
            throw new InvalidOperationException("git process did not start.");

        var output = process.StandardOutput.ReadToEnd();
        var error = process.StandardError.ReadToEnd();
        process.WaitForExit();
        if (process.ExitCode != 0)
            throw new InvalidOperationException(
                $"git {string.Join(" ", arguments)} failed ({process.ExitCode}): {error.Trim()}");
        return output;
    }

    /// <summary>
    /// 干净检出副本中的构建资产路径。
    /// </summary>
    /// <param name="Assemblies">检出副本中的程序集路径映射。</param>
    /// <param name="PackageRoot">检出副本中的 NuGet 包目录。</param>
    private sealed record CheckoutPaths(
        IReadOnlyDictionary<string, string> Assemblies,
        string PackageRoot);
}
