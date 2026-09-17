using System.Diagnostics;
using System.Text;
using System.IO.Compression;
using Bing.Offices.ApiSnapshot;

internal static class CandidateIdentityContractTest
{
    private const string GeneratorVersion = "2.7.0";

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
                    Path.Combine(assemblyRoot, "net8.0", "Bing.Offices.MiniExcel.dll")
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
                "Bing.Offices.MiniExcel.2.0.0.nupkg"
            };
            foreach (var packageName in packageNames)
                CreatePackage(Path.Combine(packageRoot, packageName), packageName, assemblyPaths);

            var approvalPath = Path.Combine(repository,
                CandidateIdentityVerifier.BreakingApprovalPath.Replace('/', Path.DirectorySeparatorChar));
            WriteText(approvalPath, "approved\r\n");

            var ignoredArtifactPath = Path.Combine(repository, "ai_docs", "tasks",
                "BO-RC-20260908-002", "artifacts", "ignored.json");
            var ignoredLockfilePath = Path.Combine(repository, "tests", "packages.lock.json");

            var captureFailures = new List<string>();
            var identity = CandidateIdentityVerifier.Capture(
                assemblyPaths, packageRoot, repository, baseCommit, captureFailures);
            AssertNoFailures("dirty capture", captureFailures);

            var sourceManifestHash = identity.CandidateSourceManifestSha256;
            WriteBytes(ignoredArtifactPath, Utf8("ignored artifact\r\n"));
            WriteBytes(ignoredLockfilePath, Utf8("ignored lockfile\r\n"));
            var ignoredCaptureFailures = new List<string>();
            var ignoredCapture = CandidateIdentityVerifier.Capture(
                assemblyPaths, packageRoot, repository, baseCommit, ignoredCaptureFailures);
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
                assemblyPaths, packageRoot, repository, baseCommit, releaseUnrelatedCaptureFailures);
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
                CandidateIdentityVerifier.BreakingApprovalPath.Replace('/', Path.DirectorySeparatorChar));
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
                CandidateIdentityVerifier.BreakingApprovalPath.Replace('/', Path.DirectorySeparatorChar));
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

    private static List<string> Validate(
        ApiBaselineDocument baseline,
        IReadOnlyDictionary<string, string> assemblyPaths,
        string packageRoot,
        string repository)
    {
        var failures = new List<string>();
        CandidateIdentityVerifier.Validate(
            baseline, assemblyPaths, packageRoot, repository, failures);
        return failures;
    }

    private static void AssertValidationFails(
        string label,
        ApiBaselineDocument baseline,
        IReadOnlyDictionary<string, string> assemblyPaths,
        string packageRoot,
        string repository)
    {
        var failures = Validate(baseline, assemblyPaths, packageRoot, repository);
        if (failures.Count == 0)
            throw new InvalidOperationException($"{label} did not fail validation.");
    }

    private static void AssertNoFailures(string label, IReadOnlyCollection<string> failures)
    {
        if (failures.Count > 0)
            throw new InvalidOperationException($"{label}: {string.Join(" | ", failures)}");
    }

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

        return new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["net6.0/Bing.Offices.Npoi.dll"] = assemblyPaths["net6.0/Bing.Offices.Npoi.dll"],
            ["net8.0/Bing.Offices.Npoi.dll"] = assemblyPaths["net8.0/Bing.Offices.Npoi.dll"]
        };
    }

    private static void WritePackageText(ZipArchive archive, string path, string value,
        bool emitUtf8Identifier = false)
    {
        var entry = archive.CreateEntry(path, CompressionLevel.Optimal);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(emitUtf8Identifier));
        writer.Write(value);
    }

    private static void CopyFile(string sourcePath, string targetPath)
    {
        var directory = Path.GetDirectoryName(targetPath);
        if (directory is not null)
            Directory.CreateDirectory(directory);
        File.Copy(sourcePath, targetPath, overwrite: true);
    }

    private static byte[] Utf8(string value) => Encoding.UTF8.GetBytes(value);

    private static void WriteText(string path, string value)
    {
        var directory = Path.GetDirectoryName(path);
        if (directory is not null)
            Directory.CreateDirectory(directory);
        File.WriteAllText(path, value, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
    }

    private static void WriteBytes(string path, byte[] value)
    {
        var directory = Path.GetDirectoryName(path);
        if (directory is not null)
            Directory.CreateDirectory(directory);
        File.WriteAllBytes(path, value);
    }

    private static void DeleteTree(string path)
    {
        if (!Directory.Exists(path))
            return;

        foreach (var file in Directory.EnumerateFiles(path, "*", SearchOption.AllDirectories))
            File.SetAttributes(file, FileAttributes.Normal);
        Directory.Delete(path, recursive: true);
    }

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

    private sealed record CheckoutPaths(
        IReadOnlyDictionary<string, string> Assemblies,
        string PackageRoot);
}
