using System.Diagnostics;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml;
using System.Xml.Linq;
using Bing.Offices.ApiSnapshot;

const string schema = "bing.offices.public-api.v2";
const string generatorVersion = "2.7.0";

var arguments = ParseArguments(args);
var root = Path.GetFullPath(arguments.GetValueOrDefault("root") ?? "output/release");
var baselinePath = Path.GetFullPath(arguments.GetValueOrDefault("baseline") ?? "build/api-snapshot-baseline.json");
var output = arguments.GetValueOrDefault("output");
var dependencies = arguments.GetValueOrDefault("dependencies");
var packages = arguments.GetValueOrDefault("packages");
var repository = Path.GetFullPath(arguments.GetValueOrDefault("repository") ?? Directory.GetCurrentDirectory());
var captureOnly = string.Equals(arguments.GetValueOrDefault("capture"), "true", StringComparison.OrdinalIgnoreCase);
var targetFrameworks = new[] { "net6.0", "net8.0" };
var jsonOptions = new JsonSerializerOptions { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase };

if (string.Equals(arguments.GetValueOrDefault("identity-self-test"), "true", StringComparison.OrdinalIgnoreCase))
    return CandidateIdentityContractTest.Run();

ApiBaselineDocument? baseline = null;
if (!captureOnly)
{
    if (!File.Exists(baselinePath))
    {
        Console.Error.WriteLine($"API baseline is missing: {baselinePath}");
        return 2;
    }
    baseline = JsonSerializer.Deserialize<ApiBaselineDocument>(
        File.ReadAllText(baselinePath, System.Text.Encoding.UTF8), jsonOptions)
        ?? throw new InvalidOperationException("API snapshot baseline is empty.");
}

var candidate = new ApiBaselineDocument
{
    Schema = schema,
    GeneratorVersion = generatorVersion,
    BaselineCommit = arguments.GetValueOrDefault("commit") ?? "",
    ApprovedBy = "",
    ApprovedAt = ""
};
var failures = new List<string>();
var diffs = new Dictionary<string, Dictionary<string, ApiMemberDiff>>(StringComparer.Ordinal);
var candidateAssemblyPaths = GetCandidateAssemblyPaths(root);

if (!captureOnly)
{
    CandidateIdentityVerifier.Validate(baseline!, candidateAssemblyPaths, packages, repository, failures);
}

foreach (var tfm in targetFrameworks)
{
    var paths = new[]
    {
        Path.Combine(root, "netstandard2.0", "Bing.Offices.Abstractions.dll"),
        Path.Combine(root, "netstandard2.0", "Bing.Offices.Core.dll"),
        Path.Combine(root, tfm, "Bing.Offices.Npoi.dll"),
        Path.Combine(root, tfm, "Bing.Offices.MiniExcel.dll")
    };
    if (paths.Any(path => !File.Exists(path)))
    {
        failures.Add($"{tfm}: one or more target assemblies are missing");
        continue;
    }

    var additionalPaths = paths.Select(Path.GetDirectoryName).Where(path => path is not null)
        .SelectMany(path => Directory.EnumerateFiles(path!, "*.dll")).ToList();
    if (!string.IsNullOrWhiteSpace(dependencies))
    {
        var dependencyDirectory = Path.GetFullPath(dependencies);
        if (!Directory.Exists(dependencyDirectory))
            throw new DirectoryNotFoundException($"Dependency directory does not exist: {dependencyDirectory}");
        additionalPaths.AddRange(Directory.EnumerateFiles(dependencyDirectory, "*.dll",
            SearchOption.AllDirectories));
    }

    var snapshots = paths.Select(path => PublicApiSnapshot.Load(path, additionalPaths)).ToDictionary(
        snapshot => snapshot.AssemblyName,
        snapshot => new ApiSnapshotRecord
        {
            MemberCount = snapshot.MemberCount,
            Hash = snapshot.Hash,
            Lines = snapshot.Lines.ToList()
        }, StringComparer.Ordinal);
    candidate.Assemblies[tfm] = snapshots;

    if (captureOnly)
        continue;

    var approvedBaseline = baseline!;
    if (approvedBaseline.Schema != schema || approvedBaseline.GeneratorVersion != generatorVersion)
        failures.Add($"{tfm}: baseline schema/generator mismatch");
    if (string.IsNullOrWhiteSpace(approvedBaseline.ApprovedBy) || string.IsNullOrWhiteSpace(approvedBaseline.ApprovedAt))
        failures.Add($"{tfm}: baseline approval metadata is missing");
    if (!approvedBaseline.Assemblies.TryGetValue(tfm, out var expectedAssemblies))
    {
        failures.Add($"{tfm}: missing baseline");
        continue;
    }

    var tfmDiffs = new Dictionary<string, ApiMemberDiff>(StringComparer.Ordinal);
    foreach (var pair in snapshots)
    {
        if (!expectedAssemblies.TryGetValue(pair.Key, out var expected))
        {
            failures.Add($"{tfm}/{pair.Key}: missing baseline assembly");
            continue;
        }
        if (string.Equals(pair.Value.Hash, expected.Hash, StringComparison.Ordinal))
            continue;
        var expectedLines = new HashSet<string>(expected.Lines ?? new List<string>(), StringComparer.Ordinal);
        var actualLines = new HashSet<string>(pair.Value.Lines ?? new List<string>(), StringComparer.Ordinal);
        tfmDiffs[pair.Key] = new ApiMemberDiff
        {
            Added = actualLines.Except(expectedLines, StringComparer.Ordinal).OrderBy(line => line, StringComparer.Ordinal).ToList(),
            Removed = expectedLines.Except(actualLines, StringComparer.Ordinal).OrderBy(line => line, StringComparer.Ordinal).ToList()
        };
        failures.Add($"{tfm}/{pair.Key}: expected={expected.Hash}; actual={pair.Value.Hash}");
    }
    diffs[tfm] = tfmDiffs;
}

if (captureOnly)
    candidate.CandidateIdentity = CandidateIdentityVerifier.Capture(
        candidateAssemblyPaths, packages, repository, candidate.BaselineCommit, failures);

if (output is not null)
{
    var directory = Path.GetFullPath(output);
    Directory.CreateDirectory(directory);
    foreach (var pair in candidate.Assemblies)
    {
        File.WriteAllText(Path.Combine(directory, $"api-snapshot-{pair.Key}.json"),
            JsonSerializer.Serialize(pair.Value, jsonOptions), System.Text.Encoding.UTF8);
    }
    File.WriteAllText(Path.Combine(directory, "api-candidate.json"),
        JsonSerializer.Serialize(candidate, jsonOptions), System.Text.Encoding.UTF8);
    File.WriteAllText(Path.Combine(directory, "api-diff.json"),
        JsonSerializer.Serialize(diffs, jsonOptions), System.Text.Encoding.UTF8);
}

if (failures.Count > 0)
{
    Console.Error.WriteLine(string.Join(Environment.NewLine, failures));
    return 1;
}

Console.WriteLine(captureOnly
    ? "API snapshot capture completed for net6.0 and net8.0."
    : "API snapshot comparison passed for net6.0 and net8.0.");
return 0;

static Dictionary<string, string> ParseArguments(string[] args)
{
    var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    for (var index = 0; index < args.Length; index++)
    {
        if (!args[index].StartsWith("--", StringComparison.Ordinal) || index + 1 >= args.Length)
            throw new ArgumentException($"Invalid argument: {args[index]}");
        result[args[index][2..]] = args[++index];
    }
    return result;
}

static Dictionary<string, string> GetCandidateAssemblyPaths(string root) =>
    new(StringComparer.Ordinal)
    {
        ["netstandard2.0/Bing.Offices.Abstractions.dll"] =
            Path.Combine(root, "netstandard2.0", "Bing.Offices.Abstractions.dll"),
        ["netstandard2.0/Bing.Offices.Core.dll"] =
            Path.Combine(root, "netstandard2.0", "Bing.Offices.Core.dll"),
        ["net6.0/Bing.Offices.Npoi.dll"] =
            Path.Combine(root, "net6.0", "Bing.Offices.Npoi.dll"),
        ["net8.0/Bing.Offices.Npoi.dll"] =
            Path.Combine(root, "net8.0", "Bing.Offices.Npoi.dll"),
        ["net6.0/Bing.Offices.MiniExcel.dll"] =
            Path.Combine(root, "net6.0", "Bing.Offices.MiniExcel.dll"),
        ["net8.0/Bing.Offices.MiniExcel.dll"] =
            Path.Combine(root, "net8.0", "Bing.Offices.MiniExcel.dll")
    };

public sealed class ApiBaselineDocument
{
    public string Schema { get; set; } = "";
    public string GeneratorVersion { get; set; } = "";
    public string BaselineCommit { get; set; } = "";
    public string ApprovedBy { get; set; } = "";
    public string ApprovedAt { get; set; } = "";
    public ApiCandidateIdentity? CandidateIdentity { get; set; }
    public Dictionary<string, Dictionary<string, ApiSnapshotRecord>> Assemblies { get; set; } = new(StringComparer.Ordinal);
}

public sealed class ApiCandidateIdentity
{
    public string ArtifactIdentityFormat { get; set; } = "";
    public string BaseCommit { get; set; } = "";
    public string WorktreeState { get; set; } = "";
    public string CandidateSourceManifestSha256 { get; set; } = "";
    public string BreakingApprovalArtifact { get; set; } = "";
    public string BreakingApprovalSha256 { get; set; } = "";
    public Dictionary<string, string> AssemblyFiles { get; set; } = new(StringComparer.Ordinal);
    public Dictionary<string, string> NupkgFiles { get; set; } = new(StringComparer.Ordinal);
}

public sealed class ApiSnapshotRecord
{
    public int MemberCount { get; set; }
    public string Hash { get; set; } = "";
    public List<string> Lines { get; set; } = new();
}

public sealed class ApiMemberDiff
{
    public List<string> Added { get; set; } = new();
    public List<string> Removed { get; set; } = new();
}

internal static class CandidateIdentityVerifier
{
    private const string ArtifactIdentityFormat = "logical-v3";
    internal const string BreakingApprovalPath =
        "ai_docs/tasks/BO-RC-20260908-002/api-breaking-approval.md";

    private static readonly string[] RequiredPackageIds =
    {
        "Bing.Offices.Abstractions",
        "Bing.Offices.Core",
        "Bing.Offices.Npoi",
        "Bing.Offices.MiniExcel"
    };

    private static readonly string[] CandidateSourceScope =
    {
        // 只绑定会影响 Release 程序集或 nupkg 内容的输入，避免文档、测试和 CI 配置导致 API 基线失效。
        "src", "asset", "build/ApiSnapshot", "README.md", "LICENSE", ".gitattributes",
        "framework.props", "common.props", "Directory.Build.targets", "version.props", "version.dev.props"
    };

    public static ApiCandidateIdentity Capture(
        IReadOnlyDictionary<string, string> assemblyPaths,
        string? packagesRoot,
        string repositoryRoot,
        string candidateCommit,
        List<string> failures)
    {
        var identity = new ApiCandidateIdentity
        {
            BaseCommit = candidateCommit,
            ArtifactIdentityFormat = ArtifactIdentityFormat
        };

        identity.AssemblyFiles = CaptureAssemblyFiles(assemblyPaths, "capture", failures);

        if (string.IsNullOrWhiteSpace(identity.BaseCommit)
            && TryRunGit(repositoryRoot, new[] { "rev-parse", "HEAD" }, out var head, out _))
            identity.BaseCommit = head.Trim();

        if (TryRunGit(repositoryRoot, new[] { "status", "--porcelain=v1", "--untracked-files=all" },
                out var status, out var statusError))
            identity.WorktreeState = string.IsNullOrEmpty(status.Trim()) ? "clean" : "dirty";
        else
            failures.Add($"capture: unable to read git status: {statusError}");

        if (TryGetCandidateSourceManifestSha256(repositoryRoot, out var sourceHash, out var sourceError))
            identity.CandidateSourceManifestSha256 = sourceHash;
        else
            failures.Add($"capture: unable to hash candidate source manifest: {sourceError}");

        identity.BreakingApprovalArtifact = BreakingApprovalPath;
        var approvalPath = ResolveWithinRoot(repositoryRoot, BreakingApprovalPath);
        if (approvalPath is null)
        {
            failures.Add($"capture: approval file path escapes repository root: {BreakingApprovalPath}");
        }
        else if (!File.Exists(approvalPath))
        {
            failures.Add($"capture: approval file is missing: {approvalPath}");
        }
        else
        {
            try
            {
                identity.BreakingApprovalSha256 = ApiSnapshotFileHash.ComputeCanonicalTextSha256(approvalPath);
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
            {
                failures.Add($"capture: approval file could not be hashed: {approvalPath}; {exception.Message}");
            }
        }

        if (!string.IsNullOrWhiteSpace(packagesRoot))
            identity.NupkgFiles = CaptureNupkgFiles(packagesRoot, identity.AssemblyFiles, assemblyPaths, failures);

        return identity;
    }

    public static void Validate(
        ApiBaselineDocument baseline,
        IReadOnlyDictionary<string, string> assemblyPaths,
        string? packagesRoot,
        string repositoryRoot,
        List<string> failures)
    {
        var identity = baseline.CandidateIdentity;
        if (identity is null)
        {
            failures.Add("baseline candidate identity metadata is missing");
            return;
        }

        if (!string.Equals(identity.ArtifactIdentityFormat, ArtifactIdentityFormat,
                StringComparison.Ordinal))
            failures.Add($"candidate identity artifactIdentityFormat is invalid: {identity.ArtifactIdentityFormat}");

        if (string.IsNullOrWhiteSpace(identity.BaseCommit)
            || !string.Equals(identity.BaseCommit, baseline.BaselineCommit, StringComparison.Ordinal))
            failures.Add("candidate identity baseCommit does not match baselineCommit");

        if (!string.Equals(identity.WorktreeState, "clean", StringComparison.OrdinalIgnoreCase)
            && !string.Equals(identity.WorktreeState, "dirty", StringComparison.OrdinalIgnoreCase))
            failures.Add($"candidate identity worktreeState is invalid: {identity.WorktreeState}");

        if (TryGetCandidateSourceManifestSha256(repositoryRoot, out var actualSourceHash, out var sourceError))
            ValidateHash("candidate source manifest", identity.CandidateSourceManifestSha256, actualSourceHash, failures);
        else
            failures.Add($"candidate identity source manifest verification failed: {sourceError}");

        var actualAssemblyFiles = ValidateAssemblyFiles(identity.AssemblyFiles, assemblyPaths, failures);
        ValidateNupkgFiles(identity.NupkgFiles, packagesRoot, actualAssemblyFiles, assemblyPaths, failures);
        ValidateApprovalFile(identity, repositoryRoot, failures);
    }

    private static Dictionary<string, string> ValidateAssemblyFiles(
        IReadOnlyDictionary<string, string>? recordedFiles,
        IReadOnlyDictionary<string, string> assemblyPaths,
        List<string> failures)
    {
        var expectedPaths = assemblyPaths.ToDictionary(
            pair => NormalizeRelativePath(pair.Key), pair => pair.Value, StringComparer.Ordinal);
        var recorded = NormalizeHashMap(recordedFiles, "assembly", failures);

        var actual = CaptureAssemblyFiles(expectedPaths, "validation", failures);
        foreach (var pair in expectedPaths)
        {
            if (!recorded.TryGetValue(pair.Key, out var expectedHash))
            {
                failures.Add($"candidate identity assembly hash is missing: {pair.Key}");
                continue;
            }

            if (actual.TryGetValue(pair.Key, out var actualHash))
                ValidateHash($"assembly {pair.Key}", expectedHash, actualHash, failures);
        }

        foreach (var path in recorded.Keys.Except(expectedPaths.Keys, StringComparer.Ordinal))
            failures.Add($"candidate identity contains an unexpected assembly path: {path}");

        return actual;
    }

    private static void ValidateNupkgFiles(
        IReadOnlyDictionary<string, string>? recordedFiles,
        string? packagesRoot,
        IReadOnlyDictionary<string, string> assemblyFiles,
        IReadOnlyDictionary<string, string> assemblyPaths,
        List<string> failures)
    {
        var recorded = NormalizeHashMap(recordedFiles, "nupkg", failures);
        if (recorded.Count < RequiredPackageIds.Length)
            failures.Add("candidate identity must contain hashes for the four production nupkg files");
        if (string.IsNullOrWhiteSpace(packagesRoot))
        {
            failures.Add("nupkg hash verification requires --packages <directory>");
            return;
        }

        var fullPackagesRoot = Path.GetFullPath(packagesRoot);
        if (!Directory.Exists(fullPackagesRoot))
        {
            failures.Add($"nupkg directory does not exist: {fullPackagesRoot}");
            return;
        }

        foreach (var pair in recorded)
        {
            if (!string.Equals(Path.GetExtension(pair.Key), ".nupkg", StringComparison.OrdinalIgnoreCase))
            {
                failures.Add($"candidate identity nupkg path is not a .nupkg file: {pair.Key}");
                continue;
            }

            var packagePath = ResolveWithinRoot(fullPackagesRoot, pair.Key);
            if (packagePath is null)
            {
                failures.Add($"candidate identity nupkg path escapes package root: {pair.Key}");
                continue;
            }

            if (!ApiSnapshotFileHash.IsSha256(pair.Value))
                continue;
            if (!File.Exists(packagePath))
            {
                failures.Add($"nupkg {pair.Key} is missing: {packagePath}");
                continue;
            }

            try
            {
                var actualHash = ComputePackageIdentityHash(packagePath, assemblyFiles, assemblyPaths);
                ValidateHash($"nupkg {pair.Key}", pair.Value, actualHash, failures);
            }
            catch (Exception exception) when (exception is IOException or InvalidDataException or XmlException
                or UnauthorizedAccessException or BadImageFormatException or FileLoadException or InvalidOperationException)
            {
                failures.Add($"nupkg {pair.Key} could not be canonicalized: {packagePath}; {exception.Message}");
            }
        }

        foreach (var packageId in RequiredPackageIds)
        {
            var matches = recorded.Keys.Where(path => string.Equals(GetPackageId(path), packageId,
                StringComparison.OrdinalIgnoreCase)).ToArray();
            if (matches.Length != 1)
            {
                failures.Add($"candidate identity must contain exactly one nupkg hash for {packageId}");
                continue;
            }

            var actualPackages = Directory.EnumerateFiles(fullPackagesRoot, "*.nupkg", SearchOption.TopDirectoryOnly)
                .Where(path => string.Equals(GetPackageId(Path.GetFileName(path)), packageId,
                    StringComparison.OrdinalIgnoreCase))
                .Select(path => NormalizeRelativePath(Path.GetRelativePath(fullPackagesRoot, path)))
                .ToArray();
            if (actualPackages.Length != 1 || !string.Equals(actualPackages[0], matches[0], StringComparison.Ordinal))
                failures.Add($"candidate identity does not identify the unique final nupkg for {packageId}");
        }
    }

    private static void ValidateApprovalFile(ApiCandidateIdentity identity, string repositoryRoot,
        List<string> failures)
    {
        var path = NormalizeRelativePath(identity.BreakingApprovalArtifact);
        if (!string.Equals(path, BreakingApprovalPath, StringComparison.Ordinal))
        {
            failures.Add($"candidate identity approval path is not the approved task-root file: {path}");
            return;
        }

        var fullPath = ResolveWithinRoot(repositoryRoot, path);
        if (fullPath is null)
        {
            failures.Add($"candidate identity approval path escapes repository root: {path}");
            return;
        }

        ValidateCanonicalTextFileHash(fullPath, identity.BreakingApprovalSha256,
            $"approval file {path}", failures);
    }

    private static Dictionary<string, string> CaptureNupkgFiles(string packagesRoot,
        IReadOnlyDictionary<string, string> assemblyFiles,
        IReadOnlyDictionary<string, string> assemblyPaths, List<string> failures)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        var fullRoot = Path.GetFullPath(packagesRoot);
        if (!Directory.Exists(fullRoot))
        {
            failures.Add($"capture: nupkg directory does not exist: {fullRoot}");
            return result;
        }

        foreach (var packageId in RequiredPackageIds)
        {
            var matches = Directory.EnumerateFiles(fullRoot, "*.nupkg", SearchOption.TopDirectoryOnly)
                .Where(path => Path.GetFileName(path).StartsWith(packageId + ".", StringComparison.OrdinalIgnoreCase))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (matches.Length != 1)
            {
                failures.Add($"capture: expected exactly one final nupkg for {packageId}, found {matches.Length}");
                continue;
            }

            var relativePath = NormalizeRelativePath(Path.GetRelativePath(fullRoot, matches[0]));
            try
            {
                result[relativePath] = ComputePackageIdentityHash(matches[0], assemblyFiles, assemblyPaths);
            }
            catch (Exception exception) when (exception is IOException or InvalidDataException or XmlException
                or UnauthorizedAccessException or BadImageFormatException or FileLoadException or InvalidOperationException)
            {
                failures.Add($"capture: nupkg could not be canonicalized: {matches[0]}; {exception.Message}");
            }
        }

        return result;
    }

    private static Dictionary<string, string> CaptureAssemblyFiles(
        IReadOnlyDictionary<string, string> assemblyPaths, string operation, List<string> failures)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        var additionalAssemblyPaths = assemblyPaths.Values
            .Select(Path.GetDirectoryName)
            .Where(path => !string.IsNullOrWhiteSpace(path) && Directory.Exists(path!))
            .SelectMany(path => Directory.EnumerateFiles(path!, "*.dll"))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray();

        foreach (var pair in assemblyPaths)
        {
            var relativePath = NormalizeRelativePath(pair.Key);
            if (!File.Exists(pair.Value))
            {
                failures.Add($"{operation}: candidate assembly is missing: {pair.Value}");
                continue;
            }

            try
            {
                result[relativePath] = ComputeAssemblyIdentityHash(pair.Value, additionalAssemblyPaths);
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException
                or BadImageFormatException or FileLoadException or InvalidOperationException)
            {
                failures.Add($"{operation}: candidate assembly could not be canonicalized: {pair.Value}; {exception.Message}");
            }
        }

        return result;
    }

    private static string ComputeAssemblyIdentityHash(string path, IEnumerable<string> additionalAssemblyPaths)
    {
        var snapshot = PublicApiSnapshot.Load(path, additionalAssemblyPaths);
        return ComputeUtf8Sha256(string.Join("\n", new[]
        {
            "bing.offices.assembly-identity.v1",
            snapshot.AssemblyName,
            snapshot.MemberCount.ToString(System.Globalization.CultureInfo.InvariantCulture),
            snapshot.Hash
        }));
    }

    private static string ComputePackageIdentityHash(string packagePath,
        IReadOnlyDictionary<string, string> assemblyFiles,
        IReadOnlyDictionary<string, string> assemblyPaths)
    {
        using var archive = ZipFile.OpenRead(packagePath);
        var entries = new List<string>();
        foreach (var entry in archive.Entries
                     .Where(entry => !string.IsNullOrEmpty(entry.Name))
                     .OrderBy(entry => entry.FullName, StringComparer.Ordinal))
        {
            var relativePath = NormalizeRelativePath(entry.FullName);
            if (string.Equals(relativePath, "package/services/metadata/core-properties/nuget.psmdcp",
                    StringComparison.OrdinalIgnoreCase))
                continue;

            entries.Add($"{relativePath}|{ComputePackageEntryIdentityHash(entry, relativePath, assemblyFiles, assemblyPaths)}");
        }

        return ComputeUtf8Sha256("bing.offices.nupkg-identity.v1\n" + string.Join("\n", entries));
    }

    private static string ComputePackageEntryIdentityHash(ZipArchiveEntry entry, string relativePath,
        IReadOnlyDictionary<string, string> assemblyFiles,
        IReadOnlyDictionary<string, string> assemblyPaths)
    {
        if (relativePath.StartsWith("lib/", StringComparison.OrdinalIgnoreCase)
            && relativePath.EndsWith(".dll", StringComparison.OrdinalIgnoreCase))
        {
            var assemblyPath = relativePath["lib/".Length..];
            if (!assemblyFiles.TryGetValue(assemblyPath, out var assemblyHash))
                throw new InvalidOperationException(
                    $"package assembly does not match a required Release assembly: {relativePath}");
            // 包内 DLL 必须独立读取和校验，不能只凭资产路径复用 Release 身份。
            var temporaryDirectory = Path.Combine(Path.GetTempPath(), $"bing-offices-package-dll-{Guid.NewGuid():N}");
            Directory.CreateDirectory(temporaryDirectory);
            try
            {
                var extractedPath = Path.Combine(temporaryDirectory, Path.GetFileName(assemblyPath));
                using (var input = entry.Open())
                using (var output = File.Create(extractedPath))
                    input.CopyTo(output);
                var dependencies = assemblyPaths.Values
                    .Select(Path.GetDirectoryName)
                    .Where(directory => directory is not null && Directory.Exists(directory))
                    .SelectMany(directory => Directory.EnumerateFiles(directory!, "*.dll"));
                var actualHash = ComputeAssemblyIdentityHash(extractedPath, dependencies);
                if (!string.Equals(actualHash, assemblyHash, StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        $"package assembly identity does not match the required Release assembly: {relativePath}");
                return "assembly:" + actualHash;
            }
            finally
            {
                Directory.Delete(temporaryDirectory, recursive: true);
            }
        }

        using var stream = entry.Open();
        if (IsXmlPackageEntry(relativePath))
        {
            var isNuspec = string.Equals(Path.GetExtension(relativePath), ".nuspec",
                StringComparison.OrdinalIgnoreCase);
            return "xml:" + ComputeXmlPackageEntryIdentityHash(stream, isNuspec);
        }

        return IsCanonicalTextPackageEntry(relativePath)
            ? "text:" + ApiSnapshotFileHash.ComputeCanonicalTextSha256(stream)
            : "binary:" + ApiSnapshotFileHash.ComputeSha256(stream);
    }

    private static string ComputeXmlPackageEntryIdentityHash(Stream stream, bool normalizeRepositoryMetadata)
    {
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            IgnoreWhitespace = true,
            XmlResolver = null
        };
        using var reader = XmlReader.Create(stream, settings);
        var document = XDocument.Load(reader, LoadOptions.None);

        if (normalizeRepositoryMetadata)
        {
            // NuGet 根据当前 checkout 注入这两个源控元数据，候选源清单已独立绑定实际源码。
            foreach (var repository in document.Descendants()
                         .Where(element => string.Equals(element.Name.LocalName, "repository", StringComparison.Ordinal)))
            {
                repository.Attribute("branch")?.Remove();
                repository.Attribute("commit")?.Remove();
            }
        }

        // XML 属性顺序、BOM、XML 声明和纯格式化空白不应使跨 OS 的同一包失效。
        foreach (var element in document.Descendants())
        {
            var attributes = element.Attributes()
                .OrderBy(attribute => attribute.Name.NamespaceName, StringComparer.Ordinal)
                .ThenBy(attribute => attribute.Name.LocalName, StringComparer.Ordinal)
                .Select(attribute => new XAttribute(attribute))
                .ToArray();
            element.ReplaceAttributes(attributes);
        }

        return ComputeUtf8Sha256(document.ToString(SaveOptions.DisableFormatting));
    }

    private static bool IsXmlPackageEntry(string relativePath) =>
        Path.GetExtension(relativePath).ToLowerInvariant() is ".xml" or ".nuspec" or ".rels";

    private static bool IsCanonicalTextPackageEntry(string relativePath)
    {
        var fileName = Path.GetFileName(relativePath);
        if (string.Equals(fileName, "LICENSE", StringComparison.OrdinalIgnoreCase)
            || fileName.StartsWith("README", StringComparison.OrdinalIgnoreCase))
            return true;

        return Path.GetExtension(relativePath).ToLowerInvariant() is ".md" or ".txt" or ".json"
            or ".props" or ".targets" or ".config";
    }

    private static Dictionary<string, string> NormalizeHashMap(
        IReadOnlyDictionary<string, string>? values, string label, List<string> failures)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        if (values is null)
            return result;

        foreach (var pair in values)
        {
            var path = NormalizeRelativePath(pair.Key);
            if (!result.TryAdd(path, pair.Value ?? string.Empty))
                failures.Add($"candidate identity contains duplicate {label} path: {path}");
            if (!ApiSnapshotFileHash.IsSha256(pair.Value))
                failures.Add($"candidate identity {label} hash is not a SHA-256 value: {path}");
        }

        return result;
    }


    private static void ValidateCanonicalTextFileHash(string path, string expectedHash, string label,
        List<string> failures)
    {
        if (!ApiSnapshotFileHash.IsSha256(expectedHash))
            return;
        if (!File.Exists(path))
        {
            failures.Add($"{label} is missing: {path}");
            return;
        }

        try
        {
            var actualHash = ApiSnapshotFileHash.ComputeCanonicalTextSha256(path);
            if (!string.Equals(expectedHash, actualHash, StringComparison.OrdinalIgnoreCase))
                failures.Add($"{label} SHA-256 mismatch: expected={expectedHash}; actual={actualHash}");
        }
        catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
        {
            failures.Add($"{label} could not be hashed: {path}; {exception.Message}");
        }
    }

    private static void ValidateHash(string label, string expectedHash, string actualHash, List<string> failures)
    {
        if (!ApiSnapshotFileHash.IsSha256(expectedHash))
        {
            failures.Add($"candidate identity {label} hash is not a SHA-256 value");
            return;
        }

        if (!string.Equals(expectedHash, actualHash, StringComparison.OrdinalIgnoreCase))
            failures.Add($"candidate identity {label} SHA-256 mismatch: expected={expectedHash}; actual={actualHash}");
    }

    private static bool TryGetCandidateSourceManifestSha256(string repositoryRoot, out string hash, out string error)
    {
        if (!TryRunGit(repositoryRoot,
                new[] { "ls-files", "--cached", "--others", "--exclude-standard", "--" }.Concat(CandidateSourceScope),
                out var paths, out error))
        {
            hash = string.Empty;
            return false;
        }

        var lines = new List<string>();
        foreach (var relativePath in paths.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                     .Select(NormalizeRelativePath).Where(path => !string.IsNullOrWhiteSpace(path))
                     .Distinct(StringComparer.Ordinal).OrderBy(path => path, StringComparer.Ordinal))
        {
            var fullPath = ResolveWithinRoot(repositoryRoot, relativePath);
            if (fullPath is null)
            {
                hash = string.Empty;
                error = $"candidate source path escapes repository root: {relativePath}";
                return false;
            }

            if (!File.Exists(fullPath))
                continue;

            if (!TryRunGit(repositoryRoot,
                    new[] { "hash-object", "--path=" + relativePath, "--", fullPath },
                    out var canonicalBlobHash, out var canonicalHashError))
            {
                hash = string.Empty;
                error = $"unable to hash candidate source path {relativePath}: {canonicalHashError}";
                return false;
            }

            var normalizedBlobHash = canonicalBlobHash.Trim();
            if (string.IsNullOrWhiteSpace(normalizedBlobHash))
            {
                hash = string.Empty;
                error = $"git returned an empty canonical hash for candidate source path: {relativePath}";
                return false;
            }

            lines.Add($"{relativePath}|{normalizedBlobHash}");
        }

        hash = ComputeUtf8Sha256(string.Join("\n", lines));
        return true;
    }

    private static bool TryRunGit(string repositoryRoot, IEnumerable<string> arguments,
        out string output, out string error)
    {
        try
        {
            using var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = "git",
                    WorkingDirectory = repositoryRoot,
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
            {
                output = string.Empty;
                error = "git process did not start";
                return false;
            }

            output = process.StandardOutput.ReadToEnd();
            var standardError = process.StandardError.ReadToEnd();
            process.WaitForExit();
            if (process.ExitCode != 0)
            {
                error = string.IsNullOrWhiteSpace(standardError)
                    ? $"git exited with code {process.ExitCode}"
                    : standardError.Trim();
                return false;
            }

            error = string.Empty;
            return true;
        }
        catch (Exception exception) when (exception is IOException or InvalidOperationException or System.ComponentModel.Win32Exception)
        {
            output = string.Empty;
            error = exception.Message;
            return false;
        }
    }

    private static string ComputeUtf8Sha256(string value)
    {
        using var sha256 = SHA256.Create();
        return BitConverter.ToString(sha256.ComputeHash(Encoding.UTF8.GetBytes(value)))
            .Replace("-", string.Empty, StringComparison.Ordinal);
    }

    private static string? ResolveWithinRoot(string root, string relativePath)
    {
        if (string.IsNullOrWhiteSpace(relativePath) || Path.IsPathRooted(relativePath))
            return null;

        try
        {
            var fullRoot = Path.GetFullPath(root);
            var fullPath = Path.GetFullPath(Path.Combine(fullRoot,
                relativePath.Replace('/', Path.DirectorySeparatorChar)));
            var relativeToRoot = Path.GetRelativePath(fullRoot, fullPath);
            if (relativeToRoot == ".."
                || relativeToRoot.StartsWith(".." + Path.DirectorySeparatorChar, StringComparison.Ordinal)
                || Path.IsPathRooted(relativeToRoot))
                return null;
            return fullPath;
        }
        catch (ArgumentException)
        {
            return null;
        }
    }

    private static string NormalizeRelativePath(string path)
    {
        var normalized = (path ?? string.Empty).Trim().Replace('\\', '/');
        while (normalized.StartsWith("./", StringComparison.Ordinal))
            normalized = normalized[2..];
        return normalized;
    }

    private static string? GetPackageId(string path)
    {
        var fileName = Path.GetFileName(path);
        return RequiredPackageIds.FirstOrDefault(packageId =>
            fileName.StartsWith(packageId + ".", StringComparison.OrdinalIgnoreCase));
    }
}
