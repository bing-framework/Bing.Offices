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
var approval = arguments.GetValueOrDefault("approval");
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
    CandidateIdentityVerifier.Validate(
        baseline!, candidateAssemblyPaths, packages, repository, failures, approval);
}

foreach (var tfm in targetFrameworks)
{
    var paths = new[]
    {
        Path.Combine(root, "netstandard2.0", "Bing.Offices.Abstractions.dll"),
        Path.Combine(root, "netstandard2.0", "Bing.Offices.Core.dll"),
        Path.Combine(root, tfm, "Bing.Offices.Npoi.dll"),
        Path.Combine(root, tfm, "Bing.Offices.MiniExcel.dll"),
        Path.Combine(root, tfm, "Bing.Offices.ClosedXml.dll"),
        Path.Combine(root, tfm, "Bing.Offices.ExcelDataReader.dll"),
        Path.Combine(root, tfm, "Bing.Offices.SpreadCheetah.dll"),
        Path.Combine(root, tfm, "Bing.Offices.AsposeCells.dll")
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
        candidateAssemblyPaths, packages, repository, candidate.BaselineCommit, failures, approval);

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
            Path.Combine(root, "net8.0", "Bing.Offices.MiniExcel.dll"),
        ["net6.0/Bing.Offices.ClosedXml.dll"] =
            Path.Combine(root, "net6.0", "Bing.Offices.ClosedXml.dll"),
        ["net8.0/Bing.Offices.ClosedXml.dll"] =
            Path.Combine(root, "net8.0", "Bing.Offices.ClosedXml.dll"),
        ["net6.0/Bing.Offices.ExcelDataReader.dll"] =
            Path.Combine(root, "net6.0", "Bing.Offices.ExcelDataReader.dll"),
        ["net8.0/Bing.Offices.ExcelDataReader.dll"] =
            Path.Combine(root, "net8.0", "Bing.Offices.ExcelDataReader.dll"),
        ["net6.0/Bing.Offices.SpreadCheetah.dll"] =
            Path.Combine(root, "net6.0", "Bing.Offices.SpreadCheetah.dll"),
        ["net8.0/Bing.Offices.SpreadCheetah.dll"] =
            Path.Combine(root, "net8.0", "Bing.Offices.SpreadCheetah.dll"),
        ["net6.0/Bing.Offices.AsposeCells.dll"] =
            Path.Combine(root, "net6.0", "Bing.Offices.AsposeCells.dll"),
        ["net8.0/Bing.Offices.AsposeCells.dll"] =
            Path.Combine(root, "net8.0", "Bing.Offices.AsposeCells.dll")
    };

/// <summary>
/// 公共 API 快照基线文档模型。
/// </summary>
public sealed class ApiBaselineDocument
{
    /// <summary>
    /// 获取或设置基线文档格式标识。
    /// </summary>
    public string Schema { get; set; } = "";
    /// <summary>
    /// 获取或设置生成快照的工具版本。
    /// </summary>
    public string GeneratorVersion { get; set; } = "";
    /// <summary>
    /// 获取或设置基线对应的仓库提交标识。
    /// </summary>
    public string BaselineCommit { get; set; } = "";
    /// <summary>
    /// 获取或设置批准基线的人员标识。
    /// </summary>
    public string ApprovedBy { get; set; } = "";
    /// <summary>
    /// 获取或设置批准基线的时间文本。
    /// </summary>
    public string ApprovedAt { get; set; } = "";
    /// <summary>
    /// 获取或设置候选程序集、包和源清单身份。
    /// </summary>
    public ApiCandidateIdentity? CandidateIdentity { get; set; }
    /// <summary>
    /// 获取或设置按目标框架和程序集组织的 API 快照。
    /// </summary>
    public Dictionary<string, Dictionary<string, ApiSnapshotRecord>> Assemblies { get; set; } = new(StringComparer.Ordinal);
}

/// <summary>
/// 候选构建产物的身份记录模型。
/// </summary>
public sealed class ApiCandidateIdentity
{
    /// <summary>
    /// 获取或设置候选身份算法格式。
    /// </summary>
    public string ArtifactIdentityFormat { get; set; } = "";
    /// <summary>
    /// 获取或设置候选身份使用的基线提交标识。
    /// </summary>
    public string BaseCommit { get; set; } = "";
    /// <summary>
    /// 获取或设置采集身份时的工作区状态。
    /// </summary>
    public string WorktreeState { get; set; } = "";
    /// <summary>
    /// 获取或设置候选源文件清单的 SHA-256 哈希。
    /// </summary>
    public string CandidateSourceManifestSha256 { get; set; } = "";
    /// <summary>
    /// 获取或设置 Breaking Change 审批文件的相对路径。
    /// </summary>
    public string BreakingApprovalArtifact { get; set; } = "";
    /// <summary>
    /// 获取或设置 Breaking Change 审批文件的 SHA-256 哈希。
    /// </summary>
    public string BreakingApprovalSha256 { get; set; } = "";
    /// <summary>
    /// 获取或设置候选程序集相对路径到身份哈希的映射。
    /// </summary>
    public Dictionary<string, string> AssemblyFiles { get; set; } = new(StringComparer.Ordinal);
    /// <summary>
    /// 获取或设置候选 NuGet 包相对路径到身份哈希的映射。
    /// </summary>
    public Dictionary<string, string> NupkgFiles { get; set; } = new(StringComparer.Ordinal);
}

/// <summary>
/// 单个程序集的 API 快照记录模型。
/// </summary>
public sealed class ApiSnapshotRecord
{
    /// <summary>
    /// 获取或设置快照中的成员数量。
    /// </summary>
    public int MemberCount { get; set; }
    /// <summary>
    /// 获取或设置快照成员清单的哈希。
    /// </summary>
    public string Hash { get; set; } = "";
    /// <summary>
    /// 获取或设置按稳定顺序保存的成员签名清单。
    /// </summary>
    public List<string> Lines { get; set; } = new();
}

/// <summary>
/// API 快照之间的成员差异模型。
/// </summary>
public sealed class ApiMemberDiff
{
    /// <summary>
    /// 获取或设置实际快照新增的成员签名。
    /// </summary>
    public List<string> Added { get; set; } = new();
    /// <summary>
    /// 获取或设置实际快照移除的成员签名。
    /// </summary>
    public List<string> Removed { get; set; } = new();
}

/// <summary>
/// 采集并校验候选构建身份的内部服务。
/// </summary>
internal static class CandidateIdentityVerifier
{
    /// <summary>
    /// 候选产物身份记录使用的逻辑格式标识。
    /// </summary>
    private const string ArtifactIdentityFormat = "logical-v3";
    /// <summary>
    /// 默认 Breaking Change 审批文件的仓库相对路径。
    /// </summary>
    internal const string BreakingApprovalPath =
        "ai_docs/tasks/BING-OFFICES-ADVANCED-EXCEL-ROADMAP-20260926-001/api-approval.md";

    /// <summary>
    /// 候选身份必须覆盖的生产 NuGet 包标识。
    /// </summary>
    private static readonly string[] RequiredPackageIds =
    {
        "Bing.Offices.Abstractions",
        "Bing.Offices.Core",
        "Bing.Offices.Npoi",
        "Bing.Offices.MiniExcel",
        "Bing.Offices.ClosedXml",
        "Bing.Offices.ExcelDataReader",
        "Bing.Offices.SpreadCheetah",
        "Bing.Offices.AsposeCells"
    };

    /// <summary>
    /// 参与候选源清单哈希的仓库相对路径范围。
    /// </summary>
    private static readonly string[] CandidateSourceScope =
    {
        // 只绑定会影响 Release 程序集或 nupkg 内容的输入，避免文档、测试和 CI 配置导致 API 基线失效。
        "src", "asset", "build/ApiSnapshot", "README.md", "LICENSE", ".gitattributes",
        "framework.props", "common.props", "Directory.Build.targets", "version.props", "version.dev.props"
    };

    /// <summary>
    /// 从当前程序集、包和仓库状态采集候选身份。
    /// </summary>
    /// <param name="assemblyPaths">待采集的程序集路径。</param>
    /// <param name="packagesRoot">NuGet 包目录；为空时不采集包身份。</param>
    /// <param name="repositoryRoot">候选仓库根目录。</param>
    /// <param name="candidateCommit">候选基线提交标识。</param>
    /// <param name="failures">用于追加采集失败信息的集合。</param>
    /// <param name="approvalPath">审批文件相对路径；省略时使用默认路径。</param>
    /// <returns>采集到的候选身份。</returns>
    public static ApiCandidateIdentity Capture(
        IReadOnlyDictionary<string, string> assemblyPaths,
        string? packagesRoot,
        string repositoryRoot,
        string candidateCommit,
        List<string> failures,
        string? approvalPath = null)
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

        if (!TryNormalizeApprovalPath(approvalPath, out var normalizedApprovalPath, out var approvalError))
        {
            failures.Add($"capture: approval file path is invalid: {approvalError}");
        }
        else
        {
            identity.BreakingApprovalArtifact = normalizedApprovalPath;
            var fullApprovalPath = ResolveWithinRoot(repositoryRoot, normalizedApprovalPath);
            if (fullApprovalPath is null)
            {
                failures.Add($"capture: approval file path escapes repository root: {normalizedApprovalPath}");
            }
            else if (!File.Exists(fullApprovalPath))
            {
                failures.Add($"capture: approval file is missing: {fullApprovalPath}");
            }
            else
            {
                try
                {
                    identity.BreakingApprovalSha256 =
                        ApiSnapshotFileHash.ComputeCanonicalTextSha256(fullApprovalPath);
                }
                catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
                {
                    failures.Add(
                        $"capture: approval file could not be hashed: {fullApprovalPath}; {exception.Message}");
                }
            }
        }

        if (!string.IsNullOrWhiteSpace(packagesRoot))
            identity.NupkgFiles = CaptureNupkgFiles(packagesRoot, identity.AssemblyFiles, assemblyPaths, failures);

        return identity;
    }

    /// <summary>
    /// 校验候选程序集、包、仓库清单和审批文件身份。
    /// </summary>
    /// <param name="baseline">包含候选身份的 API 基线。</param>
    /// <param name="assemblyPaths">待校验的程序集路径。</param>
    /// <param name="packagesRoot">NuGet 包目录。</param>
    /// <param name="repositoryRoot">候选仓库根目录。</param>
    /// <param name="failures">用于追加校验失败信息的集合。</param>
    /// <param name="approvalPath">审批文件相对路径；省略时使用默认路径。</param>
    public static void Validate(
        ApiBaselineDocument baseline,
        IReadOnlyDictionary<string, string> assemblyPaths,
        string? packagesRoot,
        string repositoryRoot,
        List<string> failures,
        string? approvalPath = null)
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
        if (!TryNormalizeApprovalPath(approvalPath, out var normalizedApprovalPath, out var approvalError))
            failures.Add($"candidate identity approval file path is invalid: {approvalError}");
        else
            ValidateApprovalFile(identity, repositoryRoot, normalizedApprovalPath, failures);
    }

    /// <summary>
    /// 校验候选程序集身份并返回当前程序集哈希。
    /// </summary>
    /// <param name="recordedFiles">基线记录的程序集哈希映射。</param>
    /// <param name="assemblyPaths">待校验的程序集路径映射。</param>
    /// <param name="failures">用于追加校验失败信息的集合。</param>
    /// <returns>当前程序集相对路径到身份哈希的映射。</returns>
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

    /// <summary>
    /// 校验候选 NuGet 包的路径、数量和内容身份。
    /// </summary>
    /// <param name="recordedFiles">基线记录的 NuGet 包哈希映射。</param>
    /// <param name="packagesRoot">待校验的 NuGet 包目录。</param>
    /// <param name="assemblyFiles">已校验的程序集身份哈希映射。</param>
    /// <param name="assemblyPaths">候选程序集路径映射。</param>
    /// <param name="failures">用于追加校验失败信息的集合。</param>
    private static void ValidateNupkgFiles(
        IReadOnlyDictionary<string, string>? recordedFiles,
        string? packagesRoot,
        IReadOnlyDictionary<string, string> assemblyFiles,
        IReadOnlyDictionary<string, string> assemblyPaths,
        List<string> failures)
    {
        var recorded = NormalizeHashMap(recordedFiles, "nupkg", failures);
        if (recorded.Count < RequiredPackageIds.Length)
            failures.Add($"candidate identity must contain hashes for the {RequiredPackageIds.Length} production nupkg files");
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

    /// <summary>
    /// 校验候选身份中的审批文件路径和内容哈希。
    /// </summary>
    /// <param name="identity">待校验的候选身份。</param>
    /// <param name="repositoryRoot">候选仓库根目录。</param>
    /// <param name="expectedPath">当前配置要求的审批文件路径。</param>
    /// <param name="failures">用于追加校验失败信息的集合。</param>
    private static void ValidateApprovalFile(ApiCandidateIdentity identity, string repositoryRoot,
        string expectedPath, List<string> failures)
    {
        if (!TryNormalizeApprovalPath(identity.BreakingApprovalArtifact, out var path, out var approvalError))
        {
            failures.Add($"candidate identity approval file path is invalid: {approvalError}");
            return;
        }

        if (!string.Equals(identity.BreakingApprovalArtifact, path, StringComparison.Ordinal))
        {
            failures.Add($"candidate identity approval path is not normalized: {identity.BreakingApprovalArtifact}");
            return;
        }

        if (!string.Equals(path, expectedPath, StringComparison.Ordinal))
        {
            failures.Add($"candidate identity approval path does not match configured path: "
                + $"expected={expectedPath}; actual={path}");
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

    /// <summary>
    /// 规范化并校验相对于仓库根目录的审批文件路径。
    /// </summary>
    /// <param name="value">待处理的路径；为空时使用默认路径。</param>
    /// <param name="normalized">返回使用正斜杠且已消除相对片段的路径。</param>
    /// <param name="error">路径无效时返回错误说明。</param>
    /// <returns>路径有效时返回 <see langword="true"/>，否则返回 <see langword="false"/>。</returns>
    private static bool TryNormalizeApprovalPath(
        string? value, out string normalized, out string error)
    {
        var candidate = string.IsNullOrWhiteSpace(value) ? BreakingApprovalPath : value.Trim();
        candidate = candidate.Replace('\\', '/');
        if (candidate.Length == 0)
        {
            normalized = string.Empty;
            error = "path is empty";
            return false;
        }

        if (candidate.StartsWith("/", StringComparison.Ordinal)
            || candidate.Contains(':', StringComparison.Ordinal))
        {
            normalized = string.Empty;
            error = $"path must be relative: {value}";
            return false;
        }

        var segments = new List<string>();
        foreach (var segment in candidate.Split('/', StringSplitOptions.RemoveEmptyEntries))
        {
            if (string.Equals(segment, ".", StringComparison.Ordinal))
                continue;
            if (string.Equals(segment, "..", StringComparison.Ordinal))
            {
                if (segments.Count == 0)
                {
                    normalized = string.Empty;
                    error = $"path escapes repository root: {value}";
                    return false;
                }

                segments.RemoveAt(segments.Count - 1);
                continue;
            }

            segments.Add(segment);
        }

        if (segments.Count == 0)
        {
            normalized = string.Empty;
            error = $"path is empty: {value}";
            return false;
        }

        normalized = string.Join("/", segments);
        error = string.Empty;
        return true;
    }

    /// <summary>
    /// 采集生产 NuGet 包的规范化身份哈希。
    /// </summary>
    /// <param name="packagesRoot">NuGet 包目录。</param>
    /// <param name="assemblyFiles">候选程序集身份哈希映射。</param>
    /// <param name="assemblyPaths">候选程序集路径映射。</param>
    /// <param name="failures">用于追加采集失败信息的集合。</param>
    /// <returns>NuGet 包相对路径到身份哈希的映射。</returns>
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

    /// <summary>
    /// 采集候选程序集的公共 API 身份哈希。
    /// </summary>
    /// <param name="assemblyPaths">候选程序集路径映射。</param>
    /// <param name="operation">当前采集或校验操作名称。</param>
    /// <param name="failures">用于追加采集失败信息的集合。</param>
    /// <returns>程序集相对路径到身份哈希的映射。</returns>
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

    /// <summary>
    /// 计算程序集公共 API 快照的身份哈希。
    /// </summary>
    /// <param name="path">程序集文件路径。</param>
    /// <param name="additionalAssemblyPaths">解析快照所需的附加程序集路径。</param>
    /// <returns>程序集身份哈希。</returns>
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

    /// <summary>
    /// 计算 NuGet 包规范化内容的身份哈希。
    /// </summary>
    /// <param name="packagePath">NuGet 包文件路径。</param>
    /// <param name="assemblyFiles">候选程序集身份哈希映射。</param>
    /// <param name="assemblyPaths">候选程序集路径映射。</param>
    /// <returns>NuGet 包身份哈希。</returns>
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

    /// <summary>
    /// 计算单个 NuGet 包资产的规范化身份哈希。
    /// </summary>
    /// <param name="entry">待处理的 ZIP 资产。</param>
    /// <param name="relativePath">资产在包内的相对路径。</param>
    /// <param name="assemblyFiles">候选程序集身份哈希映射。</param>
    /// <param name="assemblyPaths">候选程序集路径映射。</param>
    /// <returns>带资产类型前缀的身份哈希。</returns>
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

    /// <summary>
    /// 计算 XML 包资产在格式差异归一化后的身份哈希。
    /// </summary>
    /// <param name="stream">XML 资产输入流。</param>
    /// <param name="normalizeRepositoryMetadata">是否移除 NuGet 注入的仓库提交元数据。</param>
    /// <returns>规范化 XML 的身份哈希。</returns>
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

    /// <summary>
    /// 判断包内路径是否属于 XML 资产。
    /// </summary>
    /// <param name="relativePath">包内相对路径。</param>
    /// <returns>扩展名为 XML、NUSPEC 或 RELS 时返回 <see langword="true" />，否则返回 <see langword="false" />。</returns>
    private static bool IsXmlPackageEntry(string relativePath) =>
        Path.GetExtension(relativePath).ToLowerInvariant() is ".xml" or ".nuspec" or ".rels";

    /// <summary>
    /// 判断包内路径是否应按规范化文本处理。
    /// </summary>
    /// <param name="relativePath">包内相对路径。</param>
    /// <returns>路径属于受支持的文本资产时返回 <see langword="true" />，否则返回 <see langword="false" />。</returns>
    private static bool IsCanonicalTextPackageEntry(string relativePath)
    {
        var fileName = Path.GetFileName(relativePath);
        if (string.Equals(fileName, "LICENSE", StringComparison.OrdinalIgnoreCase)
            || fileName.StartsWith("README", StringComparison.OrdinalIgnoreCase))
            return true;

        return Path.GetExtension(relativePath).ToLowerInvariant() is ".md" or ".txt" or ".json"
            or ".props" or ".targets" or ".config";
    }

    /// <summary>
    /// 规范化身份哈希映射中的路径并记录非法哈希。
    /// </summary>
    /// <param name="values">待处理的哈希映射；为 <see langword="null" /> 时返回空映射。</param>
    /// <param name="label">映射类型标签。</param>
    /// <param name="failures">用于追加校验失败信息的集合。</param>
    /// <returns>使用规范化相对路径的哈希映射。</returns>
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


    /// <summary>
    /// 计算并校验规范化文本文件的 SHA-256 哈希。
    /// </summary>
    /// <param name="path">待校验的文本文件路径。</param>
    /// <param name="expectedHash">预期 SHA-256 哈希。</param>
    /// <param name="label">错误信息中的文件标签。</param>
    /// <param name="failures">用于追加校验失败信息的集合。</param>
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

    /// <summary>
    /// 比较预期和实际身份哈希并记录差异。
    /// </summary>
    /// <param name="label">错误信息中的身份标签。</param>
    /// <param name="expectedHash">预期 SHA-256 哈希。</param>
    /// <param name="actualHash">实际 SHA-256 哈希。</param>
    /// <param name="failures">用于追加校验失败信息的集合。</param>
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

    /// <summary>
    /// 计算候选源文件清单的 Git 规范化 SHA-256 哈希。
    /// </summary>
    /// <param name="repositoryRoot">候选仓库根目录。</param>
    /// <param name="hash">成功时返回源文件清单哈希。</param>
    /// <param name="error">失败时返回错误说明。</param>
    /// <returns>成功计算时返回 <see langword="true" />，否则返回 <see langword="false" />。</returns>
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

    /// <summary>
    /// 运行 Git 命令并以结果参数返回标准输出或错误。
    /// </summary>
    /// <param name="repositoryRoot">Git 命令的工作目录。</param>
    /// <param name="arguments">Git 命令参数。</param>
    /// <param name="output">成功时返回标准输出。</param>
    /// <param name="error">失败时返回错误说明。</param>
    /// <returns>命令成功退出时返回 <see langword="true" />，否则返回 <see langword="false" />。</returns>
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

    /// <summary>
    /// 计算文本 UTF-8 字节序列的 SHA-256 哈希。
    /// </summary>
    /// <param name="value">待计算哈希的文本。</param>
    /// <returns>大写十六进制 SHA-256 哈希。</returns>
    private static string ComputeUtf8Sha256(string value)
    {
        using var sha256 = SHA256.Create();
        return BitConverter.ToString(sha256.ComputeHash(Encoding.UTF8.GetBytes(value)))
            .Replace("-", string.Empty, StringComparison.Ordinal);
    }

    /// <summary>
    /// 将相对路径解析为根目录内的绝对路径。
    /// </summary>
    /// <param name="root">允许访问的根目录。</param>
    /// <param name="relativePath">待解析的相对路径。</param>
    /// <returns>路径位于根目录内时返回绝对路径，否则返回 <see langword="null" />。</returns>
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

    /// <summary>
    /// 规范化相对路径的分隔符和当前目录片段。
    /// </summary>
    /// <param name="path">待规范化的路径。</param>
    /// <returns>使用正斜杠且移除前导当前目录片段的路径。</returns>
    private static string NormalizeRelativePath(string path)
    {
        var normalized = (path ?? string.Empty).Trim().Replace('\\', '/');
        while (normalized.StartsWith("./", StringComparison.Ordinal))
            normalized = normalized[2..];
        return normalized;
    }

    /// <summary>
    /// 从 NuGet 包文件名识别生产包标识。
    /// </summary>
    /// <param name="path">NuGet 包路径或文件名。</param>
    /// <returns>识别到的生产包标识；无法识别时返回 <see langword="null" />。</returns>
    private static string? GetPackageId(string path)
    {
        var fileName = Path.GetFileName(path);
        return RequiredPackageIds.FirstOrDefault(packageId =>
            fileName.StartsWith(packageId + ".", StringComparison.OrdinalIgnoreCase));
    }
}
