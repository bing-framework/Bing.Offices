using System.Text.Json;
using Bing.Offices.ApiSnapshot;

const string schema = "bing.offices.public-api.v2";
const string generatorVersion = "2.0.0";

var arguments = ParseArguments(args);
var root = Path.GetFullPath(arguments.GetValueOrDefault("root") ?? "output/release");
var baselinePath = Path.GetFullPath(arguments.GetValueOrDefault("baseline") ?? "build/api-snapshot-baseline.json");
var output = arguments.GetValueOrDefault("output");
var dependencies = arguments.GetValueOrDefault("dependencies");
var captureOnly = string.Equals(arguments.GetValueOrDefault("capture"), "true", StringComparison.OrdinalIgnoreCase);
var targetFrameworks = new[] { "net6.0", "net8.0" };
var jsonOptions = new JsonSerializerOptions { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase };

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

foreach (var tfm in targetFrameworks)
{
    var paths = new[]
    {
        Path.Combine(root, "netstandard2.0", "Bing.Offices.Abstractions.dll"),
        Path.Combine(root, "netstandard2.0", "Bing.Offices.Core.dll"),
        Path.Combine(root, tfm, "Bing.Offices.Npoi.dll")
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

public sealed class ApiBaselineDocument
{
    public string Schema { get; set; } = "";
    public string GeneratorVersion { get; set; } = "";
    public string BaselineCommit { get; set; } = "";
    public string ApprovedBy { get; set; } = "";
    public string ApprovedAt { get; set; } = "";
    public Dictionary<string, Dictionary<string, ApiSnapshotRecord>> Assemblies { get; set; } = new(StringComparer.Ordinal);
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
