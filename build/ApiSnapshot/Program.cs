using System.Text.Json;
using Bing.Offices.ApiSnapshot;

var arguments = ParseArguments(args);
var root = Path.GetFullPath(arguments.GetValueOrDefault("root") ?? "output/release");
var baselinePath = Path.GetFullPath(arguments.GetValueOrDefault("baseline") ?? "build/api-snapshot-baseline.json");
var output = arguments.GetValueOrDefault("output");
var baseline = JsonSerializer.Deserialize<Dictionary<string, Dictionary<string, string>>>(
    File.ReadAllText(baselinePath, System.Text.Encoding.UTF8))
    ?? throw new InvalidOperationException("API snapshot baseline is empty.");
var failures = new List<string>();

foreach (var tfm in new[] { "netcoreapp3.1", "net6.0", "net8.0" })
{
    var paths = new[]
    {
        Path.Combine(root, "netstandard2.0", "Bing.Offices.Abstractions.dll"),
        Path.Combine(root, "netstandard2.0", "Bing.Offices.Core.dll"),
        Path.Combine(root, tfm, "Bing.Offices.Npoi.dll")
    };
    var additionalPaths = paths.Select(Path.GetDirectoryName).Where(path => path is not null)
        .SelectMany(path => Directory.EnumerateFiles(path!, "*.dll"));
    var snapshots = paths.Select(path => PublicApiSnapshot.Load(path, additionalPaths)).ToDictionary(
        snapshot => snapshot.AssemblyName, snapshot => snapshot, StringComparer.Ordinal);

    if (!baseline.TryGetValue(tfm, out var expected))
        failures.Add($"{tfm}: missing baseline");
    else
    {
        foreach (var pair in expected)
        {
            if (!snapshots.TryGetValue(pair.Key, out var actual))
                failures.Add($"{tfm}/{pair.Key}: missing assembly");
            else if (!string.Equals(pair.Value, actual.Hash, StringComparison.Ordinal))
                failures.Add($"{tfm}/{pair.Key}: expected={pair.Value}; actual={actual.Hash}");
        }
    }

    if (output is not null)
    {
        var directory = Path.GetFullPath(output);
        Directory.CreateDirectory(directory);
        var document = snapshots.ToDictionary(pair => pair.Key, pair => new
        {
            memberCount = pair.Value.MemberCount,
            hash = pair.Value.Hash,
            lines = pair.Value.Lines
        }, StringComparer.Ordinal);
        File.WriteAllText(Path.Combine(directory, $"api-snapshot-{tfm}.json"),
            JsonSerializer.Serialize(document, new JsonSerializerOptions { WriteIndented = true }),
            System.Text.Encoding.UTF8);
    }
}

if (failures.Count > 0)
{
    Console.Error.WriteLine(string.Join(Environment.NewLine, failures));
    return 1;
}

Console.WriteLine("API snapshot comparison passed for netcoreapp3.1, net6.0, net8.0.");
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
