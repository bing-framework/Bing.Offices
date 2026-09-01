using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace Bing.Offices.ApiSnapshot;

public sealed class PublicApiSnapshot
{
    private PublicApiSnapshot(string assemblyName, int memberCount, string hash, IReadOnlyList<string> lines)
    {
        AssemblyName = assemblyName;
        MemberCount = memberCount;
        Hash = hash;
        Lines = lines;
    }

    public string AssemblyName { get; }

    public int MemberCount { get; }

    public string Hash { get; }

    public IReadOnlyList<string> Lines { get; }

    public static PublicApiSnapshot Load(string assemblyPath, IEnumerable<string>? additionalAssemblyPaths = null)
    {
        var fullAssemblyPath = Path.GetFullPath(assemblyPath);
        var resolverPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        AddDirectory(resolverPaths, Path.GetDirectoryName(fullAssemblyPath));
        var trustedPlatformAssemblyPaths = GetTrustedPlatformAssemblyPaths();
        var coreAssemblyPath = Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), "System.Private.CoreLib.dll");
        if (!File.Exists(coreAssemblyPath))
            coreAssemblyPath = trustedPlatformAssemblyPaths.FirstOrDefault(path =>
                string.Equals(Path.GetFileName(path), "System.Private.CoreLib.dll", StringComparison.OrdinalIgnoreCase))
                ?? typeof(object).Assembly.Location;
        if (!File.Exists(coreAssemblyPath))
            coreAssemblyPath = Path.Combine(AppContext.BaseDirectory, "System.Private.CoreLib.dll");
        if (!File.Exists(coreAssemblyPath))
            throw new FileNotFoundException("Unable to locate the current runtime core assembly.", coreAssemblyPath);
        AddPath(resolverPaths, coreAssemblyPath);

        if (additionalAssemblyPaths is not null)
        {
            foreach (var path in additionalAssemblyPaths)
                AddPath(resolverPaths, path);
        }

        var trustedPlatformPaths = new List<string>(trustedPlatformAssemblyPaths);
        foreach (var path in trustedPlatformPaths)
        {
            AddPath(resolverPaths, path);
        }

        AddKnownPackageAssemblies(resolverPaths);

        const string coreAssemblyName = "System.Private.CoreLib";
        var resolver = new AssemblyPathResolver(resolverPaths, trustedPlatformPaths, coreAssemblyName,
            coreAssemblyPath);
        using var loadContext = new MetadataLoadContext(resolver, coreAssemblyName);
        var assembly = loadContext.LoadFromAssemblyPath(fullAssemblyPath);
        var lines = BuildCanonicalLines(assembly);
        var memberCount = lines.Count(line => !line.StartsWith("type|", StringComparison.Ordinal));
        var text = string.Join("\n", lines);
        using var sha256 = SHA256.Create();
        var hash = BitConverter.ToString(sha256.ComputeHash(Encoding.UTF8.GetBytes(text)))
            .Replace("-", string.Empty, StringComparison.Ordinal);
        return new PublicApiSnapshot(assembly.GetName().Name ?? string.Empty, memberCount, hash, lines);
    }

    private static List<string> BuildCanonicalLines(Assembly assembly)
    {
        var lines = new List<string>();
        foreach (var type in assembly.GetExportedTypes().OrderBy(type => type.FullName, StringComparer.Ordinal))
        {
            lines.Add($"type|{type.FullName}|generic={type.GetGenericArguments().Length}");
            foreach (var constructor in GetGovernedConstructors(type))
                lines.Add(FormatConstructor(type, constructor));
            foreach (var property in GetGovernedProperties(type))
                lines.Add(FormatProperty(type, property));
            foreach (var field in GetGovernedFields(type))
                lines.Add(FormatField(type, field));
            foreach (var method in GetGovernedMethods(type).Where(method => !method.IsSpecialName))
                lines.Add(FormatMethod(type, method));
        }

        return lines.OrderBy(line => line, StringComparer.Ordinal).ToList();
    }

    private static string FormatConstructor(Type type, ConstructorInfo constructor) =>
        $"constructor|{type.FullName}|{FormatParameters(constructor.GetParameters())}";

    private static string FormatProperty(Type type, PropertyInfo property) =>
        $"property|{type.FullName}.{property.Name}|{FormatTypeName(property.PropertyType)}";

    private static string FormatField(Type type, FieldInfo field) =>
        $"field|{type.FullName}.{field.Name}|{FormatTypeName(field.FieldType)}";

    private static string FormatMethod(Type type, MethodInfo method) =>
        $"method|{type.FullName}.{method.Name}|{(method.IsStatic ? "static" : "instance")}|{FormatTypeName(method.ReturnType)}|"
        + $"{FormatParameters(method.GetParameters())}|generic={method.GetGenericArguments().Length}";

    private static string FormatParameters(IReadOnlyList<ParameterInfo> parameters) =>
        string.Join(",", parameters.Select(parameter => FormatTypeName(parameter.ParameterType)));

    private static string FormatTypeName(Type type)
    {
        if (type.IsByRef)
            return $"{FormatTypeName(type.GetElementType()!)}&";
        if (type.IsPointer)
            return $"{FormatTypeName(type.GetElementType()!)}*";
        if (type.IsArray)
            return $"{FormatTypeName(type.GetElementType()!)}{new string(',', type.GetArrayRank() - 1)}[]";
        if (type.IsGenericParameter)
            return type.FullName ?? string.Empty;
        if (!type.IsGenericType)
            return type.FullName ?? type.Name;

        var definition = type.GetGenericTypeDefinition().FullName ?? type.Name;
        var arguments = type.GetGenericArguments()
            .Select(argument => $"[{FormatTypeName(argument)}]");
        return $"{definition}[{string.Join(",", arguments)}]";
    }

    private static List<string> GetTrustedPlatformAssemblyPaths()
    {
        var value = AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES") as string;
        return string.IsNullOrWhiteSpace(value)
            ? new List<string>()
            : value.Split(Path.PathSeparator).Where(File.Exists).ToList();
    }

    private static void AddKnownPackageAssemblies(ISet<string> paths)
    {
        var packageRoot = Environment.GetEnvironmentVariable("NUGET_PACKAGES");
        if (string.IsNullOrWhiteSpace(packageRoot))
        {
            packageRoot = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                ".nuget",
                "packages");
        }

        if (!Directory.Exists(packageRoot))
            return;

        foreach (var package in new[]
        {
            ("aspectcore.extensions.reflection", "2.4.0"),
            ("bing.utils", "1.5.0"),
            ("csvhelper", "33.1.0"),
            ("microsoft.extensions.dependencyinjection.abstractions", "2.2.0"),
            ("microsoft.extensions.dependencyinjection.abstractions", "3.1.10"),
            ("microsoft.extensions.dependencyinjection.abstractions", "5.0.0"),
            ("microsoft.extensions.dependencyinjection.abstractions", "6.0.0"),
            ("microsoft.extensions.dependencyinjection.abstractions", "7.0.0"),
            ("microsoft.extensions.dependencyinjection.abstractions", "8.0.0"),
            ("npoi", "2.7.4"),
            ("bouncycastle.cryptography", "2.4.0"),
            ("enums.net", "5.0.0"),
            ("extendednumerics.bigdecimal", "2025.1001.2.129"),
            ("mathnet.numerics.signed", "5.0.0"),
            ("microsoft.io.recyclablememorystream", "3.0.1"),
            ("sharpziplib", "1.4.2"),
            ("sixlabors.fonts", "1.0.1"),
            ("sixlabors.imagesharp", "2.1.10"),
            ("zstring", "2.6.0")
        })
        {
            var packageDirectory = Path.Combine(packageRoot, package.Item1, package.Item2);
            if (!Directory.Exists(packageDirectory))
                continue;

            foreach (var path in Directory.EnumerateFiles(packageDirectory, "*.dll", SearchOption.AllDirectories))
            {
                var fileName = Path.GetFileName(path);
                if (fileName.StartsWith("System.", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(fileName, "mscorlib.dll", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(fileName, "netstandard.dll", StringComparison.OrdinalIgnoreCase))
                    continue;

                AddPath(paths, path);
            }
        }
    }

    private static IEnumerable<ConstructorInfo> GetGovernedConstructors(Type type)
    {
        const BindingFlags flags = BindingFlags.Public | BindingFlags.NonPublic
            | BindingFlags.Instance | BindingFlags.Static | BindingFlags.DeclaredOnly;
        return type.GetConstructors(flags).Where(constructor =>
            constructor.IsPublic || constructor.IsFamily || constructor.IsFamilyOrAssembly);
    }

    private static IEnumerable<MethodInfo> GetGovernedMethods(Type type)
    {
        const BindingFlags flags = BindingFlags.Public | BindingFlags.NonPublic
            | BindingFlags.Instance | BindingFlags.Static
            | BindingFlags.DeclaredOnly;
        return type.GetMethods(flags).Where(method =>
            method.IsPublic || method.IsFamily || method.IsFamilyOrAssembly);
    }

    private static IEnumerable<PropertyInfo> GetGovernedProperties(Type type)
    {
        const BindingFlags flags = BindingFlags.Public | BindingFlags.NonPublic
            | BindingFlags.Instance | BindingFlags.Static
            | BindingFlags.DeclaredOnly;
        return type.GetProperties(flags).Where(property =>
            property.GetAccessors(true).Any(accessor =>
                accessor.IsPublic || accessor.IsFamily || accessor.IsFamilyOrAssembly));
    }

    private static IEnumerable<FieldInfo> GetGovernedFields(Type type)
    {
        const BindingFlags flags = BindingFlags.Public | BindingFlags.NonPublic
            | BindingFlags.Instance | BindingFlags.Static
            | BindingFlags.DeclaredOnly;
        return type.GetFields(flags).Where(field =>
            field.IsPublic || field.IsFamily || field.IsFamilyOrAssembly);
    }

    private static void AddDirectory(ISet<string> paths, string? directory)
    {
        if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
            return;

        foreach (var path in Directory.EnumerateFiles(directory, "*.dll"))
            AddPath(paths, path);
    }

    private static void AddPath(ISet<string> paths, string? path)
    {
        if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
            paths.Add(Path.GetFullPath(path));
    }

    private sealed class AssemblyPathResolver : MetadataAssemblyResolver
    {
        private readonly Dictionary<string, List<AssemblyCandidate>> _candidates;
        private readonly string _coreAssemblyName;
        private readonly string _coreAssemblyPath;

        public AssemblyPathResolver(IEnumerable<string> paths, IEnumerable<string> trustedPlatformPaths,
            string coreAssemblyName, string coreAssemblyPath)
        {
            _candidates = new Dictionary<string, List<AssemblyCandidate>>(StringComparer.OrdinalIgnoreCase);
            _coreAssemblyName = coreAssemblyName;
            _coreAssemblyPath = coreAssemblyPath;
            var trusted = new HashSet<string>(trustedPlatformPaths, StringComparer.OrdinalIgnoreCase);
            var identities = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var path in paths)
            {
                try
                {
                    var assemblyName = System.Reflection.AssemblyName.GetAssemblyName(path);
                    if (string.IsNullOrWhiteSpace(assemblyName.Name))
                        continue;
                    if (!identities.Add(assemblyName.FullName))
                        continue;

                    if (!_candidates.TryGetValue(assemblyName.Name, out var candidates))
                    {
                        candidates = new List<AssemblyCandidate>();
                        _candidates[assemblyName.Name] = candidates;
                    }

                    candidates.Add(new AssemblyCandidate(path, assemblyName, trusted.Contains(path)));
                }
                catch (BadImageFormatException)
                {
                }
                catch (FileLoadException)
                {
                }
            }
        }

        public override Assembly? Resolve(MetadataLoadContext context, System.Reflection.AssemblyName assemblyName)
        {
            if (string.IsNullOrWhiteSpace(assemblyName.Name))
                return null;

            var loaded = context.GetAssemblies().FirstOrDefault(candidate =>
                string.Equals(candidate.GetName().Name, assemblyName.Name, StringComparison.OrdinalIgnoreCase));
            if (loaded is not null)
                return loaded;

            if (string.Equals(assemblyName.Name, "System.Private.CoreLib", StringComparison.OrdinalIgnoreCase))
            {
                return context.LoadFromAssemblyPath(_coreAssemblyPath);
            }

            if (string.Equals(assemblyName.Name, "mscorlib", StringComparison.OrdinalIgnoreCase)
                && !_candidates.ContainsKey(assemblyName.Name))
                return context.CoreAssembly;

            if (!_candidates.TryGetValue(assemblyName.Name, out var candidates))
                return null;

            var candidate = candidates.FirstOrDefault(item =>
                    string.Equals(item.Name.FullName, assemblyName.FullName, StringComparison.OrdinalIgnoreCase))
                ?? candidates.FirstOrDefault(item => item.Name.Version == assemblyName.Version)
                ?? candidates.FirstOrDefault(item => item.IsTrustedPlatformAssembly)
                ?? candidates[0];
            return context.LoadFromAssemblyPath(candidate.Path);
        }

        private sealed class AssemblyCandidate
        {
            public AssemblyCandidate(string path, System.Reflection.AssemblyName name,
                bool isTrustedPlatformAssembly)
            {
                Path = path;
                Name = name;
                IsTrustedPlatformAssembly = isTrustedPlatformAssembly;
            }

            public string Path { get; }

            public System.Reflection.AssemblyName Name { get; }

            public bool IsTrustedPlatformAssembly { get; }
        }
    }
}