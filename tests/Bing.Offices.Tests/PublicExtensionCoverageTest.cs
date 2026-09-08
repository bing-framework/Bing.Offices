using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.CompilerServices;
using System.Text;
using Bing.Offices.Extensions;
using Bing.Offices.Npoi.Extensions;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 公开扩展方法到职责级 xUnit 测试的直接调用覆盖门禁。
/// </summary>
public class PublicExtensionCoverageTest
{
    private static readonly IReadOnlyDictionary<ushort, OpCode> OpCodesByValue = typeof(OpCodes)
        .GetFields(BindingFlags.Public | BindingFlags.Static)
        .Where(field => field.FieldType == typeof(OpCode))
        .Select(field => (OpCode)field.GetValue(null))
        .ToDictionary(opCode => unchecked((ushort)opCode.Value));

    /// <summary>
    /// Core 和 NPOI 的每个公开扩展声明都必须由真实测试方法直接调用。
    /// </summary>
    [Fact]
    public void PublicExtensions_ShouldHaveDirectBehaviorTestForEverySignature()
    {
        var coreAssembly = typeof(CsvStreamExtensions).Assembly;
        var npoiAssembly = typeof(CellExtensions).Assembly;
        var coreMethods = GetPublicExtensionMethods(coreAssembly);
        var npoiMethods = GetPublicExtensionMethods(npoiAssembly);
        var publicExtensions = coreMethods.Concat(npoiMethods).ToArray();
        var coverage = BuildDirectCallCoverage(publicExtensions);

        Assert.Equal(19, coreMethods.Length);
        Assert.Equal(66, npoiMethods.Length);
        Assert.Equal(85, publicExtensions.Length);

        var missing = publicExtensions
            .Where(method => !coverage.ContainsKey(GetSignature(method)))
            .Select(GetSignature)
            .OrderBy(signature => signature, StringComparer.Ordinal)
            .ToArray();
        Assert.True(missing.Length == 0,
            "以下公开扩展方法没有被 [Fact]/[Theory] 直接调用：" + Environment.NewLine +
            string.Join(Environment.NewLine, missing));

        Assert.All(coverage, pair =>
        {
            Assert.Contains(publicExtensions, method => GetSignature(method) == pair.Key);
            Assert.All(pair.Value, test => Assert.True(IsXunitTest(test),
                $"覆盖目标不是 xUnit 测试：{test.DeclaringType?.FullName}.{test.Name}"));
        });

        WriteTraceabilityReport(coverage);
    }

    /// <summary>
    /// IL 读取器只接受真实调用，不把方法组取址或未执行 lambda 计为覆盖。
    /// </summary>
    [Fact]
    public void CoverageReader_ShouldRejectMethodPointersAndUnexecutedLambdas()
    {
        var direct = typeof(PublicExtensionCoverageTest).GetMethod(nameof(CallCoverageTarget),
            BindingFlags.NonPublic | BindingFlags.Static);
        var methodGroup = typeof(PublicExtensionCoverageTest).GetMethod(nameof(CreateCoverageTargetDelegate),
            BindingFlags.NonPublic | BindingFlags.Static);
        var lambda = typeof(PublicExtensionCoverageTest).GetMethod(nameof(CreateUnexecutedCoverageLambda),
            BindingFlags.NonPublic | BindingFlags.Static);

        Assert.Contains(GetCalledMethods(direct), method => method.Name == nameof(CoverageTarget));
        Assert.DoesNotContain(GetCalledMethods(methodGroup), method => method.Name == nameof(CoverageTarget));
        Assert.DoesNotContain(GetCalledMethods(lambda), method => method.Name == nameof(CoverageTarget));
    }

    private static MethodInfo[] GetPublicExtensionMethods(Assembly assembly) => assembly
        .GetExportedTypes()
        .Where(type => type.IsClass && type.IsAbstract && type.IsSealed)
        .SelectMany(type => type.GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly))
        .Where(method => method.IsDefined(typeof(ExtensionAttribute), false))
        .OrderBy(GetSignature, StringComparer.Ordinal)
        .ToArray();

    private static IReadOnlyDictionary<string, IReadOnlyList<MethodInfo>> BuildDirectCallCoverage(
        IReadOnlyCollection<MethodInfo> extensionMethods)
    {
        var extensionByIdentity = extensionMethods.ToDictionary(GetMethodIdentity);
        var result = new Dictionary<string, List<MethodInfo>>(StringComparer.Ordinal);
        var tests = typeof(PublicExtensionCoverageTest).Assembly
            .GetTypes()
            .SelectMany(type => type.GetMethods(BindingFlags.Public | BindingFlags.Instance |
                                                BindingFlags.Static | BindingFlags.DeclaredOnly))
            .Where(IsXunitTest);

        foreach (var test in tests)
        {
            foreach (var method in GetTestImplementationMethods(test))
            {
                foreach (var called in GetCalledMethods(method))
                {
                    var identity = GetMethodIdentity(called);
                    if (!extensionByIdentity.TryGetValue(identity, out var extension))
                        continue;
                    var signature = GetSignature(extension);
                    if (!result.TryGetValue(signature, out var targets))
                    {
                        targets = new List<MethodInfo>();
                        result.Add(signature, targets);
                    }
                    if (!targets.Contains(test))
                        targets.Add(test);
                }
            }
        }

        return result.ToDictionary(pair => pair.Key,
            pair => (IReadOnlyList<MethodInfo>)pair.Value,
            StringComparer.Ordinal);
    }

    private static bool IsXunitTest(MethodInfo method) =>
        method.IsDefined(typeof(FactAttribute), true) || method.IsDefined(typeof(TheoryAttribute), true);

    private static IEnumerable<MethodInfo> GetTestImplementationMethods(MethodInfo test)
    {
        yield return test;
        var stateMachine = test.GetCustomAttribute<AsyncStateMachineAttribute>()?.StateMachineType;
        var moveNext = stateMachine?.GetMethod("MoveNext", BindingFlags.NonPublic | BindingFlags.Instance);
        if (moveNext != null)
            yield return moveNext;
    }

    private static IEnumerable<MethodInfo> GetCalledMethods(MethodInfo source)
    {
        var body = source.GetMethodBody();
        var bytes = body?.GetILAsByteArray();
        if (bytes == null)
            yield break;

        var offset = 0;
        while (offset < bytes.Length)
        {
            var first = bytes[offset++];
            var value = first == 0xfe
                ? (ushort)(0xfe00 | bytes[offset++])
                : first;
            if (!OpCodesByValue.TryGetValue(value, out var opCode))
                throw new InvalidOperationException($"无法解析 IL 操作码 0x{value:x4}。");

            if (opCode == OpCodes.Call || opCode == OpCodes.Callvirt)
            {
                var token = BitConverter.ToInt32(bytes, offset);
                MethodBase resolved = null;
                try
                {
                    resolved = source.Module.ResolveMethod(token,
                        source.DeclaringType?.GetGenericArguments(), source.GetGenericArguments());
                }
                catch (ArgumentException)
                {
                    // 无法解析的外部泛型成员不属于当前程序集的公开扩展集合。
                }
                if (resolved is MethodInfo method)
                    yield return Normalize(method);
            }

            offset += GetOperandSize(opCode.OperandType, bytes, offset);
        }
    }

    private static int GetOperandSize(OperandType operandType, byte[] bytes, int offset) => operandType switch
    {
        OperandType.InlineNone => 0,
        OperandType.ShortInlineBrTarget or OperandType.ShortInlineI or OperandType.ShortInlineVar => 1,
        OperandType.InlineVar => 2,
        OperandType.InlineBrTarget or OperandType.InlineField or OperandType.InlineI or OperandType.InlineMethod or
            OperandType.InlineSig or OperandType.InlineString or OperandType.InlineTok or OperandType.InlineType or
            OperandType.ShortInlineR => 4,
        OperandType.InlineI8 or OperandType.InlineR => 8,
        OperandType.InlineSwitch => 4 + BitConverter.ToInt32(bytes, offset) * 4,
        _ => throw new ArgumentOutOfRangeException(nameof(operandType), operandType, null)
    };

    private static MethodInfo Normalize(MethodInfo method) => method.IsGenericMethod
        ? method.GetGenericMethodDefinition()
        : method;

    private static (Module Module, int MetadataToken) GetMethodIdentity(MethodInfo method)
    {
        var normalized = Normalize(method);
        return (normalized.Module, normalized.MetadataToken);
    }

    private static string GetSignature(MethodInfo method)
    {
        var parameters = string.Join(", ", method.GetParameters().Select(parameter =>
            GetTypeName(parameter.ParameterType)));
        var genericArity = method.IsGenericMethodDefinition ? $"``{method.GetGenericArguments().Length}" : string.Empty;
        return $"{method.DeclaringType?.FullName}.{method.Name}{genericArity}({parameters})";
    }

    private static string GetTypeName(Type type)
    {
        if (type.IsGenericParameter)
            return type.Name;
        if (type.IsArray)
            return $"{GetTypeName(type.GetElementType())}[]";
        if (!type.IsGenericType)
            return type.FullName ?? type.Name;
        var name = type.GetGenericTypeDefinition().FullName;
        var separator = name?.IndexOf('`') ?? -1;
        if (separator >= 0)
            name = name.Substring(0, separator);
        return $"{name}<{string.Join(",", type.GetGenericArguments().Select(GetTypeName))}>";
    }

    private static void WriteTraceabilityReport(
        IReadOnlyDictionary<string, IReadOnlyList<MethodInfo>> coverage)
    {
        var path = Environment.GetEnvironmentVariable("PUBLIC_EXTENSION_COVERAGE_REPORT");
        if (string.IsNullOrWhiteSpace(path))
            return;

        var lines = new List<string>
        {
            "# Public Extension Coverage Traceability",
            string.Empty,
            "- Task-ID：`BO-RC-20260907-001`",
            "- Gate：`PublicExtensionCoverageTest.PublicExtensions_ShouldHaveDirectBehaviorTestForEverySignature`",
            "- 统计：Core `19/19`，NPOI `66/66`，总计 `85/85`",
            "- 判定：`PASS`（完整签名逐项映射）",
            string.Empty,
            "门禁仅接受真实 `call`/`callvirt`，异步测试通过 `AsyncStateMachineAttribute` 精确定位 `MoveNext`；方法组取址、未执行 lambda 和同名实例方法不会计入覆盖。",
            string.Empty,
            "| 完整生产签名 | 测试项目 | `[Fact]` / `[Theory]` 方法（方法名体现关键行为） |",
            "| --- | --- | --- |"
        };
        foreach (var pair in coverage.OrderBy(item => item.Key, StringComparer.Ordinal))
        {
            var tests = string.Join("<br>", pair.Value
                .Select(test => $"`{test.DeclaringType?.Name}.{test.Name}`")
                .OrderBy(name => name, StringComparer.Ordinal));
            lines.Add($"| `{pair.Key.Replace("|", "\\|")}` | `Bing.Offices.Tests` | {tests} |");
        }

        var directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory))
            Directory.CreateDirectory(directory);
        File.WriteAllText(path, string.Join(Environment.NewLine, lines) + Environment.NewLine,
            new UTF8Encoding(false));
    }

    private static void CoverageTarget()
    {
    }

    private static void CallCoverageTarget() => CoverageTarget();

    private static Action CreateCoverageTargetDelegate() => CoverageTarget;

    private static Action CreateUnexecutedCoverageLambda() => () => CoverageTarget();
}
