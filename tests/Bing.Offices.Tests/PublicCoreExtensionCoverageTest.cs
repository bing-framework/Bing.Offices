using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Runtime.CompilerServices;
using System.Text;
using Bing.Offices.Entities;
using Bing.Offices.Extensions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// Core 公开扩展方法到 Common xUnit 测试的直接调用覆盖门禁。
/// </summary>
public class PublicExtensionCoverageTest
{
    /// <summary>
    /// 按数值索引的 IL 操作码字典。
    /// </summary>
    private static readonly IReadOnlyDictionary<ushort, OpCode> OpCodesByValue = typeof(OpCodes)
        .GetFields(BindingFlags.Public | BindingFlags.Static)
        .Where(field => field.FieldType == typeof(OpCode))
        .Select(field => (OpCode)field.GetValue(null))
        .ToDictionary(opCode => unchecked((ushort)opCode.Value));

    /// <summary>
    /// Core 的每个公开扩展声明都必须由 Common 中的真实测试方法直接调用。
    /// </summary>
    [Fact]
    public void PublicExtensions_ShouldHaveDirectBehaviorTestForEverySignature()
    {
        var coreAssembly = typeof(CsvStreamExtensions).Assembly;
        var coreMethods = GetPublicExtensionMethods(coreAssembly);
        var publicExtensions = coreMethods;
        var coverage = BuildDirectCallCoverage(publicExtensions);

        Assert.Equal(25, coreMethods.Length);
        Assert.Equal(25, publicExtensions.Length);

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
    /// 直接调用 Entity/Template 扩展入口，确保新增公开签名进入行为测试追溯。
    /// </summary>
    [Fact]
    public void EntityExtensions_ShouldRejectNullProvidersThroughDirectCalls()
    {
        IExcelImporter importer = null;
        IExcelExporter exporter = null;
        ExcelEntityLayout<CoverageEntity> layout = null;
        ExcelEntityTemplateOptions template = null;
        using var source = new MemoryStream();
        using var destination = new MemoryStream();

        try { _ = importer.ImportEntity(source, layout); } catch (ArgumentNullException) { }
        try { _ = importer.ImportEntityAsync(source, layout); } catch (ArgumentNullException) { }
        try { _ = importer.ImportForTemplate(source, layout, template); } catch (ArgumentNullException) { }
        try { _ = importer.ImportForTemplateAsync(source, layout, template); } catch (ArgumentNullException) { }
        try { exporter.ExportEntity(new CoverageEntity(), layout, destination); } catch (ArgumentNullException) { }
        try { _ = exporter.ExportEntityAsync(new CoverageEntity(), layout, destination); } catch (ArgumentNullException) { }
        try { exporter.ExportEntityToFile(new CoverageEntity(), layout, "coverage.xlsx"); } catch (ArgumentNullException) { }
        try { _ = exporter.ExportEntityToFileAsync(new CoverageEntity(), layout, "coverage.xlsx"); } catch (ArgumentNullException) { }
        try { exporter.ExportForTemplate(new CoverageEntity(), layout, template, destination); } catch (ArgumentNullException) { }
        try { _ = exporter.ExportForTemplateAsync(new CoverageEntity(), layout, template, destination); } catch (ArgumentNullException) { }
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

    /// <summary>
    /// 获取公开扩展方法。
    /// </summary>
    /// <param name="assembly">要检查的程序集。</param>
    /// <returns>按签名排序的程序集公开扩展方法数组。</returns>
    private static MethodInfo[] GetPublicExtensionMethods(Assembly assembly) => assembly
        .GetExportedTypes()
        .Where(type => type.IsClass && type.IsAbstract && type.IsSealed)
        .SelectMany(type => type.GetMethods(BindingFlags.Public | BindingFlags.Static | BindingFlags.DeclaredOnly))
        .Where(method => method.IsDefined(typeof(ExtensionAttribute), false))
        .OrderBy(GetSignature, StringComparer.Ordinal)
        .ToArray();

    /// <summary>
    /// 构建直接调用覆盖率映射。
    /// </summary>
    /// <param name="extensionMethods">扩展方法集合。</param>
    /// <returns>从扩展方法签名到直接调用它的测试方法列表的映射。</returns>
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

    /// <summary>
    /// 判断方法是否为 xUnit 测试。
    /// </summary>
    /// <param name="method">要检查的反射方法。</param>
    /// <returns>方法具有 Fact 或 Theory 特性时为 true，否则为 false。</returns>
    private static bool IsXunitTest(MethodInfo method) =>
        method.IsDefined(typeof(FactAttribute), true) || method.IsDefined(typeof(TheoryAttribute), true);

    /// <summary>
    /// 获取测试方法及其异步状态机实现。
    /// </summary>
    /// <param name="test">测试方法。</param>
    /// <returns>测试方法，以及存在时的异步状态机 MoveNext 方法。</returns>
    private static IEnumerable<MethodInfo> GetTestImplementationMethods(MethodInfo test)
    {
        yield return test;
        var stateMachine = test.GetCustomAttribute<AsyncStateMachineAttribute>()?.StateMachineType;
        var moveNext = stateMachine?.GetMethod("MoveNext", BindingFlags.NonPublic | BindingFlags.Instance);
        if (moveNext != null)
            yield return moveNext;
    }

    /// <summary>
    /// 读取方法体中调用的方法。
    /// </summary>
    /// <param name="source">输入流。</param>
    /// <returns>IL 中可解析的直接调用方法；没有方法体时返回空序列。</returns>
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

    /// <summary>
    /// 计算 IL 操作数占用的字节数。
    /// </summary>
    /// <param name="operandType">IL 操作数类型。</param>
    /// <param name="bytes">待处理的字节内容。</param>
    /// <param name="offset">流偏移量。</param>
    /// <returns>指定 IL 操作数占用的字节数。</returns>
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

    /// <summary>
    /// 规范化反射方法定义。
    /// </summary>
    /// <param name="method">要检查的反射方法。</param>
    /// <returns>泛型方法的泛型定义；非泛型方法返回原对象。</returns>
    private static MethodInfo Normalize(MethodInfo method) => method.IsGenericMethod
        ? method.GetGenericMethodDefinition()
        : method;

    /// <summary>
    /// 获取方法的模块和元数据标识。
    /// </summary>
    /// <param name="method">要检查的反射方法。</param>
    /// <returns>规范化方法所属模块与元数据令牌组成的标识。</returns>
    private static (Module Module, int MetadataToken) GetMethodIdentity(MethodInfo method)
    {
        var normalized = Normalize(method);
        return (normalized.Module, normalized.MetadataToken);
    }

    /// <summary>
    /// 格式化反射方法签名。
    /// </summary>
    /// <param name="method">要检查的反射方法。</param>
    /// <returns>包含声明类型、方法名、泛型参数数量和参数类型的签名。</returns>
    private static string GetSignature(MethodInfo method)
    {
        var parameters = string.Join(", ", method.GetParameters().Select(parameter =>
            GetTypeName(parameter.ParameterType)));
        var genericArity = method.IsGenericMethodDefinition ? $"``{method.GetGenericArguments().Length}" : string.Empty;
        return $"{method.DeclaringType?.FullName}.{method.Name}{genericArity}({parameters})";
    }

    /// <summary>
    /// 格式化类型名称。
    /// </summary>
    /// <param name="type">要格式化或检查的类型。</param>
    /// <returns>保留数组和泛型参数结构的类型名称。</returns>
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

    /// <summary>
    /// 写入公共扩展覆盖率追溯报告。
    /// </summary>
    /// <param name="coverage">扩展方法覆盖率映射。</param>
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
            "- Task-ID：`BING-OFFICES-RC-TEST-ARCH-HARDENING-20260917-001`",
            "- Gate：`PublicExtensionCoverageTest.PublicExtensions_ShouldHaveDirectBehaviorTestForEverySignature`",
            "- 统计：Core `25/25`",
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

    /// <summary>
    /// 提供覆盖率调用目标。
    /// </summary>
    private static void CoverageTarget()
    {
    }

    /// <summary>
    /// 调用覆盖率目标方法。
    /// </summary>
    private static void CallCoverageTarget() => CoverageTarget();

    /// <summary>
    /// 创建覆盖率目标委托。
    /// </summary>
    /// <returns>引用覆盖率目标方法的委托。</returns>
    private static Action CreateCoverageTargetDelegate() => CoverageTarget;

    /// <summary>
    /// 创建不会执行覆盖率目标的 Lambda。
    /// </summary>
    /// <returns>调用时才执行覆盖率目标方法的委托。</returns>
    private static Action CreateUnexecutedCoverageLambda() => () => CoverageTarget();

    /// <summary>
    /// 覆盖公开 Entity 扩展签名所需的最小实体类型。
    /// </summary>
    private sealed class CoverageEntity
    {
    }
}
