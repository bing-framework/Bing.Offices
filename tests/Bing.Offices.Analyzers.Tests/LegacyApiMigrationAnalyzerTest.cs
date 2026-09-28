using System;
using System.Collections.Immutable;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Analyzers;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Diagnostics;
using Xunit;

namespace Bing.Offices.Analyzers.Tests;

/// <summary>
/// 旧版 Excel API 迁移诊断的职责测试。
/// </summary>
public sealed class LegacyApiMigrationAnalyzerTest
{
    /// <summary>
    /// 测试旧导出服务与选项引用会提示人工迁移。
    /// </summary>
    [Fact]
    public async Task LegacyExportTypes_ShouldReportMigrationWarnings()
    {
        const string source = """
            using Bing.Offices.Exports;
            namespace Bing.Offices.Exports
            {
                public interface IExcelExportService { }
                public sealed class ExportOptions<T> { }
            }
            public sealed class Order { }
            public sealed class LegacyUsage
            {
                public IExcelExportService Service { get; set; }
                public Bing.Offices.Exports.ExportOptions<Order> Options { get; set; }
            }
            """;

        var diagnostics = await AnalyzeAsync(source);
        Assert.Equal(2, diagnostics.Length);
        Assert.All(diagnostics, diagnostic => Assert.Equal(LegacyApiMigrationAnalyzer.LegacyExportId,
            diagnostic.Id));
        Assert.All(diagnostics, diagnostic => Assert.Equal(DiagnosticSeverity.Warning,
            diagnostic.Severity));
    }

    /// <summary>
    /// 测试仅旧版 <c>HasDynamicColumn</c> 特性会触发诊断。
    /// </summary>
    [Fact]
    public async Task LegacyDynamicMarker_ShouldWarnWithoutFlaggingPreservedAttributes()
    {
        const string source = """
            using System;
            using Bing.Offices.Attributes;
            namespace Bing.Offices.Attributes
            {
                public sealed class HasDynamicColumnAttribute : Attribute { }
                public sealed class DynamicColumnAttribute : Attribute { }
                public sealed class ColumnNameAttribute : Attribute
                {
                    public ColumnNameAttribute(string name) { }
                }
            }
            [HasDynamicColumn]
            public sealed class Order
            {
                [DynamicColumn] public object Values { get; set; }
                [ColumnName("订单号")] public string Code { get; set; }
            }
            """;

        var diagnostic = Assert.Single(await AnalyzeAsync(source));
        Assert.Equal(LegacyApiMigrationAnalyzer.LegacyDynamicColumnId, diagnostic.Id);
        Assert.Equal(DiagnosticSeverity.Warning, diagnostic.Severity);
    }

    /// <summary>
    /// 测试相同简单名称的业务类型不会产生迁移提示。
    /// </summary>
    [Fact]
    public async Task UnrelatedNames_ShouldNotReport()
    {
        const string source = """
            using System;
            namespace Business
            {
                public interface IExcelExportService { }
                public sealed class ExportOptions<T> { }
                public sealed class HasDynamicColumnAttribute : Attribute { }
                [HasDynamicColumn]
                public sealed class Order
                {
                    public IExcelExportService Service { get; set; }
                    public ExportOptions<Order> Options { get; set; }
                }
            }
            """;

        Assert.Empty(await AnalyzeAsync(source));
    }

    /// <summary>
    /// 使用独立编译单元执行迁移诊断。
    /// </summary>
    private static async Task<ImmutableArray<Diagnostic>> AnalyzeAsync(string source)
    {
        var references = ((string?)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES") ?? string.Empty)
            .Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select(path => MetadataReference.CreateFromFile(path));
        var compilation = CSharpCompilation.Create("LegacyApiMigrationCase",
            new[] { CSharpSyntaxTree.ParseText(source, new CSharpParseOptions(LanguageVersion.Latest)) },
            references, new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
        Assert.Empty(compilation.GetDiagnostics().Where(diagnostic => diagnostic.Severity == DiagnosticSeverity.Error));
        return (await compilation.WithAnalyzers(ImmutableArray.Create<DiagnosticAnalyzer>(
            new LegacyApiMigrationAnalyzer())).GetAnalyzerDiagnosticsAsync())
            .OrderBy(diagnostic => diagnostic.Id).ThenBy(diagnostic => diagnostic.Location.SourceSpan.Start)
            .ToImmutableArray();
    }
}
