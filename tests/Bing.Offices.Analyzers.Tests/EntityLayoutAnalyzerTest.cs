using System;
using System.Collections.Generic;
using System.Collections.Immutable;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Analyzers;
using Bing.Offices.Entities;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Diagnostics;
using Xunit;

namespace Bing.Offices.Analyzers.Tests;

/// <summary>
/// 实体布局编译期诊断的职责测试。
/// </summary>
public sealed class EntityLayoutAnalyzerTest
{
    /// <summary>
    /// 测试 - 属性地址、重复占位和只读导入提示应分别诊断。
    /// </summary>
    [Fact]
    public async Task AttributeCells_ShouldReportInvalidDuplicateAndReadOnly()
    {
        const string source = """
            using Bing.Offices.Entities;
            public sealed class Order
            {
                [ExcelEntityCell("采购单", "XFE1")] public string Invalid { get; set; }
                [ExcelEntityCell("采购单", "B2")] public string First { get; set; }
                [ExcelEntityCell("采购单", "$b$2")] public string Second { get; set; }
                [ExcelEntityCell("采购单", "C3")] public string ExportOnly { get; } = "PO-1";
            }
            """;

        var diagnostics = await AnalyzeAsync(source);

        Assert.Equal(new[] { EntityLayoutAnalyzer.InvalidAddressId, EntityLayoutAnalyzer.DuplicateAddressId,
                EntityLayoutAnalyzer.ReadOnlyPropertyId }, diagnostics.Select(d => d.Id).OrderBy(id => id));
    }

    /// <summary>
    /// 测试 - Fluent 固定地址重复和非法地址应只在常量参数上报。
    /// </summary>
    [Fact]
    public async Task FluentCells_ShouldReportConstantAddressConflicts()
    {
        const string source = """
            using Bing.Offices.Entities;
            public sealed class Order { public string First { get; set; } public string Second { get; set; } }
            public static class Layouts
            {
                public static void Build() => ExcelEntity.Layout<Order>(builder => builder
                    .Cell("采购单", "A1", value => value.First)
                    .Cell("采购单", "$a$1", value => value.Second)
                    .Cell("采购单", "A0", value => value.Second));
            }
            """;

        var diagnostics = await AnalyzeAsync(source);

        Assert.Equal(new[] { EntityLayoutAnalyzer.InvalidAddressId, EntityLayoutAnalyzer.DuplicateAddressId },
            diagnostics.Select(d => d.Id).OrderBy(id => id));
    }

    /// <summary>
    /// 测试 - 同一列表的内联动态列 Key、标题、别名和组标识不能重名。
    /// </summary>
    [Fact]
    public async Task DynamicGroups_ShouldReportInlineNameConflicts()
    {
        const string source = """
            using System.Collections.Generic;
            using Bing.Offices.Entities;
            using Bing.Offices.Exports;
            public sealed class Order { public List<Line> Lines { get; set; } = new(); }
            public sealed class Line
            {
                public IDictionary<string, object> Goods { get; set; } = new Dictionary<string, object>();
                public IDictionary<string, object> Product { get; set; } = new Dictionary<string, object>();
            }
            public static class Layouts
            {
                public static void Build() => ExcelEntity.Layout<Order>(builder => builder.ListRegion(
                    "采购单", "A4", order => order.Lines, region => region
                    .DynamicColumnGroup("goods", line => line.Goods, new[]
                    {
                        new ExcelDynamicColumnDefinition { Key = "brand", Title = "品牌", Aliases = new[] { "旧品牌" } }
                    })
                    .DynamicColumnGroup("goods", line => line.Product, new[]
                    {
                        new ExcelDynamicColumnDefinition { Key = "weight", Title = "品牌", Aliases = new[] { "BRAND" } }
                    })));
            }
            """;

        var diagnostics = await AnalyzeAsync(source);

        Assert.Equal(3, diagnostics.Length);
        Assert.All(diagnostics, diagnostic => Assert.Equal(EntityLayoutAnalyzer.DuplicateDynamicNameId, diagnostic.Id));
    }

    /// <summary>
    /// 测试 - 静态信息不足或类型不属于实体布局时不应误报。
    /// </summary>
    [Fact]
    public async Task UnknownValuesAndOtherApis_ShouldNotReport()
    {
        const string source = """
            using Bing.Offices.Entities;
            public sealed class Order { public string First { get; set; } public string Second { get; set; } }
            public static class Layouts
            {
                public static void Build(string address) => ExcelEntity.Layout<Order>(builder => builder
                    .Cell("采购单", address, value => value.First)
                    .Cell("采购单", "A1", value => value.Second));
                public static void Cell(string sheet, string address, object value) { }
                public static void Other() { Cell("采购单", "A0", null); }
            }
            """;

        Assert.Empty(await AnalyzeAsync(source));
    }

    /// <summary>
    /// 测试 - Fluent 命名参数换序后仍应识别地址和所属工作表。
    /// </summary>
    [Fact]
    public async Task NamedCellArguments_ShouldUseBoundParameters()
    {
        const string source = """
            using Bing.Offices.Entities;
            public sealed class Order { public string First { get; set; } public string Second { get; set; } }
            public static class Layouts
            {
                public static void Build() => ExcelEntity.Layout<Order>(builder => builder
                    .Cell(property: value => value.First, address: "A1", sheetName: "采购单")
                    .Cell(address: "$a$1", property: value => value.Second, sheetName: "采购单")
                    .Cell(property: value => value.Second, sheetName: "采购单", address: "A0"));
            }
            """;

        Assert.Equal(new[] { EntityLayoutAnalyzer.InvalidAddressId, EntityLayoutAnalyzer.DuplicateAddressId },
            (await AnalyzeAsync(source)).Select(diagnostic => diagnostic.Id));
    }

    /// <summary>
    /// 测试 - 命名参数和目标类型推断的动态定义仍应检查重复名称。
    /// </summary>
    [Fact]
    public async Task NamedDynamicArgumentsAndTargetTypedNew_ShouldReportDuplicates()
    {
        const string source = """
            using System.Collections.Generic;
            using Bing.Offices.Entities;
            using Bing.Offices.Exports;
            public sealed class Order { public List<Line> Lines { get; set; } = new(); }
            public sealed class Line
            {
                public IDictionary<string, object> Goods { get; set; } = new Dictionary<string, object>();
                public IDictionary<string, object> Product { get; set; } = new Dictionary<string, object>();
            }
            public static class Layouts
            {
                public static void Build() => ExcelEntity.Layout<Order>(builder => builder.ListRegion(
                    "采购单", "A4", order => order.Lines, region => region
                    .DynamicColumnGroup(definitions: new ExcelDynamicColumnDefinition[]
                    {
                        new() { Key = "brand", Title = "品牌" },
                        new() { Key = "weight", Title = "品牌" }
                    }, values: line => line.Goods, groupKey: "goods")
                    .DynamicColumnGroup(values: line => line.Product, definitions: new[]
                    {
                        new ExcelDynamicColumnDefinition { Key = "BRAND", Title = "产品" }
                    }, groupKey: "goods")));
            }
            """;

        Assert.Equal(3, (await AnalyzeAsync(source)).Count(diagnostic =>
            diagnostic.Id == EntityLayoutAnalyzer.DuplicateDynamicNameId));
    }

    /// <summary>
    /// 测试 - 条件表达式的互斥定义不能被合并后误报。
    /// </summary>
    [Fact]
    public async Task ConditionalDynamicDefinitions_ShouldNotMixBranches()
    {
        const string source = """
            using System.Collections.Generic;
            using Bing.Offices.Entities;
            using Bing.Offices.Exports;
            public sealed class Order { public List<Line> Lines { get; set; } = new(); }
            public sealed class Line { public IDictionary<string, object> Goods { get; set; } = new Dictionary<string, object>(); }
            public static class Layouts
            {
                public static void Build(bool flag) => ExcelEntity.Layout<Order>(builder => builder.ListRegion(
                    "采购单", "A4", order => order.Lines, region => region.DynamicColumnGroup(
                    "goods", line => line.Goods, flag
                        ? new[] { new ExcelDynamicColumnDefinition { Key = "a", Title = "名称" } }
                        : new[] { new ExcelDynamicColumnDefinition { Key = "b", Title = "名称" } })));
            }
            """;

        Assert.Empty(await AnalyzeAsync(source));
    }

    /// <summary>
    /// 测试 - 派生实体的固定地址应与继承的属性一起比较。
    /// </summary>
    [Fact]
    public async Task InheritedAttributedCell_ShouldReportDuplicateAddress()
    {
        const string source = """
            using Bing.Offices.Entities;
            public class BaseOrder
            {
                [ExcelEntityCell("采购单", "B2")] public string Number { get; set; }
            }
            public sealed class PurchaseOrder : BaseOrder
            {
                [ExcelEntityCell("采购单", "$b$2")] public string Customer { get; set; }
            }
            """;

        Assert.Equal(EntityLayoutAnalyzer.DuplicateAddressId, Assert.Single(await AnalyzeAsync(source)).Id);
    }

    /// <summary>
    /// 测试 - 同一工作表中的重叠合并区域只报告后续声明。
    /// </summary>
    [Fact]
    public async Task OverlappingMerges_ShouldReportOnce()
    {
        const string source = """
            using Bing.Offices.Entities;
            public sealed class Order { public string Code { get; set; } }
            public static class Layouts
            {
                public static void Build() => ExcelEntity.Layout<Order>(builder => builder
                    .Cell("Data", "G1", item => item.Code)
                    .Merge("Data", "A1:C3")
                    .Merge(range: "$c$3:$e$5", sheetName: "Data")
                    .Merge("Other", "A1:C3"));
            }
            """;

        Assert.Equal(EntityLayoutAnalyzer.OverlappingAreaId, Assert.Single(await AnalyzeAsync(source)).Id);
    }

    /// <summary>
    /// 测试 - 有界列表与固定单元格、合并区域的冲突按声明位置报告。
    /// </summary>
    [Fact]
    public async Task BoundedList_ShouldReportKnownAreaConflicts()
    {
        const string source = """
            using System.Collections.Generic;
            using Bing.Offices.Entities;
            public sealed class Order
            {
                public string Code { get; set; }
                public List<Line> Lines { get; set; } = new();
            }
            public sealed class Line { public string Name { get; set; } }
            public static class Layouts
            {
                public static void Build() => ExcelEntity.Layout<Order>(builder => builder
                    .Cell("Data", "B4", item => item.Code)
                    .Merge("Data", "A5:C5")
                    .ListRegion("Data", "A4", item => item.Lines,
                        region => region.Header(false).End("D6"))
                    .Cell("Data", "C6", item => item.Code)
                    .Merge("Data", "C6:E7"));
            }
            """;

        Assert.Equal(3, (await AnalyzeAsync(source)).Count(diagnostic =>
            diagnostic.Id == EntityLayoutAnalyzer.OverlappingAreaId));
    }

    /// <summary>
    /// 测试 - 属性式固定单元格与内联有界列表共用布局时检查占位。
    /// </summary>
    [Fact]
    public async Task AttributeCellAndBoundedList_ShouldReportConflict()
    {
        const string source = """
            using System.Collections.Generic;
            using Bing.Offices.Entities;
            public sealed class Order
            {
                [ExcelEntityCell("Data", "B3")] public string Code { get; set; }
                public List<Line> Lines { get; set; } = new();
            }
            public sealed class Line { public string Name { get; set; } }
            public static class Layouts
            {
                public static void Build() => ExcelEntity.LayoutFromAttributes<Order>(builder => builder
                    .ListRegion("Data", "A2", item => item.Lines, region => region.End("D4")));
            }
            """;

        Assert.Equal(EntityLayoutAnalyzer.OverlappingAreaId, Assert.Single(await AnalyzeAsync(source)).Id);
    }

    /// <summary>
    /// 测试 - 显式启用属性单元格时也检查有界列表的占位。
    /// </summary>
    [Fact]
    public async Task ExplicitAttributeCells_ShouldReportConflict()
    {
        const string source = """
            using System.Collections.Generic;
            using Bing.Offices.Entities;
            public sealed class Order
            {
                [ExcelEntityCell("Data", "B3")] public string Code { get; set; }
                public List<Line> Lines { get; set; } = new();
            }
            public sealed class Line { public string Name { get; set; } }
            public static class Layouts
            {
                public static void Build() => ExcelEntity.Layout<Order>(builder => builder
                    .CellsFromAttributes()
                    .ListRegion("Data", "A2", item => item.Lines, region => region.End("D4")));
            }
            """;

        Assert.Equal(EntityLayoutAnalyzer.OverlappingAreaId, Assert.Single(await AnalyzeAsync(source)).Id);
    }

    /// <summary>
    /// 测试 - 属性式与 Fluent 固定单元格同址时补足跨来源检测。
    /// </summary>
    [Fact]
    public async Task AttributeAndFluentCell_ShouldReportSharedAddress()
    {
        const string source = """
            using Bing.Offices.Entities;
            public sealed class Order
            {
                [ExcelEntityCell("Data", "B3")] public string Code { get; set; }
                public string Name { get; set; }
            }
            public static class Layouts
            {
                public static void Build() => ExcelEntity.LayoutFromAttributes<Order>(builder => builder
                    .Cell("Data", "$b$3", item => item.Name));
            }
            """;

        Assert.Equal(EntityLayoutAnalyzer.OverlappingAreaId, Assert.Single(await AnalyzeAsync(source)).Id);
    }

    /// <summary>
    /// 测试 - 两个有界列表相交时只标记后声明的列表。
    /// </summary>
    [Fact]
    public async Task BoundedLists_ShouldReportOverlap()
    {
        const string source = """
            using System.Collections.Generic;
            using Bing.Offices.Entities;
            public sealed class Order
            {
                public List<Line> Left { get; set; } = new();
                public List<Line> Right { get; set; } = new();
            }
            public sealed class Line { public string Name { get; set; } }
            public static class Layouts
            {
                public static void Build() => ExcelEntity.Layout<Order>(builder => builder
                    .ListRegion("Data", "A2", item => item.Left, region => region.End("C4"))
                    .ListRegion("Data", "C4", item => item.Right, region => region.End("E6")));
            }
            """;

        Assert.Equal(EntityLayoutAnalyzer.OverlappingAreaId, Assert.Single(await AnalyzeAsync(source)).Id);
    }

    /// <summary>
    /// 测试 - 同一内联列表多次设置结束地址时采用最终生效的位置。
    /// </summary>
    [Fact]
    public async Task RepeatedEnd_ShouldUseLastAddress()
    {
        const string source = """
            using System.Collections.Generic;
            using Bing.Offices.Entities;
            public sealed class Order
            {
                public string Code { get; set; }
                public List<Line> Lines { get; set; } = new();
            }
            public sealed class Line { public string Name { get; set; } }
            public static class Layouts
            {
                public static void NoConflict() => ExcelEntity.Layout<Order>(builder => builder
                    .Cell("Data", "C4", item => item.Code)
                    .ListRegion("Data", "A2", item => item.Lines,
                        region => region.End("D5").End("B3")));
                public static void Conflict() => ExcelEntity.Layout<Order>(builder => builder
                    .Cell("Data", "C4", item => item.Code)
                    .ListRegion("Data", "A2", item => item.Lines,
                        region => region.End("B3").End("D5")));
            }
            """;

        Assert.Equal(EntityLayoutAnalyzer.OverlappingAreaId, Assert.Single(await AnalyzeAsync(source)).Id);
    }

    /// <summary>
    /// 测试 - 合并区域内的两个不同固定地址会映射到同一锚点。
    /// </summary>
    [Fact]
    public async Task MergedCells_ShouldReportNormalizedAnchorCollision()
    {
        const string source = """
            using Bing.Offices.Entities;
            public sealed class Order { public string A { get; set; } public string B { get; set; } }
            public static class Layouts
            {
                public static void Build() => ExcelEntity.Layout<Order>(builder => builder
                    .Cell("Data", "A1", item => item.A)
                    .Cell("Data", "B1", item => item.B)
                    .Merge("Data", "A1:B1"));
            }
            """;

        Assert.Equal(EntityLayoutAnalyzer.OverlappingAreaId, Assert.Single(await AnalyzeAsync(source)).Id);
    }

    /// <summary>
    /// 测试 - 合法合并锚点、不同工作表和未限定列表范围不触发静态冲突。
    /// </summary>
    [Fact]
    public async Task ValidAndUnknownAreas_ShouldNotReport()
    {
        const string source = """
            using System.Collections.Generic;
            using Bing.Offices.Entities;
            public sealed class Order
            {
                public string Code { get; set; }
                public List<Line> Lines { get; set; } = new();
            }
            public sealed class Line { public string Name { get; set; } }
            public static class Layouts
            {
                public static void Build(string end) => ExcelEntity.Layout<Order>(builder => builder
                    .Cell("Data", "A1", item => item.Code)
                    .Merge("Data", "A1:B1")
                    .ListRegion("Other", "A2", item => item.Lines, region => region.End("B3"))
                    .ListRegion("Data", "A2", item => item.Lines, region => region.End(end)));
            }
            """;

        Assert.Empty(await AnalyzeAsync(source));
    }

    private static async Task<ImmutableArray<Diagnostic>> AnalyzeAsync(string source)
    {
        var references = ((string?)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES") ?? string.Empty)
            .Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries)
            .Append(typeof(ExcelEntity).Assembly.Location)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select(path => MetadataReference.CreateFromFile(path));
        var compilation = CSharpCompilation.Create("EntityLayoutAnalyzerCase",
            new[] { CSharpSyntaxTree.ParseText(source, new CSharpParseOptions(LanguageVersion.Latest)) },
            references, new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
        Assert.Empty(compilation.GetDiagnostics().Where(diagnostic => diagnostic.Severity == DiagnosticSeverity.Error));
        return (await compilation.WithAnalyzers(ImmutableArray.Create<DiagnosticAnalyzer>(
            new EntityLayoutAnalyzer())).GetAnalyzerDiagnosticsAsync())
            .OrderBy(diagnostic => diagnostic.Id).ThenBy(diagnostic => diagnostic.Location.SourceSpan.Start)
            .ToImmutableArray();
    }
}
