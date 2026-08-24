using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Bing.Offices.Attributes;
using Bing.Offices.Csv;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.AspNetCore.Http;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.Emit;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Linq;
using Xunit;

namespace Bing.Offices.Docs.Tests;

/// <summary>
/// 验证文档消费者只能通过 provider-neutral 接口使用 NPOI 注册入口。
/// </summary>
public sealed class DocsConsumerTest
{
    /// <summary>
    /// 测试 - 外部消费者应能调用 AddNpoi 并解析导入导出接口。
    /// </summary>
    [Fact]
    public void AddNpoi_ExternalConsumer_ShouldResolveProviderNeutralServices()
    {
        // Arrange
        var services = new ServiceCollection();

        // Act
        services.AddNpoi();
        using var provider = services.BuildServiceProvider();

        // Assert
        Assert.NotNull(provider.GetRequiredService<IExcelImporter>());
        Assert.NotNull(provider.GetRequiredService<IExcelExporter>());
    }

    /// <summary>
    /// 测试 - 文档请求应只构造 provider-neutral Workbook API。
    /// </summary>
    [Fact]
    public void WorkbookRequest_ExternalConsumer_ShouldBuildWithoutNpoiTypes()
    {
        // Act
        var request = ExcelImport.Workbook<DocsWorkbook>(builder =>
            builder.Sheet("Data", workbook => workbook.Rows));
        var export = ExcelExport.Workbook(builder =>
            builder.AddSheet("Data", new[] { new DocsRow { Name = "consumer" } }));

        // Assert
        Assert.Equal(1, request.SheetCount);
        Assert.Equal(1, export.SheetCount);
    }

    /// <summary>
    /// 测试 - 外部消费者应能使用 v2 文档和双模型 Profile 构建方向隔离的请求与 Registry。
    /// </summary>
    [Fact]
    public void MappingV2_ExternalConsumer_ShouldBuildDirectionalRequests()
    {
        // Arrange
        const string json = "{\"version\":2,\"profile\":\"docs\",\"modelAlias\":\"docs-row\",\"import\":{\"columns\":[{\"propertyName\":\"Name\",\"title\":\"输入名称\"}]},\"export\":{\"columns\":[{\"propertyName\":\"Label\",\"title\":\"输出标签\"}]}}";
        var document = ExcelMappingConfigurationLoader.FromJsonDocument(json);
        var services = new ServiceCollection();
        services.AddSingleton<DocsProfile>();
        services.AddMappingProfile<DocsProfile, DocsImportRow, DocsExportRow>("docs");

        // Act
        var request = ExcelImport.Workbook<DocsWorkbook>(builder =>
            builder.Sheet("Data", workbook => workbook.Rows, sheet => sheet.Mapping(document)));
        var export = ExcelExport.Workbook(builder =>
            builder.AddSheet("Data", new[] { new DocsExportRow { Label = "consumer" } }, sheet => sheet.Mapping(document)));
        using var provider = services.BuildServiceProvider();
        var registry = provider.GetRequiredService<IMappingProfileRegistry>();
        var profile = registry.Get<DocsImportRow, DocsExportRow>("docs");

        // Assert
        Assert.Equal("输入名称", Assert.Single(document.Import.Columns).Title);
        Assert.Equal("输出标签", Assert.Single(document.Export.Columns).Title);
        Assert.Equal(1, request.SheetCount);
        Assert.Equal(1, export.SheetCount);
        Assert.Equal("导入名称", Assert.Single(profile.ImportConfiguration.Columns).Title);
        Assert.Equal("导出标签", Assert.Single(profile.ExportConfiguration.Columns).Title);
    }

    /// <summary>
    /// 测试 - 文档消费者应能读取 v1/v2 JSON/XML，且加载器不关闭调用方流。
    /// </summary>
    [Fact]
    public void MappingDocuments_ExternalConsumer_ShouldMigrateAndPreserveStreams()
    {
        // Arrange
        using var jsonStream = new MemoryStream(Encoding.UTF8.GetBytes("{\"columns\":[]}"));
        using var xmlStream = new MemoryStream(Encoding.UTF8.GetBytes(
            "<ExcelMappingDocument><Version>2</Version><Import><Columns /></Import><Export><Columns /></Export></ExcelMappingDocument>"));

        // Act
        var json = ExcelMappingConfigurationLoader.FromJsonDocument(jsonStream);
        _ = ExcelMappingConfigurationLoader.FromJsonDocument("{\"columns\":[]}", out var diagnostics);
        var xml = ExcelMappingConfigurationLoader.FromXmlDocument(xmlStream);

        // Assert
        Assert.Equal(2, json.Version);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Code == "V1_MIGRATED");
        Assert.Equal(2, xml.Version);
        Assert.True(jsonStream.CanRead);
        Assert.True(xmlStream.CanRead);
    }

    /// <summary>
    /// 测试 - 文档消费者应能通过统一 CSV 主链执行 v2 校验和动态列。
    /// </summary>
    [Fact]
    public void ValidationAndDynamicColumns_ExternalConsumer_ShouldRunThroughPackageApi()
    {
        // Arrange
        using var validationSource = new MemoryStream(Encoding.UTF8.GetBytes("Code\r\nBAD\r\n"));
        using var dynamicSource = new MemoryStream(Encoding.UTF8.GetBytes("Name,区域\r\nconsumer,华东\r\n"));

        // Act
        var validation = new CsvEntityImporter().Import<DocsValidatedRow>(validationSource);
        var dynamic = new CsvEntityImporter().Import<DocsDynamicRow>(dynamicSource);

        // Assert
        Assert.Single(validation.Errors);
        Assert.Equal("consumer", Assert.Single(dynamic.Items).Name);
        Assert.Equal("华东", Assert.Single(dynamic.Items).Values["区域"]);
    }

    /// <summary>
    /// 测试 - ASP.NET Core 上传示例应只把 IFormFile 流交给 provider-neutral API。
    /// </summary>
    [Fact]
    public async System.Threading.Tasks.Task AspNetCoreUpload_ExternalConsumer_ShouldExecuteSuccessAndFailureResponses()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddLogging();
        services.AddNpoi();
        using var provider = services.BuildServiceProvider();
        var exporter = provider.GetRequiredService<IExcelExporter>();
        var importer = provider.GetRequiredService<IExcelImporter>();
        using var payload = new MemoryStream();
        DocsExamples.ReadmeExport(exporter, payload, new[]
        {
            new DocsExamples.OrderExport { DisplayName = "consumer" }
        });
        payload.Position = 0;
        IFormFile valid = new FormFile(payload, 0, payload.Length, "file", "upload.xlsx");
        using var invalidPayload = new MemoryStream(Encoding.UTF8.GetBytes("not-xlsx"));
        IFormFile invalid = new FormFile(invalidPayload, 0, invalidPayload.Length, "file", "upload.xlsx");
        var successContext = CreateHttpContext(provider);
        var failureContext = CreateHttpContext(provider);

        // Act
        await DocsExamples.Upload(valid, importer).ExecuteAsync(successContext);
        await DocsExamples.Upload(invalid, importer).ExecuteAsync(failureContext);

        // Assert
        Assert.Equal(StatusCodes.Status200OK, successContext.Response.StatusCode);
        Assert.Equal(StatusCodes.Status400BadRequest, failureContext.Response.StatusCode);
    }

    /// <summary>
    /// 测试 - 文档中的 Profile、JSON/XML、校验与动态列代码应由消费程序集真实执行。
    /// </summary>
    [Fact]
    public void DocumentationFences_ExternalConsumer_ShouldExecutePackagePaths()
    {
        // Arrange
        const string json = "{\"version\":2,\"import\":{\"columns\":[]},\"export\":{\"columns\":[]}}";
        using var validation = new MemoryStream(Encoding.UTF8.GetBytes("Code,CreatedAt\r\nBAD,2026-08-23\r\n"));
        using var dynamic = new MemoryStream(Encoding.UTF8.GetBytes("Name,区域\r\nconsumer,华东\r\n"));

        // Act
        var profile = DocsExamples.Profile();
        var migrated = DocsExamples.JsonXml(json);
        var validationResult = DocsExamples.Validation(validation);
        var dynamicResult = DocsExamples.Dynamic(dynamic);

        // Assert
        Assert.NotNull(profile);
        Assert.Equal(2, migrated.Version);
        Assert.Single(validationResult.Errors);
        Assert.Equal("华东", Assert.Single(dynamicResult.Items).Values["区域"]);
    }

    /// <summary>
    /// 测试 - docs/excel Markdown 中的每个 C# fence 都应从原文提取、编译并执行。
    /// </summary>
    [Fact]
    public void DocumentationFences_FromMarkdown_ShouldCompileAndExecuteIndividually()
    {
        // Arrange
        var documents = new[] { "README.md", "mapping-profile.md", "mapping-json-xml.md",
            "import-validation.md", "dynamic-columns.md" };
        var fences = System.Linq.Enumerable.SelectMany(documents, document => ExtractFences(document)).ToArray();

        // Act / Assert
        Assert.Equal(9, fences.Length);
        foreach (var fence in fences)
        {
            var source = BuildFenceSource(fence.FileName, fence.Index, fence.Code);
            var compilation = CSharpCompilation.Create($"DocsFence_{fence.Index}",
                new[] { CSharpSyntaxTree.ParseText(source,
                    new CSharpParseOptions(LanguageVersion.Latest)) },
                GetCompilationReferences(), new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary));
            using var assemblyStream = new MemoryStream();
            EmitResult emit = compilation.Emit(assemblyStream);
            var errors = string.Join(Environment.NewLine, System.Linq.Enumerable.Where(emit.Diagnostics, diagnostic =>
                diagnostic.Severity == DiagnosticSeverity.Error));
            Assert.True(emit.Success, $"{fence.FileName} fence {fence.Index} 编译失败:{Environment.NewLine}{errors}");
            assemblyStream.Position = 0;
            var assembly = Assembly.Load(assemblyStream.ToArray());
            assembly.GetType("FenceConsumer.FenceEntry")!.GetMethod("Run")!.Invoke(null, null);
        }
    }

    private static IEnumerable<MarkdownFence> ExtractFences(string fileName)
    {
        var path = Path.Combine(AppContext.BaseDirectory, "Docs", fileName);
        var text = File.ReadAllText(path, Encoding.UTF8);
        var matches = Regex.Matches(text, "[ \\t]*```csharp\\s*\\r?\\n(?<code>.*?)\\r?\\n[ \\t]*```",
            RegexOptions.Singleline | RegexOptions.IgnoreCase);
        for (var index = 0; index < matches.Count; index++)
            yield return new MarkdownFence(fileName, index + 1, matches[index].Groups["code"].Value);
    }

    private static IEnumerable<MetadataReference> GetCompilationReferences()
    {
        var paths = ((string)AppContext.GetData("TRUSTED_PLATFORM_ASSEMBLIES") ?? string.Empty)
            .Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries)
            .Concat(System.Linq.Enumerable.Select(AppDomain.CurrentDomain.GetAssemblies(),
                assembly => assembly.Location))
            .Where(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path))
            .Distinct(StringComparer.OrdinalIgnoreCase);
        return System.Linq.Enumerable.Select(paths, path => MetadataReference.CreateFromFile(path));
    }

    private static string BuildFenceSource(string fileName, int index, string code)
    {
        var declarations = string.Empty;
        var statements = code;
        if (code.Contains("public IResult Upload", StringComparison.Ordinal))
        {
            declarations = $"public sealed class UploadFence {{ {code} }}";
            statements = "_ = new UploadFence().Upload(file, importer);";
        }
        else if (code.Contains("public sealed class OrderProfile", StringComparison.Ordinal))
        {
            var separator = code.IndexOf("var profile", StringComparison.Ordinal);
            declarations = separator < 0 ? code : code.Substring(0, separator);
            statements = separator < 0 ? "_ = new OrderProfile();" : code.Substring(separator);
        }
        else if (Regex.IsMatch(code, "public\\s+sealed\\s+class\\s+OrderRow\\b"))
        {
            declarations = code;
            statements = "_ = new OrderRow();";
        }

        var prelude = string.Empty;
        if (fileName == "README.md")
            prelude = "var orders = new List<OrderExport> { new OrderExport { DisplayName = \"fence\" } }; "
                + "var services = new ServiceCollection(); services.AddNpoi(); using var provider = services.BuildServiceProvider(); "
                + "var exporter = provider.GetRequiredService<IExcelExporter>(); using var stream = new MemoryStream(); ";
        else if (fileName == "mapping-json-xml.md")
            prelude = "var json = \"{\\\"version\\\":2,\\\"import\\\":{\\\"columns\\\":[]},\\\"export\\\":{\\\"columns\\\":[]}}\"; ";
        else if (fileName == "dynamic-columns.md")
            prelude = "var document = new ExcelMappingDocument(); ";
        else if (fileName == "mapping-profile.md" && index == 2)
            prelude = "var services = new ServiceCollection(); ";
        else if (fileName == "import-validation.md" && index == 2)
            prelude = "var uploadServices = new ServiceCollection(); uploadServices.AddNpoi(); "
                + "using var uploadProvider = uploadServices.BuildServiceProvider(); "
                + "var exporter = uploadProvider.GetRequiredService<IExcelExporter>(); "
                + "using var uploadPayload = new MemoryStream(); "
                + "exporter.Export(ExcelExport.Workbook(builder => builder.AddSheet(\"Data\", "
                + "new[] { new OrderRow { Code = \"ORD-1\" } })), uploadPayload); uploadPayload.Position = 0; "
                + "IFormFile file = new FormFile(uploadPayload, 0, uploadPayload.Length, \"file\", \"upload.xlsx\"); "
                + "var importer = uploadProvider.GetRequiredService<IExcelImporter>(); ";

        if (fileName == "dynamic-columns.md" && code.Contains("File.OpenRead", StringComparison.Ordinal))
        {
            var path = Path.Combine(Path.GetTempPath(), $"bing-offices-doc-fence-{Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "Name,区域\r\nfence,华东\r\n", Encoding.UTF8);
            statements = statements.Replace("File.OpenRead(\"orders.csv\")",
                $"File.OpenRead(@\"{path}\")", StringComparison.Ordinal);
        }

        return $@"using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Csv;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.DependencyInjection;
namespace FenceConsumer;
{BuildSupportTypes(code)}
{declarations}
public static class FenceEntry
{{
    public static void Run()
    {{
        {prelude}
        {statements}
    }}
}}";
    }

    private static string BuildSupportTypes(string code)
    {
        var orderRow = code.Contains("public sealed class OrderRow", StringComparison.Ordinal)
            ? string.Empty
            : "public sealed class OrderRow { [ExcelIgnore] public string Code { get; set; } public string Name { get; set; } [DynamicColumn] public IDictionary<string, object> Values { get; set; } = new Dictionary<string, object>(); }";
        var profile = code.Contains("public sealed class OrderProfile", StringComparison.Ordinal)
            ? string.Empty
            : "public sealed class OrderProfile : IMappingProfile<OrderImport, OrderExport> { public void Configure(FluentSetting<OrderImport, OrderExport> setting) { } }";
        return $@"
public sealed class OrderImport {{ public string Code {{ get; set; }} }}
public sealed class OrderExport {{ public string DisplayName {{ get; set; }} public string Code {{ get; set; }} }}
public sealed class OrdersWorkbook {{ public List<OrderRow> Items {{ get; }} = new List<OrderRow>(); }}
public sealed class UploadWorkbook {{ public List<OrderRow> Rows {{ get; }} = new List<OrderRow>(); }}
{orderRow}
{profile}";
    }

    private sealed record MarkdownFence(string FileName, int Index, string Code);

    private static DefaultHttpContext CreateHttpContext(IServiceProvider services)
    {
        var context = new DefaultHttpContext { RequestServices = services };
        context.Response.Body = new MemoryStream();
        return context;
    }

    public sealed class DocsProfile : IMappingProfile<DocsImportRow, DocsExportRow>
    {
        public void Configure(FluentSetting<DocsImportRow, DocsExportRow> setting)
        {
            setting.Import.Property(row => row.Name).HasHeader("导入名称");
            setting.Export.Property(row => row.Label).HasHeader("导出标签");
        }
    }

    public sealed class DocsWorkbook
    {
        public List<DocsRow> Rows { get; } = new();
    }

    public sealed class DocsRow
    {
        public string Name { get; set; }
    }

    public sealed class DocsImportRow
    {
        public string Name { get; set; }
    }

    public sealed class DocsExportRow
    {
        public string Label { get; set; }
    }

    public sealed class DocsValidatedRow
    {
        [ExcelRequired]
        [ExcelRegex("^OK-")]
        public string Code { get; set; }
    }

    public sealed class DocsDynamicRow
    {
        public string Name { get; set; }

        [DynamicColumn]
        public IDictionary<string, object> Values { get; set; }
    }
}
