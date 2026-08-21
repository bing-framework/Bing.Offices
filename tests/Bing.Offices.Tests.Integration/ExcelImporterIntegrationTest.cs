using System;
using System.Collections.Generic;
using System.IO;
using Bing.Offices.Attributes;
using Bing.Offices.Conversions;
using Bing.Offices.Configurations;
using Bing.Offices.Csv;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Extensions;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Validations;
using Microsoft.Extensions.DependencyInjection;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests.Integration;

/// <summary>
/// Excel 导入器集成测试。
/// </summary>
public class ExcelImporterIntegrationTest
{
    /// <summary>
    /// 测试 - 通过 DI 解析的导入器应在真实 XLSX 流中执行调用方注册的自定义校验规则。
    /// </summary>
    [Fact]
    public void AddNpoi_RegisteredCustomValidationRule_ShouldValidateRealWorkbook()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddNpoi();
        services.AddSingleton<IExcelValidationRule, StartsWithExcelValidationRule>();
        using var provider = services.BuildServiceProvider();
        using var source = new MemoryStream(CreateWorkbook());

        // Act
        var result = provider.GetRequiredService<IExcelImporter>().Import(source,
            CreateImportRequest<IntegrationRow>());

        // Assert
        Assert.Empty(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal(nameof(IntegrationRow.Code), error.PropertyName);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 测试 - 通过 DI 注册的双向转换器应参与真实 XLSX 的导出与导入。
    /// </summary>
    [Fact]
    public void AddNpoi_RegisteredCustomValueConverter_ShouldRoundTripRealWorkbook()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddNpoi();
        services.AddSingleton<IExcelValueConverter, IntegrationCodeExcelValueConverter>();
        using var provider = services.BuildServiceProvider();
        using var destination = new MemoryStream();

        // Act
        provider.GetRequiredService<IExcelExporter>().Export(ExcelExport.Workbook(builder => builder.AddSheet("Data",
            new[] { new ConvertedIntegrationRow { Code = new IntegrationCode("42") } })), destination);
        destination.Position = 0;
        var result = provider.GetRequiredService<IExcelImporter>().Import(destination,
            CreateImportRequest<ConvertedIntegrationRow>());

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal("42", Assert.Single(result.Workbook.Rows).Code.Value);
        Assert.True(destination.CanRead);
        Assert.True(destination.CanWrite);
    }

    /// <summary>
    /// 测试 - AddNpoi 注册的 CSV 实体服务应通过真实流完成往返且不关闭调用方流。
    /// </summary>
    [Fact]
    public void AddNpoi_CsvServices_ShouldRoundTripStream()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddNpoi();
        using var provider = services.BuildServiceProvider();
        using var destination = new MemoryStream();

        // Act
        provider.GetRequiredService<ICsvExporter>().Export(new[]
        {
            new CsvIntegrationRow { Name = "A,\"B\"", Count = 7 }
        }, destination);
        destination.Position = 0;
        var result = provider.GetRequiredService<ICsvImporter>().Import<CsvIntegrationRow>(destination);

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal("A,\"B\"", Assert.Single(result.Items).Name);
        Assert.Equal(7, result.Items[0].Count);
        Assert.True(destination.CanRead);
    }

    /// <summary>
    /// 测试 - 通过 DI 注册的命名校验规则应由 JSON 映射配置在真实 XLSX 导入中解析。
    /// </summary>
    [Fact]
    public void AddNpoi_RegisteredNamedValidationRule_ShouldValidateConfiguredWorkbook()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddNpoi();
        services.AddSingleton<INamedExcelValidationRule, NamedStartsWithValidationRule>();
        using var provider = services.BuildServiceProvider();
        using var source = new MemoryStream(CreateWorkbook());
        var configuration = ExcelMappingConfigurationLoader.FromJson(
            "{\"columns\":[{\"propertyName\":\"Code\",\"validationRuleNames\":[\"starts-with-ok\"]}]}");

        // Act
        var result = provider.GetRequiredService<IExcelImporter>().Import(source,
            CreateImportRequest<ConfiguredIntegrationRow>(sheet => sheet.Mapping(configuration)));

        // Assert
        Assert.Empty(result.Workbook.Rows);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal(nameof(ConfiguredIntegrationRow.Code), error.PropertyName);
    }

    /// <summary>
    /// 测试 - 真实 XLS 与 XLSX 文件均应通过路径适配完成实体往返。
    /// </summary>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void StreamExtensions_RealExcelFormats_ShouldRoundTripThroughPath(ExcelFormat format)
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddNpoi();
        using var provider = services.BuildServiceProvider();
        var extension = format == ExcelFormat.Xls ? "xls" : "xlsx";
        var path = Path.Combine(Path.GetTempPath(), $"Bing.Offices.Integration.{Guid.NewGuid():N}.{extension}");

        try
        {
            // Act
            var exporter = provider.GetRequiredService<IExcelExporter>();
            var importer = provider.GetRequiredService<IExcelImporter>();
            using (var destination = new FileStream(path, FileMode.Create, FileAccess.Write, FileShare.None))
                exporter.Export(ExcelExport.Workbook(builder => builder.Format(format).AddSheet("Data",
                    new[] { new FileIntegrationRow { Name = "路径", Count = 7 } })), destination);
            using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var result = importer.Import(source, CreateImportRequest<FileIntegrationRow>());

            // Assert
            Assert.Empty(result.Errors);
            var item = Assert.Single(result.Workbook.Rows);
            Assert.Equal("路径", item.Name);
            Assert.Equal(7, item.Count);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 测试 - 已取消的真实导出请求不应向调用方目标流写入内容。
    /// </summary>
    [Fact]
    public void AddNpoi_PreCancelledExport_ShouldNotWriteDestination()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddNpoi();
        using var provider = services.BuildServiceProvider();
        using var destination = new MemoryStream();
        using var cancellationTokenSource = new System.Threading.CancellationTokenSource();
        cancellationTokenSource.Cancel();

        // Act
        var action = () => provider.GetRequiredService<IExcelExporter>().Export(
            ExcelExport.Workbook(builder => builder.AddSheet("Data",
                new[] { new FileIntegrationRow { Name = "取消", Count = 1 } })), destination,
            cancellationTokenSource.Token);

        // Assert
        Assert.Throws<OperationCanceledException>(action);
        Assert.Equal(0, destination.Length);
    }

    /// <summary>
    /// 测试 - 通过 DI 解析的导入器应支持不可寻址 XLSX 流、保持调用方流打开，并在预取消时不读取输入。
    /// </summary>
    [Fact]
    public void AddNpoi_NonSeekableAndPreCancelledImport_ShouldPreserveStreamContract()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddNpoi();
        using var provider = services.BuildServiceProvider();
        using var generatedWorkbook = new MemoryStream();
        provider.GetRequiredService<IExcelExporter>().Export(ExcelExport.Workbook(builder => builder.AddSheet("Data",
            new[] { new FileIntegrationRow { Name = "不可寻址", Count = 7 } })), generatedWorkbook);
        using var source = new NonSeekableReadStream(generatedWorkbook.ToArray());
        using var cancelledSource = new NonSeekableReadStream(generatedWorkbook.ToArray());
        using var cancellationTokenSource = new System.Threading.CancellationTokenSource();
        cancellationTokenSource.Cancel();

        // Act
        var result = provider.GetRequiredService<IExcelImporter>().Import(source,
            CreateImportRequest<FileIntegrationRow>());
        var cancelled = () => provider.GetRequiredService<IExcelImporter>().Import(cancelledSource,
            CreateImportRequest<FileIntegrationRow>(), cancellationTokenSource.Token);

        // Assert
        var item = Assert.Single(result.Workbook.Rows);
        Assert.Empty(result.Errors);
        Assert.Equal("不可寻址", item.Name);
        Assert.Equal(7, item.Count);
        Assert.True(source.CanRead);
        Assert.Throws<OperationCanceledException>(cancelled);
        Assert.True(cancelledSource.CanRead);
    }

    /// <summary>
    /// 测试 - Fluent Profile 与 XML 映射配置应在真实 XLSX 导入导出中使用相同的列定义。
    /// </summary>
    [Fact]
    public void AddNpoi_FluentProfileAndXmlConfiguration_ShouldRoundTripRealWorkbook()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddNpoi();
        using var provider = services.BuildServiceProvider();
        using var destination = new MemoryStream();
        var profile = ExcelMapping.For<MappingIntegrationRow>()
            .Property(row => row.Code).HasTitle("业务编码").And()
            .BuildProfile();
        var xmlConfiguration = ExcelMappingConfigurationLoader.FromXml(
            "<ExcelMappingConfiguration><Columns><ExcelColumnConfiguration><PropertyName>Code</PropertyName><Title>业务编码</Title></ExcelColumnConfiguration></Columns></ExcelMappingConfiguration>");

        // Act
        provider.GetRequiredService<IExcelExporter>().Export(ExcelExport.Workbook(builder => builder.AddSheet("Data",
            new[] { new MappingIntegrationRow { Code = "配置" } }, sheet => sheet.Mapping(profile))), destination);
        destination.Position = 0;
        var result = provider.GetRequiredService<IExcelImporter>().Import(destination,
            CreateImportRequest<MappingIntegrationRow>(sheet => sheet.Mapping(xmlConfiguration)));

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal("配置", Assert.Single(result.Workbook.Rows).Code);
        Assert.True(destination.CanRead);
    }

    /// <summary>
    /// 测试 - AddNpoi 应解析安全配置加载器，并将已注册的双向转换器注入 CSV 实体服务。
    /// </summary>
    [Fact]
    public void AddNpoi_ConfigurationLoaderAndCsvConverter_ShouldResolveThroughDi()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddNpoi();
        services.AddSingleton<IExcelValueConverter, IntegrationCodeExcelValueConverter>();
        using var provider = services.BuildServiceProvider();
        using var destination = new MemoryStream();

        // Act
        var configuration = provider.GetRequiredService<IExcelMappingConfigurationLoader>().FromJson(
            "{\"columns\":[{\"propertyName\":\"Code\",\"converterName\":\"integration-code\"}]}");
        provider.GetRequiredService<ICsvExporter>().Export(new[]
        {
            new ConvertedIntegrationRow { Code = new IntegrationCode("42") }
        }, destination, new CsvExportOptions<ConvertedIntegrationRow> { MappingConfiguration = configuration });
        destination.Position = 0;
        var result = provider.GetRequiredService<ICsvImporter>().Import<ConvertedIntegrationRow>(destination,
            new CsvImportOptions<ConvertedIntegrationRow> { MappingConfiguration = configuration });

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal("42", Assert.Single(result.Items).Code.Value);
        Assert.True(destination.CanRead);
    }

    /// <summary>
    /// 创建集成测试工作簿字节。
    /// </summary>
    private static ExcelWorkbookImportRequest<IntegrationWorkbook<T>> CreateImportRequest<T>(
        Action<ExcelSheetImportBuilder<T>> configure = null) where T : class, new()
    {
        return ExcelImport.Workbook<IntegrationWorkbook<T>>(builder =>
            builder.Sheet("Data", root => root.Rows, configure));
    }

    private sealed class IntegrationWorkbook<T> where T : class, new()
    {
        public List<T> Rows { get; } = new();
    }

    private static byte[] CreateWorkbook()
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(IntegrationRow.Code));
        sheet.CreateRow(1).CreateCell(0).SetCellValue("invalid");
        using var destination = new MemoryStream();
        workbook.Write(destination, false);
        return destination.ToArray();
    }

    /// <summary>
    /// 集成测试行模型。
    /// </summary>
    private class IntegrationRow
    {
        /// <summary>
        /// 业务编码。
        /// </summary>
        [StartsWithAttribute]
        public string Code { get; set; }
    }

    /// <summary>
    /// 包含领域编码的集成测试行模型。
    /// </summary>
    private class ConvertedIntegrationRow
    {
        /// <summary>
        /// 业务编码。
        /// </summary>
        public IntegrationCode Code { get; set; }
    }

    /// <summary>
    /// CSV 集成测试行模型。
    /// </summary>
    private class CsvIntegrationRow
    {
        /// <summary>
        /// 名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 数量。
        /// </summary>
        public int Count { get; set; }
    }

    /// <summary>
    /// 命名规则集成测试行模型。
    /// </summary>
    private class ConfiguredIntegrationRow
    {
        /// <summary>
        /// 业务编码。
        /// </summary>
        public string Code { get; set; }
    }

    /// <summary>
    /// 文件路径集成测试行模型。
    /// </summary>
    private class FileIntegrationRow
    {
        /// <summary>
        /// 名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 数量。
        /// </summary>
        public int Count { get; set; }
    }

    /// <summary>
    /// 映射配置集成测试行模型。
    /// </summary>
    private class MappingIntegrationRow
    {
        /// <summary>
        /// 业务编码。
        /// </summary>
        public string Code { get; set; }
    }

    /// <summary>
    /// 集成测试领域编码。
    /// </summary>
    private sealed class IntegrationCode
    {
        /// <summary>
        /// 初始化一个<see cref="IntegrationCode"/>类型的实例。
        /// </summary>
        /// <param name="value">编码文本。</param>
        public IntegrationCode(string value) => Value = value;

        /// <summary>
        /// 获取编码文本。
        /// </summary>
        public string Value { get; }
    }

    /// <summary>
    /// 前缀校验特性。
    /// </summary>
    [BindFilter(typeof(StartsWithExcelValidationRule))]
    private sealed class StartsWithAttribute : FilterAttributeBase
    {
        /// <inheritdoc />
        public override string ErrorMsg { get; set; } = "必须以 OK- 开头";
    }

    /// <summary>
    /// 前缀校验规则。
    /// </summary>
    private sealed class StartsWithExcelValidationRule : IExcelValidationRule
    {
        /// <inheritdoc />
        public bool CanValidate(FilterAttributeBase attribute) => attribute is StartsWithAttribute;

        /// <inheritdoc />
        public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context) =>
            context.Value.StartsWith("OK-", StringComparison.Ordinal);
    }

    /// <summary>
    /// 集成测试双向编码转换器。
    /// </summary>
    private sealed class IntegrationCodeExcelValueConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "integration-code";

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(IntegrationCode);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = new IntegrationCode(((string)context.Value).Substring(3));
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = $"CV-{((IntegrationCode)context.Value).Value}";
            return true;
        }
    }

    /// <summary>
    /// 通过配置名称解析的前缀校验规则。
    /// </summary>
    private sealed class NamedStartsWithValidationRule : INamedExcelValidationRule
    {
        /// <inheritdoc />
        public string Name => "starts-with-ok";

        /// <inheritdoc />
        public string ErrorMessage => "必须以 OK- 开头";

        /// <inheritdoc />
        public bool Validate(ExcelValidationContext context) =>
            context.Value.StartsWith("OK-", StringComparison.Ordinal);
    }

    /// <summary>
    /// 不支持定位的只读内存流。
    /// </summary>
    private sealed class NonSeekableReadStream : Stream
    {
        /// <summary>
        /// 被代理的内存流。
        /// </summary>
        private readonly MemoryStream _inner;

        /// <summary>
        /// 初始化一个<see cref="NonSeekableReadStream"/>类型的实例。
        /// </summary>
        /// <param name="content">输入内容。</param>
        public NonSeekableReadStream(byte[] content) => _inner = new MemoryStream(content);

        /// <inheritdoc />
        public override bool CanRead => true;

        /// <inheritdoc />
        public override bool CanSeek => false;

        /// <inheritdoc />
        public override bool CanWrite => false;

        /// <inheritdoc />
        public override long Length => throw new NotSupportedException();

        /// <inheritdoc />
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        /// <inheritdoc />
        public override void Flush()
        {
        }

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
