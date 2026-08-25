using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Csv;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.Mappings;
using Bing.Offices.Npoi;
using Bing.Offices.Npoi.Exports;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Npoi.Imports;
using Bing.Offices.Validations;
using Microsoft.Extensions.DependencyInjection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 流式 Excel 导入导出管线测试。
/// </summary>
public class StreamPipelineTest
{
    /// <summary>
    /// 测试 - 服务注册应只解析新的流式导入和导出契约。
    /// </summary>
    [Fact]
    public void AddNpoi_ShouldRegisterStreamPipeline()
    {
        // Arrange
        var services = new ServiceCollection();

        // Act
        services.AddNpoi();
        using var provider = services.BuildServiceProvider();

        // Assert
        Assert.IsType<NpoiExcelImporter>(provider.GetRequiredService<IExcelImporter>());
        Assert.IsType<NpoiExcelExporter>(provider.GetRequiredService<IExcelExporter>());
    }

    /// <summary>
    /// 测试 - AddNpoi 应将请求处理服务注册为 transient，内置规则注册为 singleton 枚举项。
    /// </summary>
    [Fact]
    public void AddNpoi_ServiceLifetimes_ShouldPreserveReplacementAndRequestIsolation()
    {
        // Arrange
        var services = new ServiceCollection();

        // Act
        services.AddNpoi();

        // Assert
        Assert.Equal(ServiceLifetime.Transient, Assert.Single(services.Where(descriptor =>
            descriptor.ServiceType == typeof(IExcelImporter))).Lifetime);
        Assert.Equal(ServiceLifetime.Transient, Assert.Single(services.Where(descriptor =>
            descriptor.ServiceType == typeof(IExcelExporter))).Lifetime);
        Assert.Equal(ServiceLifetime.Transient, Assert.Single(services.Where(descriptor =>
            descriptor.ServiceType == typeof(ICsvImporter))).Lifetime);
        Assert.Equal(ServiceLifetime.Transient, Assert.Single(services.Where(descriptor =>
            descriptor.ServiceType == typeof(ICsvExporter))).Lifetime);
        Assert.All(services.Where(descriptor => descriptor.ServiceType == typeof(IExcelValidationRule)), descriptor =>
            Assert.Equal(ServiceLifetime.Singleton, descriptor.Lifetime));
    }

    /// <summary>
    /// 测试 - AddNpoi 应注册与直接构造相同的全部内置校验规则，包括最大值规则。
    /// </summary>
    [Fact]
    public void AddNpoi_DefaultValidationRules_ShouldMatchDirectConstruction()
    {
        // Arrange
        var services = new ServiceCollection();
        services.AddNpoi();
        using var provider = services.BuildServiceProvider();

        // Act
        var registered = provider.GetServices<IExcelValidationRule>()
            .Select(rule => rule.GetType())
            .OrderBy(type => type.FullName, StringComparer.Ordinal)
            .ToArray();
        var direct = ExcelValidationRules.CreateDefault()
            .Select(rule => rule.GetType())
            .OrderBy(type => type.FullName, StringComparer.Ordinal)
            .ToArray();

        // Assert
        Assert.Equal(direct, registered);
    }

    /// <summary>
    /// 测试 - 流复制协作者应完整复制数据且不关闭调用方拥有的源流和目标流。
    /// </summary>
    [Fact]
    public void NpoiStreamCopier_Copy_ShouldPreserveDataAndStreamOwnership()
    {
        // Arrange
        var expected = Encoding.UTF8.GetBytes("Bing.Offices stream payload");
        using var source = new MemoryStream(expected);
        using var destination = new MemoryStream();

        // Act
        NpoiStreamCopier.Copy(source, destination, CancellationToken.None);

        // Assert
        Assert.Equal(expected, destination.ToArray());
        Assert.True(source.CanRead);
        Assert.True(destination.CanRead);
        Assert.True(destination.CanWrite);
    }

    /// <summary>
    /// 测试 - 流复制超过最大输入字节数时应拒绝继续写入。
    /// </summary>
    [Fact]
    public void NpoiStreamCopier_Copy_WhenMaxBytesExceeded_ShouldThrowBeforeWrite()
    {
        // Arrange
        using var source = new MemoryStream(new byte[] { 1, 2, 3 });
        using var destination = new MemoryStream();

        // Act
        var action = () => NpoiStreamCopier.Copy(source, destination, CancellationToken.None, 2);

        // Assert
        var exception = Assert.Throws<InvalidOperationException>(action);
        Assert.Equal("输入工作簿超过最大字节数: 2", exception.Message);
        Assert.Empty(destination.ToArray());
    }

    /// <summary>
    /// 测试 - 复制开始前已取消时应立即抛出取消异常且不读取或写入流。
    /// </summary>
    [Fact]
    public void NpoiStreamCopier_Copy_WhenCancelled_ShouldThrowBeforeIo()
    {
        // Arrange
        using var source = new MemoryStream(new byte[] { 1, 2, 3 });
        using var destination = new MemoryStream();
        using var cancellationTokenSource = new CancellationTokenSource();
        cancellationTokenSource.Cancel();

        // Act
        var action = () => NpoiStreamCopier.Copy(source, destination, cancellationTokenSource.Token);

        // Assert
        Assert.Throws<OperationCanceledException>(action);
        Assert.Equal(0, source.Position);
        Assert.Empty(destination.ToArray());
    }

    /// <summary>
    /// 测试 - AddNpoi 应注册可被调用方替换的无状态映射配置加载器。
    /// </summary>
    [Fact]
    public void AddNpoi_MappingConfigurationLoader_ShouldBeResolvableAndReplaceable()
    {
        // Arrange
        var services = new ServiceCollection();
        var replacement = new TestMappingConfigurationLoader();

        // Act
        services.AddSingleton<IExcelMappingConfigurationLoader>(replacement);
        services.AddNpoi();
        using var provider = services.BuildServiceProvider();

        // Assert
        Assert.Same(replacement, provider.GetRequiredService<IExcelMappingConfigurationLoader>());
    }

    /// <summary>
    /// 测试 - 文件内容应通过调用方流和 Workbook 请求导入。
    /// </summary>
    [Fact]
    public void Import_FileStream_ShouldUseWorkbookRequest()
    {
        // Arrange
        var path = Path.Combine(Path.GetTempPath(), $"Bing.Offices.Tests.{Guid.NewGuid():N}.xlsx");
        try
        {
            File.WriteAllBytes(path, CreateWorkbook(workbook =>
            {
                var sheet = workbook.CreateSheet("Data");
                sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(StreamRow.Name));
                sheet.CreateRow(1).CreateCell(0).SetCellValue("兼容路径");
            }));

            // Act
            using var source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            var result = new NpoiExcelImporter().Import(source, CreateSingleSheetRequest<StreamRow>());

            // Assert
            Assert.Empty(result.Errors);
            Assert.Equal("兼容路径", Assert.Single(result.Workbook.Items).Name);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 测试 - 导入器应支持不可寻址输入流，保留调用方流并返回已绑定数据。
    /// </summary>
    [Fact]
    public void Import_NonSeekableStream_ShouldKeepSourceOpenAndReturnItems()
    {
        // Arrange
        using var source = new NonSeekableReadStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(StreamRow.Name));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("StreamValue");
        }));

        // Act
        var result = new NpoiExcelImporter().Import(source, CreateSingleSheetRequest<StreamRow>());

        // Assert
        Assert.True(source.CanRead);
        Assert.Empty(result.Errors);
        Assert.Equal("StreamValue", Assert.Single(result.Workbook.Items).Name);
    }

    /// <summary>
    /// 测试 - 导入器应收集转换、必填和重复校验错误，且不返回失败行。
    /// </summary>
    [Fact]
    public void Import_InvalidAndDuplicateValues_ShouldReturnStructuredErrors()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue(nameof(ValidatedRow.Code));
            header.CreateCell(1).SetCellValue(nameof(ValidatedRow.Count));
            var first = sheet.CreateRow(1);
            first.CreateCell(0).SetCellValue("Code-A");
            first.CreateCell(1).SetCellValue("1");
            var duplicate = sheet.CreateRow(2);
            duplicate.CreateCell(0).SetCellValue("code-a");
            duplicate.CreateCell(1).SetCellValue("2");
            var invalid = sheet.CreateRow(3);
            invalid.CreateCell(0).SetCellValue("Code-B");
            invalid.CreateCell(1).SetCellValue("invalid");
        }));

        // Act
        var result = new NpoiExcelImporter().Import(source, CreateSingleSheetRequest<ValidatedRow>(sheet =>
            sheet.Validate(ValidateMode.Continue)));

        // Assert
        Assert.Single(result.Workbook.Items);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.Validation
                                                && error.PropertyName == nameof(ValidatedRow.Code));
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ValueConversion
                                                && error.PropertyName == nameof(ValidatedRow.Count));
    }

    /// <summary>
    /// 测试 - 转换失败行不应提交重复校验状态，后续有效行仍可导入。
    /// </summary>
    [Fact]
    public void Import_ConversionFailure_ShouldNotPolluteDuplicateValidationState()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue(nameof(ValidatedRow.Code));
            header.CreateCell(1).SetCellValue(nameof(ValidatedRow.Count));
            var invalid = sheet.CreateRow(1);
            invalid.CreateCell(0).SetCellValue("Code-A");
            invalid.CreateCell(1).SetCellValue("invalid");
            var valid = sheet.CreateRow(2);
            valid.CreateCell(0).SetCellValue("Code-A");
            valid.CreateCell(1).SetCellValue("1");
        }));

        // Act
        var result = new NpoiExcelImporter().Import(source, CreateSingleSheetRequest<ValidatedRow>());

        // Assert
        Assert.Single(result.Workbook.Items);
        Assert.Equal("Code-A", result.Workbook.Items[0].Code);
        Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, result.Errors[0].Code);
    }

    /// <summary>
    /// 测试 - 导入器应将未知列写入唯一动态列，并遵守请求级表头映射。
    /// </summary>
    [Fact]
    public void Import_HeaderMappingAndDynamicColumns_ShouldBindValues()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("业务编码");
            header.CreateCell(1).SetCellValue("扩展字段");
            var row = sheet.CreateRow(1);
            row.CreateCell(0).SetCellValue("A-1");
            row.CreateCell(1).SetCellValue("扩展值");
        }));

        // Act
        var result = new NpoiExcelImporter().Import(source, ExcelImport.Workbook<SingleWorkbook<DynamicRow>>(builder =>
            builder.Sheet("Data", root => root.Items, sheet => sheet.Mapping(new ExcelMappingConfiguration
            {
                Columns = new List<ExcelColumnConfiguration>
                {
                    new() { PropertyName = nameof(DynamicRow.Code), Title = "业务编码" }
                }
            }))));

        // Assert
        var item = Assert.Single(result.Workbook.Items);
        Assert.Equal("A-1", item.Code);
        Assert.Equal("扩展值", item.Values["扩展字段"]);
    }

    /// <summary>
    /// 测试 - 自定义筛选特性应通过绑定的行级校验规则参与 Stream-first 导入。
    /// </summary>
    [Fact]
    public void Import_CustomValidationRule_ShouldReturnStructuredValidationError()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(CustomValidationRow.Code));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("invalid");
        }));

        // Act
        var result = new NpoiExcelImporter(new IExcelValidationRule[] { new StartsWithExcelValidationRule() })
            .Import(source, CreateSingleSheetRequest<CustomValidationRow>());

        // Assert
        Assert.Empty(result.Workbook.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal(nameof(CustomValidationRow.Code), error.PropertyName);
        Assert.Equal("必须以 OK- 开头", error.Message);
    }

    /// <summary>
    /// 测试 - 自定义校验规则抛出异常时，导入器应返回结构化校验错误而不是中断管线。
    /// </summary>
    [Fact]
    public void Import_ThrowingValidationRule_ShouldReturnStructuredValidationError()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(ThrowingValidationRow.Code));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("value");
        }));

        // Act
        var result = new NpoiExcelImporter(new IExcelValidationRule[] { new ThrowingExcelValidationRule() })
            .Import(source, CreateSingleSheetRequest<ThrowingValidationRow>());

        // Assert
        Assert.Empty(result.Workbook.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal(nameof(ThrowingValidationRow.Code), error.PropertyName);
        Assert.Equal("校验器异常", error.Message);
    }

    /// <summary>
    /// 测试 - 未绑定且未注册规则的自定义筛选特性应明确失败，不能静默通过。
    /// </summary>
    [Fact]
    public void Import_UnboundCustomValidationAttribute_ShouldThrowConfigurationException()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(UnboundValidationRow.Code));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("Value");
        }));

        // Act
        var action = () => new NpoiExcelImporter().Import(source, CreateSingleSheetRequest<UnboundValidationRow>());

        // Assert
        Assert.Throws<InvalidOperationException>(action);
    }

    /// <summary>
    /// 测试 - 自定义值转换器应在导出和导入时双向处理领域值对象。
    /// </summary>
    [Fact]
    public void StreamPipeline_CustomValueConverter_ShouldRoundTripDomainValue()
    {
        // Arrange
        var converter = new OrderCodeExcelValueConverter();
        using var destination = new MemoryStream();
        var exporter = new NpoiExcelExporter(new IExcelValueConverter[] { converter });
        var importer = new NpoiExcelImporter(valueConverters: new IExcelValueConverter[] { converter });

        // Act
        exporter.Export(CreateSingleSheetExportRequest(
            new[] { new ConvertedRow { Code = new OrderCode("42") } }), destination);
        destination.Position = 0;
        var result = importer.Import(destination, CreateSingleSheetRequest<ConvertedRow>());

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal("CV-42", GetWorkbookCellValue(destination, 1, 0));
        Assert.Equal("42", Assert.Single(result.Workbook.Items).Code.Value);
        Assert.Equal(typeof(OrderCode), converter.ImportPropertyType);
        Assert.Equal(typeof(OrderCode), converter.ExportPropertyType);
        Assert.Equal(2, converter.ImportRowIndex);
        Assert.Equal(2, converter.ExportRowIndex);
    }

    /// <summary>
    /// 测试 - Workbook 请求应支持异步测试场景中的字节缓冲往返。
    /// </summary>
    [Fact]
    public async Task StreamExtensions_LegacyExcelAsyncBytes_ShouldDelegateToExporter()
    {
        // Arrange
        var exporter = new NpoiExcelExporter();
        using var destination = new MemoryStream();

        // Act
    await Task.Run(() => exporter.Export(CreateSingleSheetExportRequest(
        new[] { new StreamRow { Name = "兼容" } }), destination));

        // Assert
        Assert.NotEmpty(destination.ToArray());
        destination.Position = 0;
        var result = new NpoiExcelImporter().Import(destination, CreateSingleSheetRequest<StreamRow>());
        Assert.Empty(result.Errors);
        Assert.Equal("兼容", Assert.Single(result.Workbook.Items).Name);
    }

    /// <summary>
    /// 测试 - Fluent 映射配置应在单次真实工作簿导入导出中覆盖标题、顺序和值映射。
    /// </summary>
    [Fact]
    public void StreamPipeline_FluentMappingConfiguration_ShouldApplyToSingleRequest()
    {
        // Arrange
        var mapping = ExcelMapping.For<ConfiguredRow>()
            .Property(row => row.Count).HasTitle("数量").HasColumnIndex(0).Map("一", 1).And()
            .Property(row => row.Name).HasTitle("名称").HasColumnIndex(1).And()
            .Build();
        using var destination = new MemoryStream();
        var exporter = new NpoiExcelExporter();
        var importer = new NpoiExcelImporter();

        // Act
        exporter.Export(CreateSingleSheetExportRequest(new[] { new ConfiguredRow { Name = "配置行", Count = 1 } },
            configure: sheet => sheet.Mapping(mapping)), destination);
        destination.Position = 0;
        var result = importer.Import(destination, CreateSingleSheetRequest<ConfiguredRow>(sheet => sheet.Mapping(mapping)));

        // Assert
        Assert.Empty(result.Errors);
        Assert.True(destination.CanRead);
        Assert.Equal("数量", GetWorkbookCellValue(destination, 0, 0));
        Assert.Equal("一", GetWorkbookCellValue(destination, 1, 0));
        Assert.Equal("配置行", Assert.Single(result.Workbook.Items).Name);
        Assert.Equal(1, result.Workbook.Items[0].Count);
    }

    /// <summary>
    /// 测试 - 请求映射应拒绝与默认列冲突的标题和重复显式列索引。
    /// </summary>
    [Fact]
    public void TypeMap_ConflictingTitlesOrIndexes_ShouldThrow()
    {
        // Arrange
        var duplicateTitle = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new ExcelColumnConfiguration { PropertyName = nameof(ConfiguredRow.Name), Title = nameof(ConfiguredRow.Count) }
            }
        };
        var duplicateIndex = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new ExcelColumnConfiguration { PropertyName = nameof(ConfiguredRow.Name), ColumnIndex = 0 },
                new ExcelColumnConfiguration { PropertyName = nameof(ConfiguredRow.Count), ColumnIndex = 0 }
            }
        };

        // Act
        var titleAction = () => ExcelTypeMapFactory.Get<ConfiguredRow>(duplicateTitle);
        var indexAction = () => ExcelTypeMapFactory.Get<ConfiguredRow>(duplicateIndex);

        // Assert
        Assert.Throws<ArgumentException>(titleAction);
        Assert.Throws<ArgumentException>(indexAction);
    }

    /// <summary>
    /// 测试 - 请求配置应覆盖 Profile，Profile 应覆盖实体特性映射。
    /// </summary>
    [Fact]
    public void TypeMap_RequestConfiguration_ShouldOverrideProfile()
    {
        // Arrange
        var profile = ExcelMapping.For<ConfiguredRow>()
            .Property(row => row.Name).HasTitle("Profile 名称").And()
            .Build();
        var request = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new ExcelColumnConfiguration { PropertyName = nameof(ConfiguredRow.Name), Title = "请求名称" }
            }
        };

        // Act
        var map = ExcelTypeMapFactory.Get<ConfiguredRow>(
            MappingConfigurationMerger.Merge(profile, request, MappingSourceKind.Request));

        // Assert
        Assert.Equal("请求名称", map.Properties.Single(property => property.Name == nameof(ConfiguredRow.Name)).Title);
    }

    /// <summary>
    /// 测试 - Fluent Profile 应保存独立快照，外部配置或读取结果的修改不得影响后续映射。
    /// </summary>
    [Fact]
    public void TypeMap_Profile_ShouldKeepIndependentConfigurationSnapshot()
    {
        // Arrange
        var configuration = ExcelMapping.For<ConfiguredRow>()
            .Property(row => row.Name).HasTitle("固定名称").And()
            .Build();
        var snapshot = MappingConfigurationCloner.Clone(configuration, MappingSourceKind.Profile);
        configuration.Columns[0].Title = "外部修改";

        // Act
        var map = ExcelTypeMapFactory.Get<ConfiguredRow>(snapshot);

        // Assert
        Assert.Equal("固定名称", map.Properties.Single(property => property.Name == nameof(ConfiguredRow.Name)).Title);
    }

    /// <summary>
    /// 测试 - Attribute、Fluent、JSON 与 XML 配置应编译出等价的列定义。
    /// </summary>
    [Fact]
    public void TypeMap_ConfigurationSources_ShouldCompileEquivalentColumnDefinition()
    {
        // Arrange
        var fluent = ExcelMapping.For<EquivalentConfigurationRow>()
            .Property(row => row.Amount).HasTitle("金额").HasFormatter("0.00").HasDecimalScale(2).Map("有效", 1).And()
            .Build();
        var json = ExcelMappingConfigurationLoader.MigrateV1Json(
            "{\"columns\":[{\"propertyName\":\"Amount\",\"title\":\"金额\",\"formatter\":\"0.00\",\"decimalScale\":2,\"valueMappings\":[{\"text\":\"有效\",\"value\":\"1\"}]}]}",
            MappingDirection.Import).Import;
        var xml = ExcelMappingConfigurationLoader.MigrateV1Xml(
            "<ExcelMappingConfiguration><Columns><ExcelColumnConfiguration><PropertyName>Amount</PropertyName><Title>金额</Title><Formatter>0.00</Formatter><DecimalScale>2</DecimalScale><ValueMappings><ExcelValueMappingConfiguration><Text>有效</Text><Value>1</Value></ExcelValueMappingConfiguration></ValueMappings></ExcelColumnConfiguration></Columns></ExcelMappingConfiguration>",
            MappingDirection.Import).Import;

        // Act
        var attributeColumn = ExcelTypeMapFactory.Get<EquivalentConfigurationRow>().Properties.Single();
        var fluentColumn = ExcelTypeMapFactory.Get<EquivalentConfigurationRow>(fluent).Properties.Single();
        var jsonColumn = ExcelTypeMapFactory.Get<EquivalentConfigurationRow>(json).Properties.Single();
        var xmlColumn = ExcelTypeMapFactory.Get<EquivalentConfigurationRow>(xml).Properties.Single();

        // Assert
        AssertEquivalentColumn(attributeColumn, fluentColumn);
        AssertEquivalentColumn(attributeColumn, jsonColumn);
        AssertEquivalentColumn(attributeColumn, xmlColumn);
    }

    /// <summary>
    /// 测试 - Excel 字节数组便利扩展应保持现有流式导入导出语义。
    /// </summary>
    [Fact]
    public void StreamExtensions_ExcelBytes_ShouldRoundTripEntity()
    {
        // Arrange
        var exporter = new NpoiExcelExporter();
        var importer = new NpoiExcelImporter();

        // Act
        using var destination = new MemoryStream();
        exporter.Export(CreateSingleSheetExportRequest(new[]
        {
            new StreamRow { Name = "兼容" }
        }), destination);
        destination.Position = 0;
        var result = importer.Import(destination, CreateSingleSheetRequest<StreamRow>());

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal("兼容", Assert.Single(result.Workbook.Items).Name);
    }

    /// <summary>
    /// 测试 - JSON 与 XML 映射配置加载器应生成等价的可用配置，并拒绝 XML DTD。
    /// </summary>
    [Fact]
    public void MappingConfigurationLoader_JsonAndXml_ShouldLoadAndRejectDtd()
    {
        // Arrange
        const string json = "{\"columns\":[{\"propertyName\":\"Name\",\"title\":\"JSON 名称\"}]}";
        const string xml = "<ExcelMappingConfiguration><Columns><ExcelColumnConfiguration><PropertyName>Name</PropertyName><Title>XML 名称</Title></ExcelColumnConfiguration></Columns></ExcelMappingConfiguration>";
        const string unsafeXml = "<!DOCTYPE config [<!ENTITY value SYSTEM 'file:///not-allowed'>]><ExcelMappingConfiguration />";

        // Act
        var jsonConfiguration = ExcelMappingConfigurationLoader.MigrateV1Json(json, MappingDirection.Import).Import;
        var xmlConfiguration = ExcelMappingConfigurationLoader.MigrateV1Xml(xml, MappingDirection.Import).Import;

        // Assert
        Assert.Equal("JSON 名称", Assert.Single(jsonConfiguration.Columns).Title);
        Assert.Equal("XML 名称", Assert.Single(xmlConfiguration.Columns).Title);
            Assert.Throws<System.Xml.XmlException>(() => ExcelMappingConfigurationLoader.FromXmlDocument(unsafeXml));
    }

    /// <summary>
    /// 测试 - 配置加载器应支持 UTF-8 中文路径文件，且不得关闭调用方提供的流。
    /// </summary>
    [Fact]
    public void MappingConfigurationLoader_Utf8FilesAndStreams_ShouldKeepCallerStreamsOpen()
    {
        // Arrange
        const string json = "{\"columns\":[{\"propertyName\":\"Name\",\"title\":\"JSON 中文标题\"}]}";
        const string xml = "<ExcelMappingConfiguration><Columns><ExcelColumnConfiguration><PropertyName>Name</PropertyName><Title>XML 中文标题</Title></ExcelColumnConfiguration></Columns></ExcelMappingConfiguration>";
        var directory = Path.Combine(Path.GetTempPath(), $"Bing.Offices.配置.{Guid.NewGuid():N}");
        var jsonPath = Path.Combine(directory, "映射.json");
        var xmlPath = Path.Combine(directory, "映射.xml");
        Directory.CreateDirectory(directory);
        File.WriteAllText(jsonPath, json, System.Text.Encoding.UTF8);
        File.WriteAllText(xmlPath, xml, System.Text.Encoding.UTF8);
        using var jsonStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(json));
        using var xmlStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml));

        try
        {
            // Act
            var jsonFileConfiguration = ExcelMappingConfigurationLoader.MigrateV1Json(
                File.ReadAllText(jsonPath, System.Text.Encoding.UTF8), MappingDirection.Import).Import;
            var xmlFileConfiguration = ExcelMappingConfigurationLoader.MigrateV1Xml(
                File.ReadAllText(xmlPath, System.Text.Encoding.UTF8), MappingDirection.Import).Import;
            var jsonStreamConfiguration = ExcelMappingConfigurationLoader.MigrateV1Json(
                jsonStream, MappingDirection.Import).Import;
            var xmlStreamConfiguration = ExcelMappingConfigurationLoader.MigrateV1Xml(
                xmlStream, MappingDirection.Import).Import;

            // Assert
            Assert.Equal("JSON 中文标题", Assert.Single(jsonFileConfiguration.Columns).Title);
            Assert.Equal("XML 中文标题", Assert.Single(xmlFileConfiguration.Columns).Title);
            Assert.Equal("JSON 中文标题", Assert.Single(jsonStreamConfiguration.Columns).Title);
            Assert.Equal("XML 中文标题", Assert.Single(xmlStreamConfiguration.Columns).Title);
            Assert.True(jsonStream.CanRead);
            Assert.True(xmlStream.CanRead);
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 测试 - 文件入口也应执行与流入口相同的输入大小限制，避免无界读取。
    /// </summary>
    [Fact]
    public void MappingConfigurationLoader_OversizedFiles_ShouldRejectBeforeUnboundedRead()
    {
        // Arrange
        var directory = Path.Combine(Path.GetTempPath(), $"Bing.Offices.配置.{Guid.NewGuid():N}");
        var jsonPath = Path.Combine(directory, "oversized.json");
        var xmlPath = Path.Combine(directory, "oversized.xml");
        Directory.CreateDirectory(directory);
        File.WriteAllText(jsonPath, new string('x', 1024 * 1024 + 1), System.Text.Encoding.UTF8);
        File.WriteAllText(xmlPath, new string('x', 1024 * 1024 + 1), System.Text.Encoding.UTF8);

        try
        {
            // Act / Assert
            Assert.Throws<InvalidOperationException>(() => ExcelMappingConfigurationLoader.FromJsonDocument(
                File.ReadAllText(jsonPath, Encoding.UTF8)));
            Assert.Throws<InvalidOperationException>(() => ExcelMappingConfigurationLoader.FromXmlDocument(
                File.ReadAllText(xmlPath, Encoding.UTF8)));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 测试 - JSON v1 和 v2 应归一化为同一导入配置，且 v2 保留独立导出方向。
    /// </summary>
    [Fact]
    public void MappingConfigurationLoader_JsonV1AndV2_ShouldNormalizeToEquivalentDocument()
    {
        // Arrange
        const string v1 = "{\"columns\":[{\"propertyName\":\"Name\",\"title\":\"名称\",\"aliases\":[\"旧名称\"]}]}";
        const string v2 = "{\"version\":2,\"import\":{\"profile\":\"orders\",\"modelAlias\":\"order-row\",\"columns\":[{\"propertyName\":\"Name\",\"title\":\"名称\",\"aliases\":[\"旧名称\"]}]},\"export\":{\"profile\":\"orders\",\"modelAlias\":\"order-row\",\"columns\":[{\"propertyName\":\"Name\",\"title\":\"导出名称\"}]}}";

        // Act
        var migrated = ExcelMappingConfigurationLoader.MigrateV1Json(v1, MappingDirection.Import);
        var document = ExcelMappingConfigurationLoader.FromJsonDocument(v2);

        // Assert
        Assert.Equal(2, migrated.Version);
        Assert.Equal("名称", Assert.Single(migrated.Import.Columns).Title);
        Assert.Equal("名称", Assert.Single(document.Import.Columns).Title);
        Assert.Equal("旧名称", Assert.Single(document.Import.Columns).Aliases[0]);
        Assert.Equal("导出名称", Assert.Single(document.Export.Columns).Title);
        Assert.Equal("orders", document.Import.Profile);
        Assert.Equal("order-row", document.Import.ModelAlias);
    }

    /// <summary>
    /// 测试 - XML v2 与 JSON v2 应生成相同方向配置，并拒绝未知节点。
    /// </summary>
    [Fact]
    public void MappingConfigurationLoader_XmlV2_ShouldNormalizeAndRejectUnknownNodes()
    {
        // Arrange
        const string xml = "<ExcelMappingDocument><Version>2</Version><Import><Profile>orders</Profile><ModelAlias>order-row</ModelAlias><Columns><ExcelColumnConfiguration><PropertyName>Name</PropertyName><Title>名称</Title><Aliases><string>旧名称</string></Aliases></ExcelColumnConfiguration></Columns></Import><Export><Profile>orders</Profile><ModelAlias>order-row</ModelAlias><Columns><ExcelColumnConfiguration><PropertyName>Name</PropertyName><Title>导出名称</Title></ExcelColumnConfiguration></Columns></Export></ExcelMappingDocument>";
        const string unknownXml = "<ExcelMappingDocument><Version>2</Version><Unknown>value</Unknown></ExcelMappingDocument>";

        // Act
        var document = ExcelMappingConfigurationLoader.FromXmlDocument(xml);

        // Assert
        Assert.Equal(2, document.Version);
        Assert.Equal("orders", document.Import.Profile);
        Assert.Equal("名称", Assert.Single(document.Import.Columns).Title);
        Assert.Equal("导出名称", Assert.Single(document.Export.Columns).Title);
        var exception = Assert.Throws<InvalidOperationException>(() =>
            ExcelMappingConfigurationLoader.FromXmlDocument(unknownXml));
        Assert.Contains("/ExcelMappingDocument/Unknown", exception.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// 测试 - 配置加载器应拒绝未知 JSON 字段、过长字符串、过深文档和超大输入。
    /// </summary>
    [Fact]
    public void MappingConfigurationLoader_InputLimits_ShouldRejectUnsafeDocuments()
    {
        // Arrange
        var unknown = "{\"columns\":[{\"propertyName\":\"Name\",\"unknown\":true}]}";
        var longTitle = "{\"columns\":[{\"propertyName\":\"Name\",\"title\":\"" + new string('x', 4097) + "\"}]}";
        var deep = "{\"version\":2,\"import\":{\"columns\":[{\"aliases\":[" + new string('[', 40) + "\"x\"" + new string(']', 40) + "]}}";
        var oversized = new string('x', 1024 * 1024 + 1);

        // Act / Assert
        var unknownException = Assert.Throws<InvalidOperationException>(() => ExcelMappingConfigurationLoader.FromJsonDocument(unknown));
        Assert.Contains("$.columns[0].unknown", unknownException.Message);
        Assert.Throws<InvalidOperationException>(() => ExcelMappingConfigurationLoader.FromJsonDocument(longTitle));
        Assert.ThrowsAny<Exception>(() => ExcelMappingConfigurationLoader.FromJsonDocument(deep));
        Assert.Throws<InvalidOperationException>(() => ExcelMappingConfigurationLoader.FromJsonDocument(oversized));
    }

    /// <summary>
    /// 测试 - v2 文档流读取应保持调用方流可用，XML DTD/外部实体仍必须被禁止。
    /// </summary>
    [Fact]
    public void MappingConfigurationLoader_DocumentStreams_ShouldKeepOwnershipAndRejectDtd()
    {
        // Arrange
        const string json = "{\"version\":2,\"import\":{\"modelAlias\":\"row\",\"columns\":[]},\"export\":{\"columns\":[]}}";
        const string xml = "<ExcelMappingDocument><Version>2</Version><Import><ModelAlias>row</ModelAlias><Columns /></Import><Export><Columns /></Export></ExcelMappingDocument>";
        const string unsafeXml = "<!DOCTYPE config [<!ENTITY value SYSTEM 'file:///not-allowed'>]><ExcelMappingDocument />";
        using var jsonStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(json));
        using var xmlStream = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(xml));

        // Act
        var jsonDocument = ExcelMappingConfigurationLoader.FromJsonDocument(jsonStream);
        var xmlDocument = ExcelMappingConfigurationLoader.FromXmlDocument(xmlStream);

        // Assert
        Assert.Equal("row", jsonDocument.Import.ModelAlias);
        Assert.Equal("row", xmlDocument.Import.ModelAlias);
        Assert.True(jsonStream.CanRead);
        Assert.True(xmlStream.CanRead);
        Assert.Throws<System.Xml.XmlException>(() => ExcelMappingConfigurationLoader.FromXmlDocument(unsafeXml));
    }

    /// <summary>
    /// 测试 - normalized document 编译时应按方向选择独立配置并继续支持请求级覆盖。
    /// </summary>
    [Fact]
    public void TypeMap_NormalizedDocument_ShouldCompileSelectedDirection()
    {
        // Arrange
        var document = ExcelMappingConfigurationLoader.FromJsonDocument(
            "{\"version\":2,\"import\":{\"columns\":[{\"propertyName\":\"Name\",\"title\":\"导入名称\"}]},\"export\":{\"columns\":[{\"propertyName\":\"Name\",\"title\":\"导出名称\"}]}}");
        var request = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(ConfiguredRow.Name), Title = "请求名称" }
            }
        };

        // Act
        var importMap = ExcelTypeMapFactory.Get<ConfiguredRow>(document, MappingDirection.Import);
        var exportMap = ExcelTypeMapFactory.Get<ConfiguredRow>(document, request, MappingDirection.Export);

        // Assert
        Assert.Equal("导入名称", importMap.Properties.Single(property => property.Name == nameof(ConfiguredRow.Name)).Title);
        Assert.Equal("请求名称", exportMap.Properties.Single(property => property.Name == nameof(ConfiguredRow.Name)).Title);
    }

    /// <summary>
    /// 测试 - JSON 映射中的转换器名称应仅选择已注册的同名转换器。
    /// </summary>
    [Fact]
    public void StreamPipeline_JsonNamedConverter_ShouldRoundTripDomainValue()
    {
        // Arrange
        var configuration = ExcelMappingConfigurationLoader.MigrateV1Json(
            "{\"columns\":[{\"propertyName\":\"Code\",\"converterName\":\"order-code\"}]}",
            MappingDirection.Import).Import;
        var converter = new OrderCodeExcelValueConverter();
        using var destination = new MemoryStream();
        var exporter = new NpoiExcelExporter(new IExcelValueConverter[] { converter });
        var importer = new NpoiExcelImporter(valueConverters: new IExcelValueConverter[] { converter });

        // Act
        exporter.Export(CreateSingleSheetExportRequest(new[] { new ConvertedRow { Code = new OrderCode("42") } },
            configure: sheet => sheet.Mapping(configuration)), destination);
        destination.Position = 0;
        var result = importer.Import(destination, CreateSingleSheetRequest<ConvertedRow>(sheet =>
            sheet.Mapping(configuration)));

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal("42", Assert.Single(result.Workbook.Items).Code.Value);
    }

    /// <summary>
    /// 测试 - JSON 映射中的校验规则名称应仅使用已注册的同名规则。
    /// </summary>
    [Fact]
    public void StreamPipeline_JsonNamedValidationRule_ShouldReturnValidationError()
    {
        // Arrange
        var configuration = ExcelMappingConfigurationLoader.MigrateV1Json(
            "{\"columns\":[{\"propertyName\":\"Name\",\"validationRuleNames\":[\"starts-with-ok\"]}]}",
            MappingDirection.Import).Import;
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(StreamRow.Name));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("invalid");
        }));
        var importer = new NpoiExcelImporter(namedValidationRules: new INamedExcelValidationRule[]
        {
            new StartsWithOkValidationRule()
        });

        // Act
        var result = importer.Import(source, CreateSingleSheetRequest<StreamRow>(sheet =>
            sheet.Mapping(configuration)));

        // Assert
        Assert.Empty(result.Workbook.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal(nameof(StreamRow.Name), error.PropertyName);
    }

    /// <summary>
    /// 测试 - Fluent 与 XML 映射中的命名校验规则应使用已注册规则执行校验。
    /// </summary>
    [Fact]
    public void StreamPipeline_FluentAndXmlNamedValidationRules_ShouldReturnValidationErrors()
    {
        // Arrange
        var fluent = ExcelMapping.For<StreamRow>()
            .Property(row => row.Name).HasValidationRule("starts-with-ok").And()
            .Build();
        var xml = ExcelMappingConfigurationLoader.MigrateV1Xml(
            "<ExcelMappingConfiguration><Columns><ExcelColumnConfiguration><PropertyName>Name</PropertyName><ValidationRuleNames><string>starts-with-ok</string></ValidationRuleNames></ExcelColumnConfiguration></Columns></ExcelMappingConfiguration>",
            MappingDirection.Import).Import;
        var importer = new NpoiExcelImporter(namedValidationRules: new INamedExcelValidationRule[]
        {
            new StartsWithOkValidationRule()
        });
        using var fluentSource = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(StreamRow.Name));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("invalid");
        }));
        using var xmlSource = new MemoryStream(fluentSource.ToArray());

        // Act
        var fluentResult = importer.Import(fluentSource, CreateSingleSheetRequest<StreamRow>(sheet =>
            sheet.Mapping(fluent)));
        var xmlResult = importer.Import(xmlSource, CreateSingleSheetRequest<StreamRow>(sheet =>
            sheet.Mapping(xml)));

        // Assert
        Assert.Empty(fluentResult.Workbook.Items);
        Assert.Empty(xmlResult.Workbook.Items);
        Assert.Equal(ExcelImportErrorCode.Validation, Assert.Single(fluentResult.Errors).Code);
        Assert.Equal(ExcelImportErrorCode.Validation, Assert.Single(xmlResult.Errors).Code);
    }

    /// <summary>
    /// 测试 - 命名校验规则应接收单元格语义、请求区域性和目标属性元数据。
    /// </summary>
    [Fact]
    public void Import_NamedValidationRule_ShouldExposeFullValidationContext()
    {
        // Arrange
        var configuration = ExcelMappingConfigurationLoader.MigrateV1Json(
            "{\"columns\":[{\"propertyName\":\"Name\",\"validationRuleNames\":[\"context\"]}]}",
            MappingDirection.Import).Import;
        var rule = new ContextCapturingValidationRule();
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(StreamRow.Name));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("上下文");
        }));

        // Act
        var result = new NpoiExcelImporter(namedValidationRules: new INamedExcelValidationRule[] { rule })
            .Import(source, CreateSingleSheetRequest<StreamRow>(sheet => sheet
                .Mapping(configuration)
                .Culture(CultureInfo.GetCultureInfo("fr-FR"))));

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal(ExcelCellKind.Text, rule.CellKind);
        Assert.Equal("fr-FR", rule.CultureName);
        Assert.Equal(nameof(StreamRow.Name), rule.PropertyName);
    }

    /// <summary>
    /// 测试 - 并发请求使用不同动态列和表头映射时，不应污染类型静态映射缓存。
    /// </summary>
    [Fact]
    public async Task StreamPipeline_ConcurrentRequestOptions_ShouldRemainIsolated()
    {
        // Arrange
        const int requestCount = 24;
        var tasks = Enumerable.Range(0, requestCount).Select(index => Task.Run(() =>
        {
            var header = $"业务编码{index}";
            var dynamicColumn = $"扩展字段{index}";
            using var importSource = new MemoryStream(CreateWorkbook(workbook =>
            {
                var sheet = workbook.CreateSheet("Data");
                var headerRow = sheet.CreateRow(0);
                headerRow.CreateCell(0).SetCellValue(header);
                headerRow.CreateCell(1).SetCellValue(dynamicColumn);
                var dataRow = sheet.CreateRow(1);
                dataRow.CreateCell(0).SetCellValue($"Code-{index}");
                dataRow.CreateCell(1).SetCellValue($"Value-{index}");
            }));
            var importResult = new NpoiExcelImporter().Import(importSource,
                ExcelImport.Workbook<SingleWorkbook<DynamicRow>>(builder => builder.Sheet("Data",
                    root => root.Items, sheet => sheet.Mapping(new ExcelMappingConfiguration
                    {
                        Columns = new List<ExcelColumnConfiguration>
                        {
                            new() { PropertyName = nameof(DynamicRow.Code), Title = header }
                        }
                    }))));

            using var destination = new MemoryStream();
            new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(new[]
            {
                new DynamicRow
                {
                    Code = $"Code-{index}",
                    Values = new Dictionary<string, object> { [dynamicColumn] = $"Value-{index}" }
                }
            }, configure: sheet => sheet.DynamicColumns(row => row.Values, new[]
            {
                new ExcelDynamicColumnDefinition
                {
                    Key = dynamicColumn,
                    Title = dynamicColumn,
                    DataType = typeof(object)
                }
            })), destination);
            destination.Position = 0;
            using var workbook = WorkbookFactory.Create(destination);
            return (Item: Assert.Single(importResult.Workbook.Items), Sheet: workbook.GetSheetAt(0),
                DynamicColumn: dynamicColumn, Index: index);
        }));

        // Act
        var results = await Task.WhenAll(tasks);

        // Assert
        foreach (var result in results)
        {
            Assert.Equal($"Code-{result.Index}", result.Item.Code);
            Assert.Equal($"Value-{result.Index}", result.Item.Values[result.DynamicColumn]);
            Assert.Equal(result.DynamicColumn, result.Sheet.GetRow(0).GetCell(1).StringCellValue);
            Assert.Equal($"Value-{result.Index}", result.Sheet.GetRow(1).GetCell(1).StringCellValue);
        }
    }

    /// <summary>
    /// 测试 - 导入器应读取公式缓存值，并以固定区域性转换数值。
    /// </summary>
    [Fact]
    public void Import_FormulaAndInvariantNumber_ShouldBindValues()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(NumberRow.Amount));
            var formulaCell = sheet.CreateRow(1).CreateCell(0);
            formulaCell.SetCellFormula("1.5+2.25");
            workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateFormulaCell(formulaCell);
        }));
        var originalCulture = CultureInfo.CurrentCulture;
        CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");

        try
        {
            // Act
            var result = new NpoiExcelImporter().Import(source, CreateSingleSheetRequest<NumberRow>());

            // Assert
            Assert.Empty(result.Errors);
            Assert.Equal(3.75m, Assert.Single(result.Workbook.Items).Amount);
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
        }
    }

    /// <summary>
    /// 测试 - 自定义转换器应能读取与 NPOI 无关的公式及缓存单元格语义。
    /// </summary>
    [Fact]
    public void Import_FormulaCell_ShouldExposeProviderIndependentConversionContext()
    {
        // Arrange
        var converter = new FormulaContextExcelValueConverter();
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(NumberRow.Amount));
            var formula = sheet.CreateRow(1).CreateCell(0);
            formula.SetCellFormula("1.5+2.25");
            workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateFormulaCell(formula);
        }));

        // Act
        var result = new NpoiExcelImporter(valueConverters: new IExcelValueConverter[] { converter })
            .Import(source, CreateSingleSheetRequest<NumberRow>());

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal(3.75m, Assert.Single(result.Workbook.Items).Amount);
        Assert.Equal(ExcelCellKind.Formula, converter.Kind);
        Assert.Equal(ExcelCellKind.Number, converter.CachedKind);
        Assert.Equal("1.5+2.25", converter.Formula);
        Assert.Equal(3.75d, converter.RawValue);
    }

    /// <summary>
    /// 测试 - 旧版文本转换器应仅作为输入文本薄桥参与 Stream-first 导入。
    /// </summary>
    [Fact]
    public void Import_LegacyCellValueConverter_ShouldAdaptTextOnly()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(StreamRow.Name));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("legacy");
        }));

        // Act
#pragma warning disable CS0618
        var result = new NpoiExcelImporter(legacyValueConverters: new ICellValueConverter[]
        {
            new PrefixLegacyCellValueConverter()
        }).Import(source, CreateSingleSheetRequest<StreamRow>());
#pragma warning restore CS0618

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal("adapted-legacy", Assert.Single(result.Workbook.Items).Name);
    }

    /// <summary>
    /// 测试 - 格式化数值应以不变区域性写入，确保可被导入器无损解析。
    /// </summary>
    [Fact]
    public void StreamPipeline_FormattedNumberUnderDifferentCulture_ShouldRoundTrip()
    {
        // Arrange
        var originalCulture = CultureInfo.CurrentCulture;
        CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("fr-FR");
        using var destination = new MemoryStream();

        try
        {
            // Act
            new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(
                new[] { new FormattedNumberRow { Amount = 1.5m } }), destination);
            destination.Position = 0;
            var result = new NpoiExcelImporter().Import(destination, CreateSingleSheetRequest<FormattedNumberRow>());

            // Assert
            Assert.Empty(result.Errors);
            Assert.Equal(1.5m, Assert.Single(result.Workbook.Items).Amount);
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
        }
    }

    /// <summary>
    /// 测试 - 导入器应按公式缓存结果绑定数值、日期、布尔和字符串类型。
    /// </summary>
    [Fact]
    public void Import_FormulaCachedValues_ShouldBindSupportedTypes()
    {
        // Arrange
        var expectedDate = new DateTime(2026, 8, 13);
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue(nameof(FormulaRow.Amount));
            header.CreateCell(1).SetCellValue(nameof(FormulaRow.OccurredAt));
            header.CreateCell(2).SetCellValue(nameof(FormulaRow.Enabled));
            header.CreateCell(3).SetCellValue(nameof(FormulaRow.Name));
            header.CreateCell(4).SetCellValue(nameof(FormulaRow.OptionalOccurredAt));
            var row = sheet.CreateRow(1);
            var amount = row.CreateCell(0);
            amount.SetCellFormula("1.5+2.25");
            var date = row.CreateCell(1);
            date.SetCellFormula("DATE(2026,8,13)");
            date.CellStyle = workbook.CreateCellStyle();
            date.CellStyle.DataFormat = workbook.CreateDataFormat().GetFormat("yyyy-mm-dd");
            var enabled = row.CreateCell(2);
            enabled.SetCellFormula("1=1");
            var name = row.CreateCell(3);
            name.SetCellFormula("\"Formula\"&\"Value\"");
            row.CreateCell(4, CellType.Blank);
            var evaluator = workbook.GetCreationHelper().CreateFormulaEvaluator();
            evaluator.EvaluateFormulaCell(amount);
            evaluator.EvaluateFormulaCell(date);
            evaluator.EvaluateFormulaCell(enabled);
            evaluator.EvaluateFormulaCell(name);
        }));

        // Act
        var result = new NpoiExcelImporter().Import(source, CreateSingleSheetRequest<FormulaRow>());

        // Assert
        var item = Assert.Single(result.Workbook.Items);
        Assert.Empty(result.Errors);
        Assert.Equal(3.75m, item.Amount);
        Assert.Equal(expectedDate, item.OccurredAt);
        Assert.Null(item.OptionalOccurredAt);
        Assert.True(item.Enabled);
        Assert.Equal("FormulaValue", item.Name);
    }

    /// <summary>
    /// 测试 - 非可空值为空或 Guid 格式无效时应返回定位准确的转换错误且不返回失败行。
    /// </summary>
    [Fact]
    public void Import_BlankRequiredValueAndInvalidGuid_ShouldReturnStructuredConversionErrors()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue(nameof(ConversionErrorRow.Count));
            header.CreateCell(1).SetCellValue(nameof(ConversionErrorRow.Id));
            var blankCount = sheet.CreateRow(1);
            blankCount.CreateCell(0, CellType.Blank);
            blankCount.CreateCell(1).SetCellValue("de305d54-75b4-431b-adb2-eb6b9e546014");
            var invalidGuid = sheet.CreateRow(2);
            invalidGuid.CreateCell(0).SetCellValue(1);
            invalidGuid.CreateCell(1).SetCellValue("not-a-guid");
        }));

        // Act
        var result = new NpoiExcelImporter().Import(source, CreateSingleSheetRequest<ConversionErrorRow>());

        // Assert
        Assert.Empty(result.Workbook.Items);
        Assert.Collection(result.Errors,
            error =>
            {
                Assert.Equal(ExcelImportErrorCode.ValueConversion, error.Code);
                Assert.Equal(2, error.RowIndex);
                Assert.Equal(1, error.ColumnIndex);
                Assert.Equal(nameof(ConversionErrorRow.Count), error.PropertyName);
            },
            error =>
            {
                Assert.Equal(ExcelImportErrorCode.ValueConversion, error.Code);
                Assert.Equal(3, error.RowIndex);
                Assert.Equal(2, error.ColumnIndex);
                Assert.Equal(nameof(ConversionErrorRow.Id), error.PropertyName);
            });
    }

    /// <summary>
    /// 测试 - 公式缓存为错误类型时应返回结构化转换错误，不得将错误码绑定为业务数值。
    /// </summary>
    [Fact]
    public void Import_FormulaErrorCachedValue_ShouldReturnStructuredConversionError()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(NumberRow.Amount));
            var formula = sheet.CreateRow(1).CreateCell(0);
            formula.SetCellFormula("1/0");
            workbook.GetCreationHelper().CreateFormulaEvaluator().EvaluateFormulaCell(formula);
        }));

        // Act
        var result = new NpoiExcelImporter().Import(source, CreateSingleSheetRequest<NumberRow>());

        // Assert
        Assert.Empty(result.Workbook.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, error.Code);
        Assert.Equal("Data", error.SheetName);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
        Assert.Equal(nameof(NumberRow.Amount), error.PropertyName);
    }

    /// <summary>
    /// 测试 - 导入导出应保留 Guid、Version、布尔、长整型枚举和 Flags 枚举的业务值。
    /// </summary>
    [Fact]
    public void StreamPipeline_SupportedScalarTypes_ShouldRoundTripValues()
    {
        // Arrange
        var item = new ScalarRow
        {
            Id = Guid.Parse("de305d54-75b4-431b-adb2-eb6b9e546014"),
            OptionalId = null,
            Version = new Version(1, 2, 3, 4),
            Enabled = true,
            Status = LongStatus.Archived,
            Access = AccessFlags.Read | AccessFlags.Write
        };
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(new[] { item }), destination);
        destination.Position = 0;
        var result = new NpoiExcelImporter().Import(destination, CreateSingleSheetRequest<ScalarRow>());

        // Assert
        var actual = Assert.Single(result.Workbook.Items);
        Assert.Empty(result.Errors);
        Assert.Equal(item.Id, actual.Id);
        Assert.Null(actual.OptionalId);
        Assert.Equal(item.Version, actual.Version);
        Assert.Equal(item.Enabled, actual.Enabled);
        Assert.Equal(item.Status, actual.Status);
        Assert.Equal(item.Access, actual.Access);
    }

    /// <summary>
    /// 测试 - 自定义值映射的显示文本和业务值均必须唯一，确保导入导出可逆。
    /// </summary>
    [Fact]
    public void TypeMap_DuplicateCustomValueMappings_ShouldThrowOfficeException()
    {
        // Act
        var duplicateText = () => ExcelTypeMapFactory.Get<DuplicateTextMappingRow>();
        var duplicateValue = () => ExcelTypeMapFactory.Get<DuplicateValueMappingRow>();

        // Assert
        Assert.Throws<OfficeException>(duplicateText);
        Assert.Throws<OfficeException>(duplicateValue);
    }

    /// <summary>
    /// 测试 - 动态列属性必须使用可写的对象字典，避免导入导出能力不对称。
    /// </summary>
    [Fact]
    public void TypeMap_InvalidDynamicColumnType_ShouldThrowOfficeException()
    {
        // Arrange
        var action = () => ExcelTypeMapFactory.Get<InvalidDynamicRow>();

        // Act and Assert
        Assert.Throws<OfficeException>(action);
    }

    /// <summary>
    /// 测试 - 匹配到只读属性的导入模板应在绑定前被明确拒绝。
    /// </summary>
    [Fact]
    public void Import_ReadOnlyMappedProperty_ShouldThrowConfigurationException()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(ReadOnlyRow.Code));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("Value");
        }));

        // Act
        var action = () => new NpoiExcelImporter().Import(source, CreateSingleSheetRequest<ReadOnlyRow>());

        // Assert
        Assert.Throws<InvalidOperationException>(action);
    }

    /// <summary>
    /// 测试 - 多工作表导入应跳过隐藏工作表，且重复校验范围限定在各自工作表。
    /// </summary>
    [Fact]
    public void Import_MultiSheet_ShouldSkipHiddenSheetsAndScopeDuplicatesPerSheet()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            CreateValidatedSheet(workbook, "First", "Code-A");
            CreateValidatedSheet(workbook, "Second", "code-a");
            CreateValidatedSheet(workbook, "Hidden", "Code-B");
            workbook.SetSheetHidden(2, true);
        }));

        // Act
        var result = new NpoiExcelImporter().Import(source,
            ExcelImport.Workbook<MultiSheetWorkbook<ValidatedRow>>(builder => builder
                .Sheet("First", root => root.Items)
                .Sheet("Second", root => root.Items)));

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal(2, result.Workbook.Items.Count);
    }

    /// <summary>
    /// 测试 - 非法最大列数和多个动态列声明应被明确拒绝。
    /// </summary>
    [Fact]
    public void Import_InvalidColumnConfigurationAndMultipleDynamicProperties_ShouldThrow()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(StreamRow.Name));
        }));

        // Act
        var invalidLength = () => new NpoiExcelImporter().Import(source,
            CreateSingleSheetRequest<StreamRow>(sheet => sheet.MaxColumnCount(0)));
        source.Position = 0;
        var multipleDynamic = () => new NpoiExcelImporter().Import(source,
            CreateSingleSheetRequest<MultipleDynamicRow>(sheet => sheet.HeaderMatch(false)));

        // Assert
        Assert.Throws<ArgumentOutOfRangeException>(invalidLength);
        Assert.Throws<InvalidOperationException>(multipleDynamic);
    }

    /// <summary>
    /// 测试 - 数据起始行必须位于表头行之后。
    /// </summary>
    [Fact]
    public void Import_DataRowIndexBeforeOrEqualToHeaderRowIndex_ShouldThrow()
    {
        // Arrange
        var workbookBytes = CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(1).CreateCell(0).SetCellValue(nameof(StreamRow.Name));
            sheet.CreateRow(2).CreateCell(0).SetCellValue("Value");
        });

        // Act
        var sameRow = () =>
        {
            using var source = new MemoryStream(workbookBytes, writable: false);
            return new NpoiExcelImporter().Import(source, CreateSingleSheetRequest<StreamRow>(sheet => sheet
                .HeaderRowIndex(1)
                .DataRowStartIndex(1)));
        };
        var precedingRow = () =>
        {
            using var source = new MemoryStream(workbookBytes, writable: false);
            return new NpoiExcelImporter().Import(source, CreateSingleSheetRequest<StreamRow>(sheet => sheet
                .HeaderRowIndex(1)
                .DataRowStartIndex(0)));
        };

        // Assert
        Assert.Throws<ArgumentOutOfRangeException>(sameRow);
        Assert.Throws<ArgumentOutOfRangeException>(precedingRow);
    }

    /// <summary>
    /// 测试 - 未定义的校验模式应在读取工作簿前被拒绝。
    /// </summary>
    [Fact]
    public void Import_UndefinedValidateMode_ShouldThrow()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(StreamRow.Name));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("Value");
        }));

        // Act
        var import = () => new NpoiExcelImporter().Import(source,
            CreateSingleSheetRequest<StreamRow>(sheet => sheet.Validate((ValidateMode)99)));

        // Assert
        Assert.Throws<ArgumentOutOfRangeException>(import);
    }

    /// <summary>
    /// 测试 - 请求级表头映射指向动态列属性时应被拒绝，避免配置被静默忽略。
    /// </summary>
    [Fact]
    public void Import_HeaderMappingForDynamicProperty_ShouldThrow()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("扩展字段");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("扩展值");
        }));

        // Act
        var import = () => new NpoiExcelImporter().Import(source,
            ExcelImport.Workbook<SingleWorkbook<DynamicRow>>(builder => builder.Sheet("Data", root => root.Items,
                sheet => sheet.Mapping(new ExcelMappingConfiguration
                {
                    Columns = new List<ExcelColumnConfiguration>
                    {
                        new() { PropertyName = nameof(DynamicRow.Values), Title = "扩展字段" }
                    }
                }))));

        // Assert
        Assert.Throws<InvalidOperationException>(import);
    }

    /// <summary>
    /// 测试 - 请求级表头映射引用未知属性或将多个属性映射为同一表头时应被拒绝。
    /// </summary>
    [Fact]
    public void Import_InvalidHeaderMappings_ShouldThrow()
    {
        // Arrange
        var workbookBytes = CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue("共享表头");
            sheet.CreateRow(1).CreateCell(0).SetCellValue("值");
        });
        var unknownMapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = "Unknown", Title = "共享表头" }
            }
        };
        var duplicateMapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(MappingRow.First), Title = "共享表头" },
                new() { PropertyName = nameof(MappingRow.Second), Title = "共享表头" }
            }
        };

        // Act
        var importUnknownProperty = () =>
        {
            using var source = new MemoryStream(workbookBytes, writable: false);
            return new NpoiExcelImporter().Import(source, CreateSingleSheetRequest<MappingRow>(sheet =>
                sheet.Mapping(unknownMapping)));
        };
        var importDuplicateHeader = () =>
        {
            using var source = new MemoryStream(workbookBytes, writable: false);
            return new NpoiExcelImporter().Import(source, CreateSingleSheetRequest<MappingRow>(sheet =>
                sheet.Mapping(duplicateMapping)));
        };

        // Assert
        Assert.Throws<ArgumentException>(importUnknownProperty);
        Assert.Throws<ArgumentException>(importDuplicateHeader);
    }

    /// <summary>
    /// 测试 - 导入器应支持 Xls 格式，并在模板读取失败时保持调用方输入流打开。
    /// </summary>
    [Fact]
    public void Import_XlsAndInvalidWorkbook_ShouldSupportFormatAndKeepSourceOpen()
    {
        // Arrange
        using var xlsSource = new MemoryStream(CreateWorkbook<HSSFWorkbook>(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(StreamRow.Name));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("XlsValue");
        }));
        using var invalidSource = new NonSeekableReadStream(new byte[] { 1, 2, 3, 4 });

        // Act
        var result = new NpoiExcelImporter().Import(xlsSource, CreateSingleSheetRequest<StreamRow>());
        var invalidImport = () => new NpoiExcelImporter().Import(invalidSource,
            CreateSingleSheetRequest<StreamRow>());

        // Assert
        Assert.Equal("XlsValue", Assert.Single(result.Workbook.Items).Name);
        Assert.ThrowsAny<Exception>(invalidImport);
        Assert.True(invalidSource.CanRead);
    }

    /// <summary>
    /// 测试 - 导出器应写入单一工作簿、保留目标流并应用动态列与自定义表头布局。
    /// </summary>
    [Fact]
    public void Export_StreamDestination_ShouldWriteLayoutAndDynamicValues()
    {
        // Arrange
        using var destination = new MemoryStream();
        var data = new[]
        {
            new DynamicRow
            {
                Code = "A-1",
                Values = new Dictionary<string, object> { ["扩展字段"] = "扩展值" }
            }
        };
        var request = CreateSingleSheetExportRequest(data, configure: sheet => sheet
            .DynamicColumns(row => row.Values, new[]
            {
                new ExcelDynamicColumnDefinition { Key = "扩展字段", Title = "扩展字段" }
            })
            .HeaderRowIndex(2)
            .DataRowStartIndex(3)
            .HeaderRows(new[]
            {
                new ExcelHeaderRow(0, new[] { new ExcelHeaderCell(0, "总标题", columnSpan: 2, rowSpan: 2) })
            }));

        // Act
        new NpoiExcelExporter().Export(request, destination);

        // Assert
        Assert.True(destination.CanWrite);
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var sheet = workbook.GetSheetAt(0);
        Assert.Equal("总标题", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal(nameof(DynamicRow.Code), sheet.GetRow(2).GetCell(0).StringCellValue);
        Assert.Equal("扩展字段", sheet.GetRow(2).GetCell(1).StringCellValue);
        Assert.Equal("A-1", sheet.GetRow(3).GetCell(0).StringCellValue);
        Assert.Equal("扩展值", sheet.GetRow(3).GetCell(1).StringCellValue);
        Assert.Single(sheet.MergedRegions);
    }

    /// <summary>
    /// 测试 - 导出器应支持不可寻址目标流，并保持调用方流打开。
    /// </summary>
    [Fact]
    public void Export_NonSeekableDestination_ShouldKeepStreamOpenAndWriteWorkbook()
    {
        // Arrange
        using var destination = new NonSeekableWriteStream();

        // Act
        new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(
            new[] { new StreamRow { Name = "Written" } }), destination);

        // Assert
        Assert.True(destination.CanWrite);
        using var workbook = WorkbookFactory.Create(new MemoryStream(destination.ToArray(), writable: false));
        Assert.Equal("Written", workbook.GetSheetAt(0).GetRow(1).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 测试 - 动态 byte[] 图片应按 GIF 文件签名注册为 GIF 类型。
    /// </summary>
    [Fact]
    public void Export_DynamicGifPicture_ShouldRegisterGifPictureType()
    {
        // Arrange
        var gifBytes = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");
        using var destination = new MemoryStream();
        var data = new[]
        {
            new DynamicRow
            {
                Code = "GIF",
                Values = new Dictionary<string, object> { ["图片"] = gifBytes }
            }
        };

        // Act
        new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(data, configure: sheet =>
            sheet.DynamicColumns(row => row.Values, new[]
            {
                new ExcelDynamicColumnDefinition { Key = "图片", Title = "图片", DataType = typeof(byte[]) }
            })), destination);

        // Assert
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var picture = Assert.IsAssignableFrom<IPictureData>(Assert.Single(workbook.GetAllPictures()));
        Assert.Equal(PictureType.GIF, picture.PictureType);
    }

    /// <summary>
    /// 测试 - 动态列值可由任意 IDictionary 实现提供，导出器不应依赖具体字典类型。
    /// </summary>
    [Fact]
    public void Export_DynamicValuesWithNonDictionaryImplementation_ShouldWriteValue()
    {
        // Arrange
        using var destination = new MemoryStream();
        var data = new[]
        {
            new DynamicRow
            {
                Code = "Sorted",
                Values = new SortedDictionary<string, object> { ["扩展字段"] = "扩展值" }
            }
        };

        // Act
        new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(data, configure: sheet =>
            sheet.DynamicColumns(row => row.Values, new[]
            {
                new ExcelDynamicColumnDefinition { Key = "扩展字段", Title = "扩展字段" }
            })), destination);

        // Assert
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        Assert.Equal("扩展值", workbook.GetSheetAt(0).GetRow(1).GetCell(1).StringCellValue);
    }

    /// <summary>
    /// 测试 - 导出选项的行索引、工作表名称和动态列标题冲突应在写入前被拒绝。
    /// </summary>
    [Fact]
    public void Export_InvalidOptionsAndDuplicateColumns_ShouldThrow()
    {
        // Arrange
        var data = new[] { new DynamicRow { Code = "A-1" } };

        // Act
        var invalidDataRow = () => new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(data,
            configure: sheet => sheet.DataRowStartIndex(0)), new MemoryStream());
        var invalidSheetName = () => new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(data, " "),
            new MemoryStream());
        var duplicateColumns = () => new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(data,
            configure: sheet => sheet.DynamicColumns(row => row.Values, new[]
            {
                new ExcelDynamicColumnDefinition { Key = "Code", Title = "Code" }
            })), new MemoryStream());

        // Assert
        Assert.Throws<ArgumentOutOfRangeException>(invalidDataRow);
        Assert.Throws<ArgumentException>(invalidSheetName);
        Assert.Throws<ArgumentException>(duplicateColumns);
    }

    /// <summary>
    /// 测试 - 自定义表头不能覆盖属性表头、数据行或其他自定义表头单元格。
    /// </summary>
    [Fact]
    public void Export_OverlappingCustomHeaderLayout_ShouldThrow()
    {
        // Arrange
        var data = new[] { new StreamRow { Name = "Value" } };

        // Act
        var propertyHeaderOverlap = () => new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(data,
            configure: sheet => sheet.HeaderRows(new[]
            {
                new ExcelHeaderRow(0, new[] { new ExcelHeaderCell(0, "总标题") })
            })), new MemoryStream());
        var dataRowOverlap = () => new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(data,
            configure: sheet => sheet.HeaderRowIndex(1).DataRowStartIndex(2).HeaderRows(new[]
            {
                new ExcelHeaderRow(0, new[] { new ExcelHeaderCell(0, "总标题", rowSpan: 2) })
            })), new MemoryStream());
        var customHeaderOverlap = () => new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(data,
            configure: sheet => sheet.HeaderRowIndex(2).DataRowStartIndex(3).HeaderRows(new[]
            {
                new ExcelHeaderRow(0, new[] { new ExcelHeaderCell(0, "总标题", columnSpan: 2) }),
                new ExcelHeaderRow(0, new[] { new ExcelHeaderCell(1, "副标题") })
            })), new MemoryStream());

        // Assert
        Assert.Throws<ArgumentException>(propertyHeaderOverlap);
        Assert.Throws<ArgumentException>(dataRowOverlap);
        Assert.Throws<ArgumentException>(customHeaderOverlap);
    }

    /// <summary>
    /// 测试 - 自定义表头集合中的空元素应在创建工作簿前被明确拒绝。
    /// </summary>
    [Fact]
    public void Export_CustomHeaderLayoutContainingNullElement_ShouldThrow()
    {
        // Arrange
        var data = new[] { new StreamRow { Name = "Value" } };

        // Act
        var nullRow = () => new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(data,
            configure: sheet => sheet.HeaderRows(new ExcelHeaderRow[] { null })), new MemoryStream());
        var nullCell = () => new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(data,
            configure: sheet => sheet.HeaderRows(new[]
            {
                new ExcelHeaderRow(0, new ExcelHeaderCell[] { null })
            })), new MemoryStream());

        // Assert
        Assert.Throws<ArgumentException>(nullRow);
        Assert.Throws<ArgumentException>(nullCell);
    }

    /// <summary>
    /// 测试 - 标记为连续合并的固定列应合并相邻相同且非空的值。
    /// </summary>
    [Fact]
    public void Export_MergeColumnsAttribute_ShouldMergeAdjacentValues()
    {
        // Arrange
        using var destination = new MemoryStream();
        var data = new[]
        {
            new MergeRow { Category = "A" },
            new MergeRow { Category = "A" },
            new MergeRow { Category = "B" }
        };

        // Act
        new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(data), destination);

        // Assert
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var mergedRegion = Assert.Single(workbook.GetSheetAt(0).MergedRegions);
        Assert.Equal(1, mergedRegion.FirstRow);
        Assert.Equal(2, mergedRegion.LastRow);
        Assert.Equal(0, mergedRegion.FirstColumn);
        Assert.Equal(0, mergedRegion.LastColumn);
    }

    /// <summary>
    /// 测试 - 表头与自动换行特性应写入真实工作簿的字体和单元格样式。
    /// </summary>
    [Fact]
    public void Export_HeaderAndWrapTextAttributes_ShouldApplyWorkbookStyles()
    {
        // Arrange
        using var destination = new MemoryStream();

        // Act
        new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(
            new[] { new StyledExportRow { Name = "第一行\n第二行" } }), destination);

        // Assert
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var sheet = workbook.GetSheetAt(0);
        var headerCell = sheet.GetRow(0).GetCell(0);
        var contentCell = sheet.GetRow(1).GetCell(0);
        var headerFont = workbook.GetFontAt(headerCell.CellStyle.FontIndex);
        Assert.Equal("Arial", headerFont.FontName);
        Assert.Equal(14, headerFont.FontHeightInPoints);
        Assert.False(headerFont.IsBold);
        Assert.True(headerCell.CellStyle.WrapText);
        Assert.True(contentCell.CellStyle.WrapText);
    }

    /// <summary>
    /// 测试 - 导出器应以日期单元格和格式样式写出日期，并支持 Xls 格式。
    /// </summary>
    [Fact]
    public void Export_DateAndXlsFormat_ShouldWriteTypedDateWithConfiguredStyle()
    {
        // Arrange
        using var destination = new MemoryStream();
        var date = new DateTime(2026, 8, 13, 14, 15, 16, DateTimeKind.Unspecified);

        // Act
        new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(
            new[] { new DateRow { OccurredAt = date } }, format: ExcelFormat.Xls), destination);

        // Assert
        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        var cell = workbook.GetSheetAt(0).GetRow(1).GetCell(0);
        Assert.IsType<HSSFWorkbook>(workbook);
        Assert.Equal(CellType.Numeric, cell.CellType);
        Assert.True(DateUtil.IsCellDateFormatted(cell));
        Assert.Equal(date, cell.DateCellValue);
        Assert.Contains("yyyy", cell.CellStyle.GetDataFormatString());
    }

    /// <summary>
    /// 测试 - 工作簿扩展应识别格式、排除隐藏工作表，并以非过时 API 设置粗体。
    /// </summary>
    [Fact]
    public void WorkbookExtensions_VisibleSheetsAndFontWeight_ShouldPreserveExpectedBehavior()
    {
        // Arrange
        using var workbook = ExcelHelper.PrepareWorkbook(ExcelFormat.Xls);
        workbook.CreateSheet("Visible");
        workbook.CreateSheet("Hidden");
        workbook.SetSheetHidden(1, SheetVisibility.Hidden);
        var font = workbook.CreateFont();

        // Act
        font.SetBoldWeight(700);
        var visibleSheets = workbook.GetSheets().ToList();

        // Assert
        Assert.Equal(ExcelFormat.Xls, workbook.GetExcelFormat());
        Assert.True(font.IsBold);
        Assert.Single(visibleSheets);
        Assert.Equal("Visible", visibleSheets[0].SheetName);
    }

    /// <summary>
    /// 测试 - 添加工作表时应按表头序列位置写入，包括重复表头。
    /// </summary>
    [Fact]
    public void WorkbookExtensions_AddSheetWithDuplicateHeaders_ShouldPreservePositions()
    {
        // Arrange
        using var workbook = ExcelHelper.PrepareWorkbook(ExcelFormat.Xlsx);

        // Act
        var sheet = workbook.AddSheet("Data", new List<string> { "编号", "编号", "名称" });

        // Assert
        var header = sheet.GetRow(0);
        Assert.Equal("编号", header.GetCell(0).StringCellValue);
        Assert.Equal("编号", header.GetCell(1).StringCellValue);
        Assert.Equal("名称", header.GetCell(2).StringCellValue);
    }

    /// <summary>
    /// 测试 - 删除包含稀疏空行的区间时应跳过缺失物理行并上移后续行。
    /// </summary>
    [Fact]
    public void SheetExtensions_RemoveRowsWithSparseRange_ShouldShiftFollowingRows()
    {
        // Arrange
        using var workbook = ExcelHelper.PrepareWorkbook(ExcelFormat.Xlsx);
        var sheet = workbook.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("First");
        sheet.CreateRow(3).CreateCell(0).SetCellValue("Following");

        // Act
        var removedCount = sheet.RemoveRows(1, 2);

        // Assert
        Assert.Equal(2, removedCount);
        Assert.Equal("First", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Following", sheet.GetRow(1).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 测试 - 行插入、删除和移动操作应拒绝无效索引、行数及反向区间。
    /// </summary>
    [Fact]
    public void SheetExtensions_InvalidRowOperationArguments_ShouldThrow()
    {
        // Arrange
        using var workbook = ExcelHelper.PrepareWorkbook(ExcelFormat.Xlsx);
        var sheet = workbook.CreateSheet("Data");

        // Act
        var deleteNegativeIndex = () => sheet.DeleteRows(-1, 1);
        var deleteZeroCount = () => sheet.DeleteRows(0, 0);
        var insertNegativeIndex = () => sheet.InsertRows(-1, 1);
        var insertZeroCount = () => sheet.InsertRows(0, 0);
        var removeNegativeIndex = () => { sheet.RemoveRows(-1, 0); };
        var removeReversedRange = () => { sheet.RemoveRows(1, 0); };

        // Assert
        Assert.Throws<ArgumentOutOfRangeException>(deleteNegativeIndex);
        Assert.Throws<ArgumentOutOfRangeException>(deleteZeroCount);
        Assert.Throws<ArgumentOutOfRangeException>(insertNegativeIndex);
        Assert.Throws<ArgumentOutOfRangeException>(insertZeroCount);
        Assert.Throws<ArgumentOutOfRangeException>(removeNegativeIndex);
        Assert.Throws<ArgumentOutOfRangeException>(removeReversedRange);
    }

    /// <summary>
    /// 测试 - 合并区域移动到负坐标时应拒绝请求且不修改任何区域。
    /// </summary>
    [Fact]
    public void SheetExtensions_MoveMergedRegionsBeyondWorksheetStart_ShouldThrowWithoutMutation()
    {
        // Arrange
        using var workbook = ExcelHelper.PrepareWorkbook(ExcelFormat.Xlsx);
        var sheet = workbook.CreateSheet("Data");
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 0, 1));
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(3, 4, 3, 4));

        // Act
        var action = () => sheet.MoveMergedRegions(-1, 0);

        // Assert
        Assert.Throws<ArgumentOutOfRangeException>(action);
        var first = sheet.GetMergedRegion(0);
        var second = sheet.GetMergedRegion(1);
        Assert.Equal(0, first.FirstRow);
        Assert.Equal(0, first.FirstColumn);
        Assert.Equal(3, second.FirstRow);
        Assert.Equal(3, second.FirstColumn);
    }

    /// <summary>
    /// 测试 - 合并区域移动至既有区域时应拒绝请求且保持所有区域不变。
    /// </summary>
    [Fact]
    public void SheetExtensions_MoveMergedRegionsOverlappingExistingRegion_ShouldThrowWithoutMutation()
    {
        // Arrange
        using var workbook = ExcelHelper.PrepareWorkbook(ExcelFormat.Xlsx);
        var sheet = workbook.CreateSheet("Data");
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 0, 1));
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(3, 4, 0, 1));

        // Act
        var action = () => sheet.MoveMergedRegions(0, 1, 0, 1, 3);

        // Assert
        Assert.Throws<ArgumentException>(action);
        var first = sheet.GetMergedRegion(0);
        var second = sheet.GetMergedRegion(1);
        Assert.Equal(0, first.FirstRow);
        Assert.Equal(1, first.LastRow);
        Assert.Equal(3, second.FirstRow);
        Assert.Equal(4, second.LastRow);
    }

    /// <summary>
    /// 测试 - 插入行时应保留原始数据，并随 NPOI 行移动同步调整合并区域。
    /// </summary>
    [Fact]
    public void SheetExtensions_InsertRows_ShouldShiftExistingRowsAndMergedRegions()
    {
        // Arrange
        using var workbook = ExcelHelper.PrepareWorkbook(ExcelFormat.Xlsx);
        var sheet = workbook.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("Header");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("Value");
        sheet.CreateRow(2).CreateCell(0).SetCellValue("Merged");
        sheet.CreateRow(3);
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(2, 3, 0, 1));

        // Act
        var insertedRows = sheet.InsertRows(1, 2);

        // Assert
        Assert.Equal(2, insertedRows.Length);
        Assert.Equal("Header", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Value", sheet.GetRow(3).GetCell(0).StringCellValue);
        Assert.Equal("Merged", sheet.GetRow(4).GetCell(0).StringCellValue);
        var region = sheet.GetMergedRegion(0);
        Assert.Equal(4, region.FirstRow);
        Assert.Equal(5, region.LastRow);
        Assert.Equal(0, region.FirstColumn);
        Assert.Equal(1, region.LastColumn);
    }

    /// <summary>
    /// 测试 - 批量移除合并区域时应逆序删除索引，避免留下部分区域。
    /// </summary>
    [Fact]
    public void SheetExtensions_RemoveMergedRegions_ShouldRemoveAllRegions()
    {
        // Arrange
        using var workbook = ExcelHelper.PrepareWorkbook(ExcelFormat.Xlsx);
        var sheet = workbook.CreateSheet("Data");
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 0, 1));
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(3, 4, 0, 1));

        // Act
        sheet.RemoveMergedRegions();

        // Assert
        Assert.Equal(0, sheet.NumMergedRegions);
    }

    /// <summary>
    /// 测试 - 删除行时应解除与删除区间相交的跨行合并区域。
    /// </summary>
    [Fact]
    public void SheetExtensions_DeleteRowsIntersectingMergedRegion_ShouldRemoveRegion()
    {
        // Arrange
        using var workbook = ExcelHelper.PrepareWorkbook(ExcelFormat.Xlsx);
        var sheet = workbook.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue("First");
        sheet.CreateRow(1).CreateCell(0).SetCellValue("Last");
        sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(0, 1, 0, 0));

        // Act
        sheet.DeleteRows(1, 1);

        // Assert
        Assert.Equal(0, sheet.NumMergedRegions);
    }

    /// <summary>
    /// 测试 - 空表头应在创建工作表前被拒绝，工作簿状态保持不变。
    /// </summary>
    [Fact]
    public void WorkbookExtensions_AddSheetWithNullHeaders_ShouldThrowWithoutAddingSheet()
    {
        // Arrange
        using var workbook = ExcelHelper.PrepareWorkbook(ExcelFormat.Xlsx);

        // Act
        var action = () => workbook.AddSheet("Data", null);

        // Assert
        Assert.Throws<ArgumentNullException>(action);
        Assert.Equal(0, workbook.NumberOfSheets);
    }

    /// <summary>
    /// 测试 - 空图片样式应在注册图片资源前被拒绝，避免留下孤立图片。
    /// </summary>
    [Fact]
    public void SheetExtensions_AddPictureWithNullStyle_ShouldThrowWithoutAddingPicture()
    {
        // Arrange
        using var workbook = ExcelHelper.PrepareWorkbook(ExcelFormat.Xlsx);
        var sheet = workbook.CreateSheet("Data");
        var picture = new Bing.Offices.Metadata.PictureInfo(0, 1, 0, 1,
            Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw=="), null);

        // Act
        var action = () => sheet.AddPicture(picture);

        // Assert
        Assert.Throws<ArgumentException>(action);
        Assert.Empty(workbook.GetAllPictures());
    }

    /// <summary>
    /// 测试 - 未专门支持的非空类型应以不变区域性文本保留，而非静默写为空值。
    /// </summary>
    [Fact]
    public void CellExtensions_SetValueWithUnsupportedObject_ShouldPreserveText()
    {
        // Arrange
        using var workbook = ExcelHelper.PrepareWorkbook(ExcelFormat.Xlsx);
        var cell = workbook.CreateSheet("Data").CreateRow(0).CreateCell(0);

        // Act
        cell.SetValue(TimeSpan.FromMinutes(90));

        // Assert
        Assert.Equal("01:30:00", cell.StringCellValue);
    }

    /// <summary>
    /// 测试 - 取消令牌在开始处理前取消时，导入器和导出器都应立即终止。
    /// </summary>
    [Fact]
    public void StreamPipeline_PreCancelledToken_ShouldThrowOperationCanceledException()
    {
        // Arrange
        using var source = new MemoryStream(CreateWorkbook(_ => { }));
        using var destination = new MemoryStream();
        using var cancellationTokenSource = new CancellationTokenSource();
        cancellationTokenSource.Cancel();

        // Act
        var import = () => new NpoiExcelImporter().Import(source, CreateSingleSheetRequest<StreamRow>(),
            cancellationTokenSource.Token);
        var export = () => new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(
            Array.Empty<StreamRow>()), destination, cancellationTokenSource.Token);

        // Assert
        Assert.Throws<OperationCanceledException>(import);
        Assert.Throws<OperationCanceledException>(export);
    }

    /// <summary>
    /// 测试 - 流复制期间取消时，导入器应保持源流打开，导出器不应污染调用方目标流。
    /// </summary>
    [Fact]
    public void StreamPipeline_CancelDuringProcessing_ShouldPreserveCallerStreams()
    {
        // Arrange
        using var importCancellationTokenSource = new CancellationTokenSource();
        using var exportCancellationTokenSource = new CancellationTokenSource();
        using var source = new NonSeekableReadStream(CreateWorkbook(workbook =>
        {
            var sheet = workbook.CreateSheet("Data");
            sheet.CreateRow(0).CreateCell(0).SetCellValue(nameof(StreamRow.Name));
            sheet.CreateRow(1).CreateCell(0).SetCellValue("Value");
        }), importCancellationTokenSource.Cancel);
        var originalBytes = new byte[] { 1, 2, 3 };
        using var destination = new MemoryStream(originalBytes.ToArray());
        destination.Position = destination.Length;

        // Act
        var import = () => new NpoiExcelImporter().Import(source, CreateSingleSheetRequest<StreamRow>(),
            importCancellationTokenSource.Token);
        var export = () => new NpoiExcelExporter().Export(CreateSingleSheetExportRequest(
            CancelAfterFirstItem(exportCancellationTokenSource)), destination,
            exportCancellationTokenSource.Token);

        // Assert
        Assert.Throws<OperationCanceledException>(import);
        Assert.True(source.CanRead);
        Assert.Throws<OperationCanceledException>(export);
        Assert.True(destination.CanWrite);
        Assert.Equal(originalBytes, destination.ToArray());
    }

    /// <summary>
    /// 测试 - 公共 API 应固定为以调用方流为所有者的导入导出契约，且不再暴露旧服务类型。
    /// </summary>
    [Fact]
    public void PublicApi_ShouldExposeOnlyStreamFirstContracts()
    {
        // Arrange
        var abstractionsAssembly = typeof(IExcelImporter).Assembly;

        // Act
        var importer = typeof(IExcelImporter).GetMethods().Single(method =>
            method.GetParameters().Length == 3
            && method.GetParameters()[1].ParameterType.Name.StartsWith("ExcelWorkbookImportRequest", StringComparison.Ordinal));
        var exporter = typeof(IExcelExporter).GetMethods().Single(method =>
            method.GetParameters().Length == 3
            && method.GetParameters()[0].ParameterType == typeof(Bing.Offices.Exports.ExcelWorkbookExportRequest));

        // Assert
        Assert.NotNull(importer);
        Assert.NotNull(exporter);
        Assert.Equal(typeof(Stream), importer.GetParameters()[0].ParameterType);
        Assert.Equal(typeof(Stream), exporter.GetParameters()[1].ParameterType);
        Assert.Null(abstractionsAssembly.GetType("Bing.Offices.Imports.IExcelImportService"));
        Assert.Null(abstractionsAssembly.GetType("Bing.Offices.Imports.IExcelImportProvider"));
        Assert.Null(abstractionsAssembly.GetType("Bing.Offices.Exports.IExcelExportService"));
        Assert.Null(abstractionsAssembly.GetType("Bing.Offices.Exports.IExcelExportProvider"));
        Assert.Null(abstractionsAssembly.GetType("Bing.Offices.Exports.IExportOptions`1"));
        Assert.Null(abstractionsAssembly.GetType("Bing.Offices.Filters.IFilter"));
        Assert.Null(abstractionsAssembly.GetType("Bing.Offices.Metadata.Excels.IWorkbook"));
        Assert.Null(abstractionsAssembly.GetType("Bing.Offices.Metadata.Excels.IWorkSheet"));
        Assert.Null(abstractionsAssembly.GetType("Bing.Offices.Metadata.Excels.IRow"));
        Assert.Null(abstractionsAssembly.GetType("Bing.Offices.Metadata.Excels.ICell"));
        Assert.Null(abstractionsAssembly.GetType("Bing.Offices.Metadata.Excels.IRange"));
    }

    /// <summary>
    /// 创建测试工作簿字节。
    /// </summary>
    /// <param name="configure">配置工作簿的操作。</param>
    /// <returns>工作簿字节。</returns>
    private static byte[] CreateWorkbook(Action<XSSFWorkbook> configure) => CreateWorkbook<XSSFWorkbook>(configure);

    private static ExcelWorkbookImportRequest<SingleWorkbook<T>> CreateSingleSheetRequest<T>(
        Action<ExcelSheetImportBuilder<T>> configure = null) where T : class, new()
    {
        return ExcelImport.Workbook<SingleWorkbook<T>>(builder =>
            builder.Sheet("Data", root => root.Items, configure));
    }

    private static ExcelWorkbookExportRequest CreateSingleSheetExportRequest<T>(IEnumerable<T> data,
        string name = "Data", ExcelFormat format = ExcelFormat.Xlsx,
        Action<ExcelSheetExportBuilder<T>> configure = null) where T : class, new()
    {
        return ExcelExport.Workbook(builder => builder.Format(format).AddSheet(name, data, configure));
    }

    /// <summary>
    /// 断言两个列定义具有相同的映射语义。
    /// </summary>
    private static void AssertEquivalentColumn(ExcelPropertyMap expected, ExcelPropertyMap actual)
    {
        Assert.Equal(expected.Title, actual.Title);
        Assert.Equal(expected.Formatter, actual.Formatter);
        Assert.Equal(expected.DecimalScale, actual.DecimalScale);
        Assert.Equal(expected.ValueMap, actual.ValueMap);
    }

    /// <summary>
    /// 创建指定 NPOI 工作簿类型的测试字节。
    /// </summary>
    /// <typeparam name="TWorkbook">NPOI 工作簿类型。</typeparam>
    /// <param name="configure">配置工作簿的操作。</param>
    /// <returns>工作簿字节。</returns>
    private static byte[] CreateWorkbook<TWorkbook>(Action<TWorkbook> configure) where TWorkbook : IWorkbook, new()
    {
        using var workbook = new TWorkbook();
        configure(workbook);
        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        return stream.ToArray();
    }

    /// <summary>
    /// 创建包含重复校验列的测试工作表。
    /// </summary>
    /// <param name="workbook">测试工作簿。</param>
    /// <param name="name">工作表名称。</param>
    /// <param name="code">业务编码。</param>
    private static void CreateValidatedSheet(XSSFWorkbook workbook, string name, string code)
    {
        var sheet = workbook.CreateSheet(name);
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue(nameof(ValidatedRow.Code));
        header.CreateCell(1).SetCellValue(nameof(ValidatedRow.Count));
        var row = sheet.CreateRow(1);
        row.CreateCell(0).SetCellValue(code);
        row.CreateCell(1).SetCellValue(1);
    }

    /// <summary>
    /// 从导出工作簿读取指定单元格文本。
    /// </summary>
    private static string GetWorkbookCellValue(Stream source, int rowIndex, int columnIndex)
    {
        var position = source.Position;
        source.Position = 0;
        using var snapshot = new MemoryStream();
        source.CopyTo(snapshot);
        source.Position = position;
        snapshot.Position = 0;
        using var workbook = WorkbookFactory.Create(snapshot);
        return workbook.GetSheetAt(0).GetRow(rowIndex).GetCell(columnIndex).StringCellValue;
    }

    /// <summary>
    /// 在枚举第一个数据项后取消令牌。
    /// </summary>
    /// <param name="cancellationTokenSource">待取消的令牌源。</param>
    /// <returns>导出数据序列。</returns>
    private static IEnumerable<StreamRow> CancelAfterFirstItem(CancellationTokenSource cancellationTokenSource)
    {
        yield return new StreamRow { Name = "First" };
        cancellationTokenSource.Cancel();
        yield return new StreamRow { Name = "Second" };
    }

    private sealed class SingleWorkbook<T> where T : class, new()
    {
        public List<T> Items { get; } = new();
    }

    private sealed class MultiSheetWorkbook<T> where T : class, new()
    {
        public List<T> Items { get; } = new();
    }

    /// <summary>
    /// 流式导入导出测试模型。
    /// </summary>
    private class StreamRow
    {
        /// <summary>
        /// 名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 带导入校验的测试模型。
    /// </summary>
    private class ValidatedRow
    {
        /// <summary>
        /// 业务编码。
        /// </summary>
        [Required]
        [Duplication]
        public string Code { get; set; }

        /// <summary>
        /// 数量。
        /// </summary>
        public int Count { get; set; }
    }

    /// <summary>
    /// 包含自定义校验规则的测试模型。
    /// </summary>
    private class CustomValidationRow
    {
        /// <summary>
        /// 业务编码。
        /// </summary>
        [StartsWithAttribute]
        public string Code { get; set; }
    }

    /// <summary>
    /// 包含领域值对象的转换测试模型。
    /// </summary>
    private class ConvertedRow
    {
        /// <summary>
        /// 订单编码。
        /// </summary>
        public OrderCode Code { get; set; }
    }

    /// <summary>
    /// 请求级映射配置测试模型。
    /// </summary>
    private class ConfiguredRow
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
    /// 订单编码领域值对象。
    /// </summary>
    private sealed class OrderCode
    {
        /// <summary>
        /// 初始化一个<see cref="OrderCode"/>类型的实例。
        /// </summary>
        /// <param name="value">订单编码文本。</param>
        public OrderCode(string value) => Value = value;

        /// <summary>
        /// 获取订单编码文本。
        /// </summary>
        public string Value { get; }
    }

    /// <summary>
    /// 订单编码双向转换器。
    /// </summary>
    private sealed class OrderCodeExcelValueConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "order-code";

        /// <summary>
        /// 获取导入时收到的属性类型。
        /// </summary>
        public Type ImportPropertyType { get; private set; }

        /// <summary>
        /// 获取导出时收到的属性类型。
        /// </summary>
        public Type ExportPropertyType { get; private set; }

        /// <summary>
        /// 获取导入时收到的行号。
        /// </summary>
        public int ImportRowIndex { get; private set; }

        /// <summary>
        /// 获取导出时收到的行号。
        /// </summary>
        public int ExportRowIndex { get; private set; }

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(OrderCode);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            ImportPropertyType = context.PropertyType;
            ImportRowIndex = context.RowIndex;
            value = new OrderCode(((string)context.Value).Substring(3));
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            ExportPropertyType = context.PropertyType;
            ExportRowIndex = context.RowIndex;
            value = $"CV-{((OrderCode)context.Value).Value}";
            return true;
        }
    }

    /// <summary>
    /// 捕获公式单元格抽象上下文的转换器。
    /// </summary>
    private sealed class FormulaContextExcelValueConverter : IExcelValueConverter
    {
        /// <summary>
        /// 获取单元格逻辑类型。
        /// </summary>
        public ExcelCellKind Kind { get; private set; }

        /// <summary>
        /// 获取公式缓存类型。
        /// </summary>
        public ExcelCellKind? CachedKind { get; private set; }

        /// <summary>
        /// 获取公式文本。
        /// </summary>
        public string Formula { get; private set; }

        /// <summary>
        /// 获取原始 typed 值。
        /// </summary>
        public object RawValue { get; private set; }

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(decimal);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            Kind = context.Cell.Kind;
            CachedKind = context.Cell.CachedKind;
            Formula = context.Cell.Formula;
            RawValue = context.Cell.Value;
            value = null;
            return false;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = null;
            return false;
        }
    }

    /// <summary>
    /// 历史文本转换器测试实现。
    /// </summary>
#pragma warning disable CS0618
    private sealed class PrefixLegacyCellValueConverter : ICellValueConverter
    {
        /// <inheritdoc />
        public string GetStringValue(object cell) => $"adapted-{((ICell)cell).GetStringValue()}";
    }
#pragma warning restore CS0618

    /// <summary>
    /// 映射配置加载器替换测试实现。
    /// </summary>
    private sealed class TestMappingConfigurationLoader : IExcelMappingConfigurationLoader
    {
        /// <inheritdoc />
        /// <inheritdoc />
        public ExcelMappingDocument FromJsonDocument(string json) => ExcelMappingConfigurationLoader.FromJsonDocument(json);

        /// <inheritdoc />
        public ExcelMappingDocument FromJsonDocument(Stream source) => ExcelMappingConfigurationLoader.FromJsonDocument(source);

        /// <inheritdoc />
        /// <inheritdoc />
        public ExcelMappingDocument FromXmlDocument(string xml) => ExcelMappingConfigurationLoader.FromXmlDocument(xml);

        /// <inheritdoc />
        public ExcelMappingDocument FromXmlDocument(Stream source) => ExcelMappingConfigurationLoader.FromXmlDocument(source);
    }

    /// <summary>
    /// 命名前缀校验规则。
    /// </summary>
    private sealed class StartsWithOkValidationRule : INamedExcelValidationRule
    {
        /// <inheritdoc />
        public string Name => "starts-with-ok";

        /// <inheritdoc />
        public string ErrorMessage => "必须以 OK- 开头";

        /// <inheritdoc />
        public bool Validate(ExcelValidationContext context) => context.Value.StartsWith("OK-", StringComparison.Ordinal);
    }

    /// <summary>
    /// 捕获命名校验规则上下文的测试实现。
    /// </summary>
    private sealed class ContextCapturingValidationRule : INamedExcelValidationRule
    {
        /// <inheritdoc />
        public string Name => "context";

        /// <inheritdoc />
        public string ErrorMessage => "上下文错误";

        /// <summary>
        /// 获取单元格逻辑类型。
        /// </summary>
        public ExcelCellKind CellKind { get; private set; }

        /// <summary>
        /// 获取区域性名称。
        /// </summary>
        public string CultureName { get; private set; }

        /// <summary>
        /// 获取属性名称。
        /// </summary>
        public string PropertyName { get; private set; }

        /// <inheritdoc />
        public bool Validate(ExcelValidationContext context)
        {
            CellKind = context.Cell.Kind;
            CultureName = context.Culture.Name;
            PropertyName = context.PropertyName;
            return true;
        }
    }

    /// <summary>
    /// 包含未绑定校验特性的测试模型。
    /// </summary>
    private class UnboundValidationRow
    {
        /// <summary>
        /// 业务编码。
        /// </summary>
        [UnboundValidationAttribute]
        public string Code { get; set; }
    }

    /// <summary>
    /// 要求文本使用指定前缀的自定义校验特性。
    /// </summary>
    [BindFilter(typeof(StartsWithExcelValidationRule))]
    private sealed class StartsWithAttribute : FilterAttributeBase
    {
        /// <inheritdoc />
        public override string ErrorMsg { get; set; } = "必须以 OK- 开头";
    }

    /// <summary>
    /// 未绑定规则的自定义校验特性。
    /// </summary>
    private sealed class UnboundValidationAttribute : FilterAttributeBase
    {
    }

    /// <summary>
    /// 会抛出异常的校验特性。
    /// </summary>
    [BindFilter(typeof(ThrowingExcelValidationRule))]
    private sealed class ThrowingValidationAttribute : FilterAttributeBase
    {
    }

    /// <summary>
    /// 自定义前缀校验规则。
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
    /// 会抛出异常的校验规则。
    /// </summary>
    private sealed class ThrowingExcelValidationRule : IExcelValidationRule
    {
        /// <inheritdoc />
        public bool CanValidate(FilterAttributeBase attribute) => attribute is ThrowingValidationAttribute;

        /// <inheritdoc />
        public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context) =>
            throw new InvalidOperationException("校验器异常");
    }

    /// <summary>
    /// 包含会抛出异常的校验特性的测试模型。
    /// </summary>
    private class ThrowingValidationRow
    {
        /// <summary>
        /// 业务编码。
        /// </summary>
        [ThrowingValidationAttribute]
        public string Code { get; set; }
    }

    /// <summary>
    /// 数值转换测试模型。
    /// </summary>
    private class NumberRow
    {
        /// <summary>
        /// 金额。
        /// </summary>
        public decimal Amount { get; set; }
    }

    /// <summary>
    /// 包含格式化数值的往返测试模型。
    /// </summary>
    private class FormattedNumberRow
    {
        /// <summary>
        /// 金额。
        /// </summary>
        [DataFormat("0.00")]
        public decimal Amount { get; set; }
    }

    /// <summary>
    /// 公式缓存类型转换测试模型。
    /// </summary>
    private class FormulaRow
    {
        /// <summary>
        /// 金额。
        /// </summary>
        public decimal Amount { get; set; }

        /// <summary>
        /// 发生日期。
        /// </summary>
        public DateTime OccurredAt { get; set; }

        /// <summary>
        /// 可选发生日期。
        /// </summary>
        public DateTime? OptionalOccurredAt { get; set; }

        /// <summary>
        /// 是否启用。
        /// </summary>
        public bool Enabled { get; set; }

        /// <summary>
        /// 名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 转换错误定位测试模型。
    /// </summary>
    private class ConversionErrorRow
    {
        /// <summary>
        /// 数量。
        /// </summary>
        public int Count { get; set; }

        /// <summary>
        /// 唯一标识。
        /// </summary>
        public Guid Id { get; set; }
    }

    /// <summary>
    /// 标量类型往返测试模型。
    /// </summary>
    private class ScalarRow
    {
        /// <summary>
        /// 唯一标识。
        /// </summary>
        public Guid Id { get; set; }

        /// <summary>
        /// 可选唯一标识。
        /// </summary>
        public Guid? OptionalId { get; set; }

        /// <summary>
        /// 版本。
        /// </summary>
        public Version Version { get; set; }

        /// <summary>
        /// 是否启用。
        /// </summary>
        public bool Enabled { get; set; }

        /// <summary>
        /// 长整型状态。
        /// </summary>
        public LongStatus Status { get; set; }

        /// <summary>
        /// 访问权限。
        /// </summary>
        public AccessFlags Access { get; set; }
    }

    /// <summary>
    /// 包含重复显示文本映射的测试模型。
    /// </summary>
    private class DuplicateTextMappingRow
    {
        /// <summary>
        /// 状态。
        /// </summary>
        [ValueMapping("相同", 1)]
        [ValueMapping("相同", 2)]
        public int Status { get; set; }
    }

    /// <summary>
    /// 包含重复业务值映射的测试模型。
    /// </summary>
    private class DuplicateValueMappingRow
    {
        /// <summary>
        /// 状态。
        /// </summary>
        [ValueMapping("第一个", 1)]
        [ValueMapping("第二个", 1)]
        public int Status { get; set; }
    }

    /// <summary>
    /// 长整型底层枚举。
    /// </summary>
    private enum LongStatus : long
    {
        /// <summary>
        /// 已归档。
        /// </summary>
        Archived = 9007199254740992
    }

    /// <summary>
    /// 访问权限枚举。
    /// </summary>
    [Flags]
    private enum AccessFlags
    {
        /// <summary>
        /// 无权限。
        /// </summary>
        None = 0,

        /// <summary>
        /// 读取权限。
        /// </summary>
        Read = 1,

        /// <summary>
        /// 写入权限。
        /// </summary>
        Write = 2
    }

    /// <summary>
    /// 日期导出测试模型。
    /// </summary>
    private class DateRow
    {
        /// <summary>
        /// 发生时间。
        /// </summary>
        public DateTime OccurredAt { get; set; }
    }

    /// <summary>
    /// 配置来源等价测试模型。
    /// </summary>
    private class EquivalentConfigurationRow
    {
        /// <summary>
        /// 金额。
        /// </summary>
        [ColumnName("金额")]
        [DataFormat("0.00")]
        [DecimalScale(2)]
        [ValueMapping("有效", 1)]
        public int Amount { get; set; }
    }

    /// <summary>
    /// 带动态列的测试模型。
    /// </summary>
    private class DynamicRow
    {
        /// <summary>
        /// 业务编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 动态列数据。
        /// </summary>
        [DynamicColumn]
        public IDictionary<string, object> Values { get; set; }
    }

    /// <summary>
    /// 包含不兼容动态列类型的测试模型。
    /// </summary>
    private class InvalidDynamicRow
    {
        /// <summary>
        /// 动态列数据。
        /// </summary>
        [DynamicColumn]
        public IDictionary<string, string> Values { get; set; }
    }

    /// <summary>
    /// 包含只读属性的导入测试模型。
    /// </summary>
    private class ReadOnlyRow
    {
        /// <summary>
        /// 业务编码。
        /// </summary>
        public string Code { get; } = "Default";
    }

    /// <summary>
    /// 表头映射校验测试模型。
    /// </summary>
    private class MappingRow
    {
        /// <summary>
        /// 第一个属性。
        /// </summary>
        public string First { get; set; }

        /// <summary>
        /// 第二个属性。
        /// </summary>
        public string Second { get; set; }
    }

    /// <summary>
    /// 包含无效多动态列声明的测试模型。
    /// </summary>
    private class MultipleDynamicRow
    {
        /// <summary>
        /// 第一个动态列。
        /// </summary>
        [DynamicColumn]
        public IDictionary<string, object> First { get; set; }

        /// <summary>
        /// 第二个动态列。
        /// </summary>
        [DynamicColumn]
        public IDictionary<string, object> Second { get; set; }
    }

    /// <summary>
    /// 连续合并测试模型。
    /// </summary>
    private class MergeRow
    {
        /// <summary>
        /// 分类。
        /// </summary>
        [MergeColumns]
        public string Category { get; set; }
    }

    /// <summary>
    /// 包含表头与自动换行特性的导出测试模型。
    /// </summary>
    [Header(FontName = "Arial", FontSize = 14, Bold = false)]
    [WrapText]
    private class StyledExportRow
    {
        /// <summary>
        /// 名称。
        /// </summary>
        public string Name { get; set; }
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
        /// 读取数据块后的回调。
        /// </summary>
        private readonly Action _afterRead;

        /// <summary>
        /// 初始化一个<see cref="NonSeekableReadStream"/>类型的实例。
        /// </summary>
        /// <param name="bytes">输入字节。</param>
        /// <param name="afterRead">读取数据块后的回调。</param>
        public NonSeekableReadStream(byte[] bytes, Action afterRead = null)
        {
            _inner = new MemoryStream(bytes);
            _afterRead = afterRead;
        }

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
        public override void Flush() => throw new NotSupportedException();

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count)
        {
            var readCount = _inner.Read(buffer, offset, count);
            if (readCount > 0)
                _afterRead?.Invoke();
            return readCount;
        }

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

    /// <summary>
    /// 不支持定位的只写内存流。
    /// </summary>
    private sealed class NonSeekableWriteStream : Stream
    {
        /// <summary>
        /// 被代理的内存流。
        /// </summary>
        private readonly MemoryStream _inner = new();

        /// <inheritdoc />
        public override bool CanRead => false;

        /// <inheritdoc />
        public override bool CanSeek => false;

        /// <inheritdoc />
        public override bool CanWrite => true;

        /// <inheritdoc />
        public override long Length => throw new NotSupportedException();

        /// <inheritdoc />
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        /// <inheritdoc />
        public override void Flush() => _inner.Flush();

        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);

        /// <summary>
        /// 获取当前写入的数据副本。
        /// </summary>
        /// <returns>已写入的字节。</returns>
        public byte[] ToArray() => _inner.ToArray();

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
