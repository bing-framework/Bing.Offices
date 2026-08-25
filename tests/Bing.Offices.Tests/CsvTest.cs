using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Globalization;
using System.Text;
using Bing.Offices.Csv;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Validations;
using Xunit;

namespace Bing.Offices.Tests;

public class CsvTest
{
    [Fact]
    public void Test_ExportByDataTable()
    {
        var dt = new DataTable();
        dt.Columns.AddRange(new[]
        {
            new DataColumn("Name"),
            new DataColumn("Age"),
            new DataColumn("Desc"),
            new DataColumn("Separator"),
            new DataColumn("Quote"),
        });
        for (var i = 0; i < 10; i++)
        {
            var row = dt.NewRow();
            row.ItemArray = new object[] { $"Test_{i}", i + 10, $"Desc_{i}",$"Separator , {i}",$"Quote , \" {i}" };
            dt.Rows.Add(row);
        }

        var filePath = Path.Combine(Path.GetTempPath(), $"Bing.Offices.Tests.{Guid.NewGuid():N}.csv");
        try
        {
            Assert.True(CsvHelper.ToCsvFile(dt, filePath, true, ',', '"'));
            var content = File.ReadAllText(filePath, System.Text.Encoding.UTF8);
            Assert.Contains("Name,Age,Desc,Separator,Quote", content);
            Assert.Contains("\"Separator , 0\"", content);
            Assert.Contains("\"Quote , \"\" 0\"", content);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

    /// <summary>
    /// 测试 - DataTable CSV 兼容层应为无数据表保留表头，并按显式参数处理公式字段和引用字符。
    /// </summary>
    [Fact]
    public void DataTableCompatibility_ExplicitOptions_ShouldKeepHeaderAndEscapeFormula()
    {
        // Arrange
        var empty = new DataTable();
        empty.Columns.Add("=标题");
        var data = new DataTable();
        data.Columns.Add("Name");
        data.Rows.Add("=SUM(A1:A2)");

        // Act
        var emptyContent = CsvHelper.GetCsvText(empty, true, ';', '|');
        var content = CsvHelper.GetCsvText(data, true, ';', '|');

        // Assert
        Assert.Equal("'=标题\r\n", emptyContent);
        Assert.Contains("'=SUM(A1:A2)", content);
        Assert.DoesNotContain("\"", content);
    }

    /// <summary>
    /// 测试 - 实体 CSV 管线应正确转义分隔符、引号和记录内换行，并保持调用方流打开。
    /// </summary>
    [Fact]
    public void EntityPipeline_EscapedValues_ShouldRoundTripAndKeepStreamsOpen()
    {
        // Arrange
        var exporter = new CsvEntityExporter();
        var importer = new CsvEntityImporter();
        using var destination = new MemoryStream();
        var sourceItems = new[]
        {
            new CsvRow { Name = "A,\"B\"", Count = 7, Description = "第一行\r\n第二行" }
        };

        // Act
        exporter.Export(sourceItems, destination);
        var content = Encoding.UTF8.GetString(destination.ToArray());
        destination.Position = 0;
        var result = importer.Import<CsvRow>(destination);

        // Assert
        Assert.Contains("\"A,\"\"B\"\"\"", content);
        Assert.True(destination.CanRead);
        Assert.True(destination.CanWrite);
        Assert.Empty(result.Errors);
        var item = Assert.Single(result.Items);
        Assert.Equal(sourceItems[0].Name, item.Name);
        Assert.Equal(sourceItems[0].Description, item.Description);
        Assert.Equal(7, item.Count);
    }

    /// <summary>
    /// 测试 - CSV 应按请求写入 UTF-8 BOM，保留空值，并可从不可寻址输入流读取且不关闭调用方流。
    /// </summary>
    [Fact]
    public void EntityPipeline_BomNullValuesAndNonSeekableStream_ShouldRoundTripAndKeepStreamOpen()
    {
        // Arrange
        using var destination = new MemoryStream();
        var encoding = new UTF8Encoding(true);

        // Act
        new CsvEntityExporter().Export(new[]
        {
            new CsvRow { Name = null, Count = 1, Description = null }
        }, destination, new CsvExportOptions<CsvRow> { Encoding = encoding });
        using var source = new NonSeekableReadStream(destination.ToArray());
        var result = new CsvEntityImporter().Import<CsvRow>(source,
            new CsvImportOptions<CsvRow> { Encoding = encoding });

        // Assert
        var content = destination.ToArray();
        Assert.Equal(0xEF, content[0]);
        Assert.Equal(0xBB, content[1]);
        Assert.Equal(0xBF, content[2]);
        Assert.Empty(result.Errors);
        var item = Assert.Single(result.Items);
        Assert.Null(item.Name);
        Assert.Equal(1, item.Count);
        Assert.Null(item.Description);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 测试 - CSV 转换失败应返回行列和属性定位，且不返回失败行。
    /// </summary>
    [Fact]
    public void EntityPipeline_InvalidValue_ShouldReturnStructuredError()
    {
        // Arrange
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("Name,Count,Description\r\n有效,invalid,说明\r\n"));

        // Act
        var result = new CsvEntityImporter().Import<CsvRow>(source);

        // Assert
        Assert.Empty(result.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(2, error.ColumnIndex);
        Assert.Equal(nameof(CsvRow.Count), error.PropertyName);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 测试 - CSV 字节数组和路径便利扩展应委托流式实体管线完成往返。
    /// </summary>
    [Fact]
    public void StreamExtensions_BytesAndFile_ShouldRoundTripEntity()
    {
        // Arrange
        var exporter = new CsvEntityExporter();
        var importer = new CsvEntityImporter();
        var source = new[] { new CsvRow { Name = "兼容", Count = 2, Description = "路径" } };
        var filePath = Path.Combine(Path.GetTempPath(), $"Bing.Offices.Tests.{Guid.NewGuid():N}.csv");

        try
        {
            // Act
            var fromBytes = importer.ImportFromBytes<CsvRow>(exporter.ExportToBytes(source));
            exporter.ExportToFile(source, filePath);
            var fromFile = importer.ImportFromFile<CsvRow>(filePath);

            // Assert
            Assert.Empty(fromBytes.Errors);
            Assert.Empty(fromFile.Errors);
            Assert.Equal("兼容", Assert.Single(fromBytes.Items).Name);
            Assert.Equal("路径", Assert.Single(fromFile.Items).Description);
        }
        finally
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
    }

        /// <summary>
        /// 测试 - 同步字节数组入口应委托当前 CSV Stream-first 导出器。
        /// </summary>
        [Fact]
        public void StreamExtensions_Bytes_ShouldDelegateToExporter()
        {
        // Arrange
        var exporter = new CsvEntityExporter();

        // Act
        var content = exporter.ExportToBytes(new[] { new CsvRow { Name = "兼容", Count = 1 } });

        // Assert
        Assert.NotEmpty(content);
        Assert.Contains("兼容", Encoding.UTF8.GetString(content));
        }

    /// <summary>
    /// 测试 - CSV 请求映射应按转换器名称使用已提供的双向转换器。
    /// </summary>
    [Fact]
    public void EntityPipeline_NamedConverterConfiguration_ShouldRoundTripDomainValue()
    {
        // Arrange
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new ExcelColumnConfiguration { PropertyName = nameof(CsvConvertedRow.Code), ConverterName = "csv-code" }
            }
        };
        var converter = new CsvCodeConverter();
        using var destination = new MemoryStream();
        var exporter = new CsvEntityExporter(new IExcelValueConverter[] { converter });
        var importer = new CsvEntityImporter(new IExcelValueConverter[] { converter });

        // Act
        exporter.Export(new[] { new CsvConvertedRow { Code = new CsvCode("42") } }, destination,
            new CsvExportOptions<CsvConvertedRow> { MappingConfiguration = mapping });
        destination.Position = 0;
        var result = importer.Import<CsvConvertedRow>(destination,
            new CsvImportOptions<CsvConvertedRow> { MappingConfiguration = mapping });

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal("42", Assert.Single(result.Items).Code.Value);
    }

    /// <summary>
    /// 测试 - CSV 请求映射应使用已提供的命名校验规则并返回结构化错误。
    /// </summary>
    [Fact]
    public void EntityPipeline_NamedValidationRule_ShouldReturnStructuredError()
    {
        // Arrange
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(CsvRow.Name),
                    ValidationRuleNames = new List<string> { "starts-with-ok" }
                }
            }
        };
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("Name,Count,Description\r\ninvalid,1,说明\r\n"));
        var importer = new CsvEntityImporter(namedValidationRules: new INamedExcelValidationRule[]
        {
            new CsvStartsWithOkValidationRule()
        });

        // Act
        var result = importer.Import<CsvRow>(source, new CsvImportOptions<CsvRow> { MappingConfiguration = mapping });

        // Assert
        Assert.Empty(result.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(nameof(CsvRow.Name), error.PropertyName);
    }

    /// <summary>
    /// 测试 - CSV 命名校验规则抛出异常时，应返回包含位置和属性的结构化错误。
    /// </summary>
    [Fact]
    public void EntityPipeline_ThrowingNamedValidationRule_ShouldReturnStructuredError()
    {
        // Arrange
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(CsvRow.Name),
                    ValidationRuleNames = new List<string> { "throwing" }
                }
            }
        };
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("Name,Count,Description\r\nvalue,1,说明\r\n"));
        var importer = new CsvEntityImporter(namedValidationRules: new INamedExcelValidationRule[]
        {
            new ThrowingCsvValidationRule()
        });

        // Act
        var result = importer.Import<CsvRow>(source, new CsvImportOptions<CsvRow> { MappingConfiguration = mapping });

        // Assert
        Assert.Empty(result.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
        Assert.Equal(nameof(CsvRow.Name), error.PropertyName);
        Assert.Equal("校验器异常", error.Message);
    }

    /// <summary>
    /// 测试 - CSV 导出默认应防止标题和数据字段被表格软件解释为公式。
    /// </summary>
    [Fact]
    public void EntityPipeline_FormulaLikeValues_ShouldEscapeByDefault()
    {
        // Arrange
        using var destination = new MemoryStream();

        // Act
        new CsvEntityExporter().Export(new[]
        {
            new CsvRow { Name = "=SUM(A1:A2)", Count = 1, Description = "+formula" }
        }, destination);

        // Assert
        var content = Encoding.UTF8.GetString(destination.ToArray());
        Assert.Contains("'=SUM(A1:A2)", content);
        Assert.Contains("'+formula", content);
    }

    /// <summary>
    /// 测试 - 预取消的 CSV 导入导出应在读取或写入前取消。
    /// </summary>
    [Fact]
    public void EntityPipeline_PreCancelledOperation_ShouldThrow()
    {
        // Arrange
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("Name,Count,Description\r\n"));
        using var destination = new MemoryStream();
        using var cancellationTokenSource = new System.Threading.CancellationTokenSource();
        cancellationTokenSource.Cancel();

        // Act
        var import = () => new CsvEntityImporter().Import<CsvRow>(source, cancellationToken: cancellationTokenSource.Token);
        var export = () => new CsvEntityExporter().Export(Array.Empty<CsvRow>(), destination,
            cancellationToken: cancellationTokenSource.Token);

        // Assert
        Assert.Throws<OperationCanceledException>(import);
        Assert.Throws<OperationCanceledException>(export);
    }

    /// <summary>
    /// 测试 - CSV 导入应执行特性校验，且重复值仅在成功行后提交。
    /// </summary>
    [Fact]
    public void EntityPipeline_AttributeValidationAndDuplicates_ShouldReturnStructuredErrors()
    {
        // Arrange
        const string content = "Code,Count\r\nA,invalid\r\nA,1\r\nA,2\r\n,3\r\n";
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(content));

        // Act
        var result = new CsvEntityImporter().Import<CsvValidatedRow>(source);

        // Assert
        Assert.Single(result.Items);
        Assert.Equal("A", result.Items[0].Code);
        Assert.Equal(3, result.Errors.Count);
        Assert.Contains(result.Errors, error => error.RowIndex == 4 && error.PropertyName == nameof(CsvValidatedRow.Code));
        Assert.Contains(result.Errors, error => error.RowIndex == 5 && error.PropertyName == nameof(CsvValidatedRow.Code));
    }

    /// <summary>
    /// 测试 - CSV 导入应通过 BindFilterAttribute 解析自定义校验规则并返回结构化错误。
    /// </summary>
    [Fact]
    public void EntityPipeline_BoundCustomValidationRule_ShouldReturnStructuredError()
    {
        // Arrange
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("Code\r\ninvalid\r\n"));
        var importer = new CsvEntityImporter(validationRules: new IExcelValidationRule[]
        {
            new CsvStartsWithOkAttributeValidationRule()
        });

        // Act
        var result = importer.Import<CsvBoundValidationRow>(source);

        // Assert
        Assert.Empty(result.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal(2, error.RowIndex);
        Assert.Equal(nameof(CsvBoundValidationRow.Code), error.PropertyName);
        Assert.Equal("必须以 OK- 开头", error.Message);
    }

    /// <summary>
    /// 测试 - CSV 闭合引号后的普通字符必须被拒绝，不能静默改变字段内容。
    /// </summary>
    [Fact]
    public void EntityPipeline_InvalidCharacterAfterQuotedField_ShouldThrow()
    {
        // Arrange
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("Name,Count,Description\r\n\"value\"x,1,test\r\n"));

        // Act
        var action = () => new CsvEntityImporter().Import<CsvRow>(source);

        // Assert
        Assert.Throws<InvalidOperationException>(action);
    }

    /// <summary>
    /// 测试 - CSV 实体管线应在导出和导入时保留请求定义的动态列。
    /// </summary>
    [Fact]
    public void EntityPipeline_DynamicColumns_ShouldRoundTrip()
    {
        // Arrange
        using var destination = new MemoryStream();
        var source = new CsvDynamicRow
        {
            Name = "固定",
            Values = new Dictionary<string, object> { ["扩展一"] = "A", ["扩展二"] = 2 }
        };

        // Act
        new CsvEntityExporter().Export(new[] { source }, destination,
            new CsvExportOptions<CsvDynamicRow> { DynamicColumns = new[] { "扩展一", "扩展二" } });
        destination.Position = 0;
        var result = new CsvEntityImporter().Import<CsvDynamicRow>(destination);

        // Assert
        var item = Assert.Single(result.Items);
        Assert.Empty(result.Errors);
        Assert.Equal("固定", item.Name);
        Assert.Equal("A", item.Values["扩展一"]);
        Assert.Equal("2", item.Values["扩展二"]);
        Assert.True(destination.CanRead);
    }

    /// <summary>
    /// 测试 - CSV 应按请求指定的区域性格式化并转换数值。
    /// </summary>
    [Fact]
    public void EntityPipeline_Culture_ShouldRoundTripDecimalValue()
    {
        // Arrange
        var culture = CultureInfo.GetCultureInfo("de-DE");
        using var destination = new MemoryStream();

        // Act
        new CsvEntityExporter().Export(new[] { new CsvDecimalRow { Value = 1.5m } }, destination,
            new CsvExportOptions<CsvDecimalRow> { Delimiter = ';', Culture = culture });
        destination.Position = 0;
        var result = new CsvEntityImporter().Import<CsvDecimalRow>(destination,
            new CsvImportOptions<CsvDecimalRow> { Delimiter = ';', Culture = culture });

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal(1.5m, Assert.Single(result.Items).Value);
    }

    /// <summary>
    /// 测试 - CSV 表头应复用统一映射中的历史 Alias 绑定固定属性。
    /// </summary>
    [Fact]
    public void EntityPipeline_AliasHeader_ShouldBindConfiguredProperty()
    {
        // Arrange
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(CsvRow.Name),
                    Aliases = new List<string> { "旧名称" }
                }
            }
        };
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("旧名称,Count,Description\r\n兼容,1,说明\r\n"));

        // Act
        var result = new CsvEntityImporter().Import<CsvRow>(source,
            new CsvImportOptions<CsvRow> { MappingConfiguration = mapping });

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal("兼容", Assert.Single(result.Items).Name);
    }

    /// <summary>
    /// 测试 - CSV 应按统一映射的列级 ImportWhitespace 策略规范化输入文本。
    /// </summary>
    [Fact]
    public void EntityPipeline_ColumnWhitespace_ShouldTrimBeforeConversion()
    {
        // Arrange
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new ExcelColumnConfiguration
                {
                    PropertyName = nameof(CsvRow.Name),
                    ImportWhitespace = ExcelWhitespacePolicy.Trim
                }
            }
        };
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("Name,Count,Description\r\n  带空白  ,1,说明\r\n"));

        // Act
        var result = new CsvEntityImporter().Import<CsvRow>(source,
            new CsvImportOptions<CsvRow> { MappingConfiguration = mapping });

        // Assert
        Assert.Empty(result.Errors);
        Assert.Equal("带空白", Assert.Single(result.Items).Name);
    }

    /// <summary>
    /// 测试 - CSV 应使用 normalized JSON 文档的 Import/Export 独立方向配置。
    /// </summary>
    [Fact]
    public void EntityPipeline_NormalizedDocument_ShouldUseDirectionalMappings()
    {
        // Arrange
        var document = ExcelMappingConfigurationLoader.FromJsonDocument(
            "{\"version\":2,\"import\":{\"columns\":[{\"propertyName\":\"Name\",\"title\":\"输入名称\"}]},\"export\":{\"columns\":[{\"propertyName\":\"Name\",\"title\":\"输出名称\"}]}}");
        using var source = new MemoryStream(Encoding.UTF8.GetBytes("输入名称,Count,Description\r\nCSV,1,说明\r\n"));
        using var destination = new MemoryStream();

        // Act
        var imported = new CsvEntityImporter().Import<CsvRow>(source,
            new CsvImportOptions<CsvRow> { MappingDocument = document });
        new CsvEntityExporter().Export(new[] { new CsvRow { Name = "CSV", Count = 1, Description = "说明" } },
            destination, new CsvExportOptions<CsvRow> { MappingDocument = document });

        // Assert
        Assert.Empty(imported.Errors);
        Assert.Equal("CSV", Assert.Single(imported.Items).Name);
        Assert.StartsWith("输出名称,Count,Description", Encoding.UTF8.GetString(destination.ToArray()),
            StringComparison.Ordinal);
    }

    /// <summary>
    /// 测试 - Unique 跟踪达到上限后应拒绝新值，已提交的首行仍应保留。
    /// </summary>
    [Fact]
    public void EntityPipeline_UniqueLimit_ShouldKeepCommittedRowsAndRejectOverflow()
    {
        // Arrange
        const string content = "Code,Count\r\nA,1\r\nB,2\r\n";
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(content));

        // Act
        var result = new CsvEntityImporter().Import<CsvValidatedRow>(source,
            new CsvImportOptions<CsvValidatedRow> { MaxTrackedUniqueValues = 1 });

        // Assert
        Assert.Single(result.Items);
        Assert.Equal("A", result.Items[0].Code);
        Assert.Single(result.Errors);
        Assert.Contains("Unique", result.Errors[0].Message);
    }

    /// <summary>
    /// 测试 - CSV 重复值错误应携带首次成功行号，便于定位唯一约束冲突。
    /// </summary>
    [Fact]
    public void EntityPipeline_UniqueDuplicate_ShouldReportFirstRowNumber()
    {
        // Arrange
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(
            "Code,Date,Amount\r\nOK-1,2026-01-01,1\r\nOK-1,2026-01-02,2\r\n"));

        // Act
        var result = new CsvEntityImporter().Import<CsvV2ValidatedRow>(source);

        // Assert
        var error = Assert.Single(result.Errors);
        Assert.Equal(2, error.FirstRowNumber);
    }

    /// <summary>
    /// 测试 - v2 校验特性应在 CSV 主链按原始值、转换值和唯一值顺序执行。
    /// </summary>
    [Fact]
    public void EntityPipeline_V2ValidationAttributes_ShouldValidateAllRules()
    {
        // Arrange
        const string content = "Code,Date,Amount\r\nOK-1,2024-01-02,5\r\nBAD,2024-01-02,5\r\n"
            + "OK-123456789,2024-01-02,5\r\nOK-2,not-date,5\r\nOK-3,2024-01-02,11\r\n"
            + "OK-1,2024-01-02,5\r\n,2024-01-02,5\r\n";
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(content));

        // Act
        var result = new CsvEntityImporter().Import<CsvV2ValidatedRow>(source);

        // Assert
        Assert.Single(result.Items);
        Assert.Equal(6, result.Errors.Count);
        Assert.Contains(result.Errors, error => error.RowIndex == 3 && error.PropertyName == nameof(CsvV2ValidatedRow.Code));
        Assert.Contains(result.Errors, error => error.RowIndex == 4 && error.PropertyName == nameof(CsvV2ValidatedRow.Code));
        Assert.Contains(result.Errors, error => error.RowIndex == 5 && error.PropertyName == nameof(CsvV2ValidatedRow.Date));
        Assert.Contains(result.Errors, error => error.RowIndex == 6 && error.PropertyName == nameof(CsvV2ValidatedRow.Amount));
        Assert.Contains(result.Errors, error => error.RowIndex == 7 && error.PropertyName == nameof(CsvV2ValidatedRow.Code));
        Assert.Contains(result.Errors, error => error.RowIndex == 8 && error.PropertyName == nameof(CsvV2ValidatedRow.Code));
    }

    /// <summary>
    /// 测试 - v2 Range Attribute 应在构建时拒绝最小值大于最大值。
    /// </summary>
    [Fact]
    public void ExcelRangeAttribute_InvalidBounds_ShouldThrow()
    {
        // Act
        var action = () => new ExcelRangeAttribute(5, 1);

        // Assert
        Assert.Throws<ArgumentException>(action);
    }

    /// <summary>
    /// CSV 实体测试模型。
    /// </summary>
    private class CsvRow
    {
        /// <summary>
        /// 名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 数量。
        /// </summary>
        public int Count { get; set; }

        /// <summary>
        /// 描述。
        /// </summary>
        public string Description { get; set; }
    }

    /// <summary>
    /// CSV 领域值对象测试模型。
    /// </summary>
    private class CsvConvertedRow
    {
        /// <summary>
        /// 领域编码。
        /// </summary>
        public CsvCode Code { get; set; }
    }

    /// <summary>
    /// CSV 特性校验测试模型。
    /// </summary>
    private class CsvValidatedRow
    {
        [Required]
        [Duplication]
        public string Code { get; set; }

        [Range(1, 9)]
        public int Count { get; set; }
    }

    /// <summary>
    /// 绑定自定义校验规则的 CSV 测试模型。
    /// </summary>
    private class CsvBoundValidationRow
    {
        [CsvStartsWithOkAttribute]
        public string Code { get; set; }
    }

    /// <summary>
    /// CSV 动态列测试模型。
    /// </summary>
    private class CsvDynamicRow
    {
        public string Name { get; set; }

        [DynamicColumn]
        public IDictionary<string, object> Values { get; set; }
    }

    /// <summary>
    /// CSV 区域性测试模型。
    /// </summary>
    private class CsvDecimalRow
    {
        public decimal Value { get; set; }
    }

    /// <summary>
    /// CSV v2 校验特性测试模型。
    /// </summary>
    private class CsvV2ValidatedRow
    {
        [ExcelRequired]
        [ExcelRegex("^OK-")]
        [ExcelMaxLength(8)]
        [ExcelUnique]
        public string Code { get; set; }

        [ExcelDate(Format = "yyyy-MM-dd")]
        public DateTime Date { get; set; }

        [ExcelRange(1, 10)]
        [ExcelMaxValue(10)]
        public int Amount { get; set; }
    }

    /// <summary>
    /// CSV 领域编码。
    /// </summary>
    private sealed class CsvCode
    {
        /// <summary>
        /// 初始化一个<see cref="CsvCode"/>类型的实例。
        /// </summary>
        /// <param name="value">编码值。</param>
        public CsvCode(string value) => Value = value;

        /// <summary>
        /// 获取编码值。
        /// </summary>
        public string Value { get; }
    }

    /// <summary>
    /// 命名 CSV 领域编码转换器。
    /// </summary>
    private sealed class CsvCodeConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "csv-code";

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(CsvCode);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = new CsvCode(((string)context.Value).Substring(3));
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = $"CV-{((CsvCode)context.Value).Value}";
            return true;
        }
    }

    /// <summary>
    /// 命名 CSV 前缀校验规则。
    /// </summary>
    private sealed class CsvStartsWithOkValidationRule : INamedExcelValidationRule
    {
        /// <inheritdoc />
        public string Name => "starts-with-ok";

        /// <inheritdoc />
        public string ErrorMessage => "必须以 OK- 开头";

        /// <inheritdoc />
        public bool Validate(ExcelValidationContext context) => context.Value.StartsWith("OK-", StringComparison.Ordinal);
    }

    /// <summary>
    /// 要求文本使用指定前缀的 CSV 自定义校验特性。
    /// </summary>
    [BindFilter(typeof(CsvStartsWithOkAttributeValidationRule))]
    private sealed class CsvStartsWithOkAttribute : FilterAttributeBase
    {
        /// <inheritdoc />
        public override string ErrorMsg { get; set; } = "必须以 OK- 开头";
    }

    /// <summary>
    /// CSV 自定义前缀校验规则。
    /// </summary>
    private sealed class CsvStartsWithOkAttributeValidationRule : IExcelValidationRule
    {
        /// <inheritdoc />
        public bool CanValidate(FilterAttributeBase attribute) => attribute is CsvStartsWithOkAttribute;

        /// <inheritdoc />
        public bool Validate(FilterAttributeBase attribute, ExcelValidationContext context) =>
            context.Value.StartsWith("OK-", StringComparison.Ordinal);
    }

    /// <summary>
    /// 会抛出异常的命名 CSV 校验规则。
    /// </summary>
    private sealed class ThrowingCsvValidationRule : INamedExcelValidationRule
    {
        /// <inheritdoc />
        public string Name => "throwing";

        /// <inheritdoc />
        public string ErrorMessage => "不应使用默认消息";

        /// <inheritdoc />
        public bool Validate(ExcelValidationContext context) => throw new InvalidOperationException("校验器异常");
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
