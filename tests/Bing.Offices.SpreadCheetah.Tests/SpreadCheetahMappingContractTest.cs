using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.SpreadCheetah.Tests;

/// <summary>
/// 验证 SpreadCheetah 流式导出对 Core 映射计划的实际消费行为。
/// </summary>
public sealed class SpreadCheetahMappingContractTest
{
    /// <summary>
    /// 验证导出应用值映射、命名转换器和动态列物理位置。
    /// </summary>
    [Fact]
    public async Task ExportBatchesAsync_ShouldApplyValueMapNamedConverterAndDynamicPhysicalLayout()
    {
        var converter = new UpperCaseConverter();
        var factory = ExcelMappingPlanFactoryProvider.CreateDefault(new[] { converter });
        var plan = factory.Create<MappingRow>(new ExcelMappingDocument { UseConventionFallback = true },
            new ExcelMappingConfiguration
            {
                Columns = new List<ExcelColumnConfiguration>
                {
                    new()
                    {
                        PropertyName = nameof(MappingRow.Code),
                        Title = "代码",
                        ValueMappings = new List<ExcelValueMappingConfiguration>
                        {
                            new() { Text = "启用", Value = "1" }
                        }
                    }
                }
            }, MappingDirection.Export);
        Assert.Equal("1", plan.Columns.Single(column => column.Name == nameof(MappingRow.Code))
            .ValueMap["启用"]);
        var exporter = new SpreadCheetahStreamingExcelExporter(factory);
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
            new[]
            {
                new MappingRow
                {
                    Code = 1,
                    Name = "plain",
                    Extras = new Dictionary<string, object> { ["Alias"] = "mixed" }
                }
            }, sheet => sheet
                .Mapping(new ExcelMappingConfiguration
                {
                    Columns = new List<ExcelColumnConfiguration>
                    {
                        new()
                        {
                            PropertyName = nameof(MappingRow.Code),
                            Title = "代码",
                            ValueMappings = new List<ExcelValueMappingConfiguration>
                            {
                                new() { Text = "启用", Value = "1" }
                            }
                        }
                    }
                })
                .DynamicColumns(row => row.Extras, new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "Alias",
                        Title = "别名",
                        ConverterName = converter.Name,
                        PhysicalColumnIndex = 1
                    }
                })));
        Assert.NotNull(request.Sheets[0].MappingConfiguration);
        Assert.Equal("1", request.Sheets[0].MappingConfiguration.Columns.Single().ValueMappings.Single().Value);
        using var destination = new MemoryStream();

        await exporter.ExportBatchesAsync(request, destination);

        using var workbook = Open(destination);
        var sheet = workbook.GetSheet("Rows");
        Assert.Equal("代码", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("别名", sheet.GetRow(0).GetCell(1).StringCellValue);
        Assert.Equal("Name", sheet.GetRow(0).GetCell(2).StringCellValue);
        Assert.Equal("启用", sheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("MIXED", sheet.GetRow(1).GetCell(1).StringCellValue);
        Assert.Equal("PLAIN", sheet.GetRow(1).GetCell(2).StringCellValue);
    }

    /// <summary>
    /// 验证无效映射在枚举数据和写入输出前被拒绝。
    /// </summary>
    [Fact]
    public async Task ExportBatchesAsync_ShouldRejectInvalidMappingBeforeEnumeratingOrWriting()
    {
        var source = new CountingEnumerable<MappingRow>(new[]
        {
            new MappingRow { Code = 1, Name = "should-not-be-read" }
        });
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows", source, sheet =>
            sheet.Mapping(new ExcelMappingConfiguration
            {
                Columns = new List<ExcelColumnConfiguration>
                {
                    new() { PropertyName = "MissingProperty", Title = "无效" }
                }
            })));
        using var destination = new MemoryStream();

        await Assert.ThrowsAnyAsync<Exception>(() =>
            new SpreadCheetahStreamingExcelExporter().ExportBatchesAsync(request, destination));

        Assert.Equal(0, source.EnumeratorCount);
        Assert.Equal(0, source.MoveNextCount);
        Assert.Equal(0, destination.Length);
    }

    /// <summary>
    /// 验证缺失命名转换器在枚举前产生结构化配置错误。
    /// </summary>
    [Fact]
    public async Task ExportBatchesAsync_WhenNamedConverterIsMissing_ShouldReturnStructuredConfigurationErrorBeforeEnumerating()
    {
        var source = new CountingEnumerable<MappingRow>(new[]
        {
            new MappingRow { Code = 1, Name = "should-not-be-read" }
        });
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows", source, sheet =>
            sheet.Mapping(new ExcelMappingConfiguration
            {
                Columns = new List<ExcelColumnConfiguration>
                {
                    new()
                    {
                        PropertyName = nameof(MappingRow.Name),
                        ConverterName = "missing-converter"
                    }
                }
            })));
        using var destination = new MemoryStream();

        var exception = await Assert.ThrowsAsync<BingOfficesConfigurationException>(() =>
            new SpreadCheetahStreamingExcelExporter().ExportBatchesAsync(request, destination));

        Assert.Equal(BingOfficesErrorCode.ConfigurationInvalid, exception.Code);
        Assert.Equal(BingOfficesOperation.Configuration, exception.Operation);
        Assert.Equal("Core", exception.Provider);
        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Contains("missing-converter", exception.Message, StringComparison.Ordinal);
        Assert.IsType<InvalidOperationException>(exception.InnerException);
        Assert.Equal(0, source.EnumeratorCount);
        Assert.Equal(0, source.MoveNextCount);
        Assert.Equal(0, destination.Length);
    }

    /// <summary>
    /// 验证值转换器失败产生结构化导出错误。
    /// </summary>
    /// <param name="asynchronous">是否调用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExportBatches_WhenValueConverterFails_ShouldReturnStructuredExportError(bool asynchronous)
    {
        var converter = new ThrowingExportConverter();
        var factory = ExcelMappingPlanFactoryProvider.CreateDefault(new[] { converter });
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
            new[] { new SingleValueRow { Value = "value" } }, sheet => sheet.Mapping(
                new ExcelMappingConfiguration
                {
                    Columns = new List<ExcelColumnConfiguration>
                    {
                        new() { PropertyName = nameof(SingleValueRow.Value), ConverterName = converter.Name }
                    }
                })));
        using var destination = new MemoryStream();

        var exception = await AssertExportExceptionAsync(new SpreadCheetahStreamingExcelExporter(factory),
            request, destination, asynchronous);

        Assert.Equal(BingOfficesErrorCode.UserExtensionFailed, exception.Code);
        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal("SpreadCheetah", exception.Provider);
        Assert.Equal(BingOfficesStage.Validate, exception.Stage);
        Assert.Equal("Rows", exception.SheetName);
        Assert.Equal(2, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
        Assert.Equal(nameof(SingleValueRow.Value), exception.PropertyName);
        Assert.IsType<InvalidOperationException>(exception.InnerException);
        Assert.Equal("导出转换器异常", exception.InnerException.Message);
    }

    /// <summary>
    /// 验证动态值读取失败产生结构化导出错误。
    /// </summary>
    /// <param name="asynchronous">是否调用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExportBatches_WhenDynamicGetterFails_ShouldReturnStructuredExportError(bool asynchronous)
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
            new[] { new ThrowingDynamicRow() }, sheet => sheet.DynamicColumns(row => row.Values,
                new[] { new ExcelDynamicColumnDefinition { Key = "amount", Title = "金额" } })));
        using var destination = new MemoryStream();

        var exception = await AssertExportExceptionAsync(new SpreadCheetahStreamingExcelExporter(),
            request, destination, asynchronous);

        Assert.Equal(BingOfficesErrorCode.UserExtensionFailed, exception.Code);
        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal("SpreadCheetah", exception.Provider);
        Assert.Equal(BingOfficesStage.Validate, exception.Stage);
        Assert.Equal("Rows", exception.SheetName);
        Assert.Equal(2, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
        Assert.Equal("amount", exception.PropertyName);
        Assert.IsType<InvalidOperationException>(exception.InnerException);
        Assert.Equal("动态读取器异常", exception.InnerException.Message);
    }

    /// <summary>
    /// 验证动态值转换失败产生结构化导出错误。
    /// </summary>
    /// <param name="asynchronous">是否调用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExportBatches_WhenDynamicValueCannotConvert_ShouldReturnStructuredExportError(
        bool asynchronous)
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
            new[]
            {
                new DynamicValueRow
                {
                    Values = new Dictionary<string, object> { ["amount"] = "not-a-number" }
                }
            }, sheet => sheet.DynamicColumns(row => row.Values,
                new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "amount",
                        Title = "金额",
                        DataType = typeof(int)
                    }
                })));
        using var destination = new MemoryStream();

        var exception = await AssertExportExceptionAsync(new SpreadCheetahStreamingExcelExporter(),
            request, destination, asynchronous);

        Assert.Equal(BingOfficesErrorCode.ExportFailed, exception.Code);
        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal("SpreadCheetah", exception.Provider);
        Assert.Equal(BingOfficesStage.Validate, exception.Stage);
        Assert.Equal("Rows", exception.SheetName);
        Assert.Equal(2, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
        Assert.Equal("amount", exception.PropertyName);
        Assert.NotNull(exception.InnerException);
    }

    /// <summary>
    /// 验证属性读取失败产生结构化导出错误。
    /// </summary>
    /// <param name="asynchronous">是否调用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExportBatches_WhenPropertyGetterFails_ShouldReturnStructuredExportError(bool asynchronous)
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
            new[] { new ThrowingPropertyRow() }, sheet => sheet.Mapping(
                new ExcelMappingConfiguration
                {
                    Columns = new List<ExcelColumnConfiguration>
                    {
                        new() { PropertyName = nameof(ThrowingPropertyRow.Value) }
                    }
                })));
        using var destination = new MemoryStream();

        var exception = await AssertExportExceptionAsync(new SpreadCheetahStreamingExcelExporter(),
            request, destination, asynchronous);

        Assert.Equal(BingOfficesErrorCode.UserExtensionFailed, exception.Code);
        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal("SpreadCheetah", exception.Provider);
        Assert.Equal(BingOfficesStage.Validate, exception.Stage);
        Assert.Equal("Rows", exception.SheetName);
        Assert.Equal(2, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
        Assert.Equal(nameof(ThrowingPropertyRow.Value), exception.PropertyName);
        Assert.IsType<InvalidOperationException>(exception.InnerException);
        Assert.Equal("属性读取器异常", exception.InnerException.Message);
    }

    /// <summary>
    /// 验证动态转换器失败产生结构化导出错误。
    /// </summary>
    /// <param name="asynchronous">是否调用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExportBatches_WhenDynamicConverterFails_ShouldReturnStructuredExportError(
        bool asynchronous)
    {
        var converter = new ThrowingDynamicConverter();
        var factory = ExcelMappingPlanFactoryProvider.CreateDefault(new[] { converter });
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
            new[]
            {
                new DynamicValueRow
                {
                    Values = new Dictionary<string, object> { ["amount"] = "value" }
                }
            }, sheet => sheet.DynamicColumns(row => row.Values,
                new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "amount",
                        Title = "金额",
                        ConverterName = converter.Name
                    }
                })));
        using var destination = new MemoryStream();

        var exception = await AssertExportExceptionAsync(new SpreadCheetahStreamingExcelExporter(factory),
            request, destination, asynchronous);

        Assert.Equal(BingOfficesErrorCode.UserExtensionFailed, exception.Code);
        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal("SpreadCheetah", exception.Provider);
        Assert.Equal(BingOfficesStage.Validate, exception.Stage);
        Assert.Equal("Rows", exception.SheetName);
        Assert.Equal(2, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
        Assert.Equal("amount", exception.PropertyName);
        Assert.IsType<InvalidOperationException>(exception.InnerException);
        Assert.Equal("动态转换器异常", exception.InnerException.Message);
    }

    /// <summary>
    /// 验证格式化失败产生结构化导出错误。
    /// </summary>
    /// <param name="asynchronous">是否调用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExportBatches_WhenFormatterFails_ShouldReturnStructuredExportError(bool asynchronous)
    {
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
            new[] { new FormatterRow { Value = 1.2m } }, sheet => sheet.Mapping(
                new ExcelMappingConfiguration
                {
                    Columns = new List<ExcelColumnConfiguration>
                    {
                        new() { PropertyName = nameof(FormatterRow.Value), Formatter = "Q" }
                    }
                })));
        using var destination = new MemoryStream();

        var exception = await AssertExportExceptionAsync(new SpreadCheetahStreamingExcelExporter(),
            request, destination, asynchronous);

        Assert.Equal(BingOfficesErrorCode.ExportFailed, exception.Code);
        Assert.Equal(BingOfficesOperation.Export, exception.Operation);
        Assert.Equal("SpreadCheetah", exception.Provider);
        Assert.Equal(BingOfficesStage.Validate, exception.Stage);
        Assert.Equal("Rows", exception.SheetName);
        Assert.Equal(2, exception.RowIndex);
        Assert.Equal(1, exception.ColumnIndex);
        Assert.Equal(nameof(FormatterRow.Value), exception.PropertyName);
        Assert.IsType<FormatException>(exception.InnerException);
    }

    /// <summary>
    /// 验证属性读取取消保留原始异常对象。
    /// </summary>
    /// <param name="asynchronous">是否调用异步导出入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ExportBatches_WhenPropertyGetterCancels_ShouldPreserveOriginalException(bool asynchronous)
    {
        var expected = new OperationCanceledException("属性读取取消");
        var row = new CancelingPropertyRow();
        row.SetException(expected);
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
            new[] { row }, sheet => sheet.Mapping(
                new ExcelMappingConfiguration
                {
                    Columns = new List<ExcelColumnConfiguration>
                    {
                        new() { PropertyName = nameof(CancelingPropertyRow.Value) }
                    }
                })));
        using var destination = new MemoryStream();

        var exception = await AssertCancellationAsync(new SpreadCheetahStreamingExcelExporter(),
            request, destination, asynchronous);

        Assert.Same(expected, exception);
    }

    /// <summary>
    /// 执行导出并断言结构化导出异常。
    /// </summary>
    /// <param name="exporter">待验证的流式导出器。</param>
    /// <param name="request">触发目标行为的导出请求。</param>
    /// <param name="destination">接收测试导出结果的内存流。</param>
    /// <param name="asynchronous">是否调用异步导出入口。</param>
    /// <returns>断言捕获的结构化导出异常。</returns>
    private static async Task<BingOfficesExportException> AssertExportExceptionAsync(
        SpreadCheetahStreamingExcelExporter exporter, ExcelWorkbookExportRequest request,
        MemoryStream destination, bool asynchronous)
    {
        if (asynchronous)
            return await Assert.ThrowsAsync<BingOfficesExportException>(() =>
                exporter.ExportBatchesAsync(request, destination));
        return Assert.Throws<BingOfficesExportException>(() => exporter.ExportBatches(request, destination));
    }

    /// <summary>
    /// 执行导出并断言取消异常。
    /// </summary>
    /// <param name="exporter">待验证的流式导出器。</param>
    /// <param name="request">触发目标行为的导出请求。</param>
    /// <param name="destination">接收测试导出结果的内存流。</param>
    /// <param name="asynchronous">是否调用异步导出入口。</param>
    /// <returns>断言捕获的取消异常。</returns>
    private static async Task<OperationCanceledException> AssertCancellationAsync(
        SpreadCheetahStreamingExcelExporter exporter, ExcelWorkbookExportRequest request,
        MemoryStream destination, bool asynchronous)
    {
        if (asynchronous)
            return await Assert.ThrowsAsync<OperationCanceledException>(() =>
                exporter.ExportBatchesAsync(request, destination));
        return Assert.Throws<OperationCanceledException>(() => exporter.ExportBatches(request, destination));
    }

    /// <summary>
    /// 从导出流的副本打开工作簿。
    /// </summary>
    /// <param name="stream">包含导出工作簿的内存流。</param>
    /// <returns>从独立流副本加载的工作簿。</returns>
    private static XSSFWorkbook Open(MemoryStream stream) =>
        new(new MemoryStream(stream.ToArray()));

    /// <summary>
    /// 用于固定列映射和动态列布局测试的数据行。
    /// </summary>
    private sealed class MappingRow
    {
        /// <summary>
        /// 获取或设置用于值映射验证的代码。
        /// </summary>
        public int Code { get; set; }
        /// <summary>
        /// 获取或设置测试数据名称。
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置测试动态列值。
        /// </summary>
        public IDictionary<string, object> Extras { get; set; } = new Dictionary<string, object>();
    }

    /// <summary>
    /// 将导出文本转换为大写的测试转换器。
    /// </summary>
    private sealed class UpperCaseConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "upper-case";

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = context.Value;
            return false;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = context.Value == null
                ? null
                : Convert.ToString(context.Value, context.Culture)?.ToUpperInvariant();
            return true;
        }
    }

    /// <summary>
    /// 导出时抛出异常的测试转换器。
    /// </summary>
    private sealed class ThrowingExportConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "throwing-export";

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = null;
            return false;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value) =>
            throw new InvalidOperationException("导出转换器异常");
    }

    /// <summary>
    /// 动态列导出时抛出异常的测试转换器。
    /// </summary>
    private sealed class ThrowingDynamicConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "throwing-dynamic";

        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = null;
            return false;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value) =>
            throw new InvalidOperationException("动态转换器异常");
    }

    /// <summary>
    /// 用于单列转换器测试的数据行。
    /// </summary>
    private sealed class SingleValueRow
    {
        /// <summary>
        /// 获取或设置待导出的测试值。
        /// </summary>
        public string Value { get; set; }
    }

    /// <summary>
    /// 读取动态值时抛出异常的测试数据行。
    /// </summary>
    private sealed class ThrowingDynamicRow
    {
        /// <summary>
        /// 获取动态值字典并触发预设读取异常。
        /// </summary>
        public IDictionary<string, object> Values =>
            throw new InvalidOperationException("动态读取器异常");
    }

    /// <summary>
    /// 用于动态值转换测试的数据行。
    /// </summary>
    private sealed class DynamicValueRow
    {
        /// <summary>
        /// 获取或设置待导出的动态值字典。
        /// </summary>
        public IDictionary<string, object> Values { get; set; }
    }

    /// <summary>
    /// 读取属性时抛出异常的测试数据行。
    /// </summary>
    private sealed class ThrowingPropertyRow
    {
        /// <summary>
        /// 获取测试值并触发预设读取异常。
        /// </summary>
        public string Value => throw new InvalidOperationException("属性读取器异常");
    }

    /// <summary>
    /// 用于验证数值格式化失败的数据行。
    /// </summary>
    private sealed class FormatterRow
    {
        /// <summary>
        /// 获取或设置待导出的测试值。
        /// </summary>
        public decimal Value { get; set; }
    }

    /// <summary>
    /// 读取属性时抛出指定取消异常的数据行。
    /// </summary>
    private sealed class CancelingPropertyRow
    {
        /// <summary>
        /// 属性读取时抛出的原始取消异常。
        /// </summary>
        private OperationCanceledException _exception;

        /// <summary>
        /// 初始化一个 <see cref="CancelingPropertyRow"/> 类型的实例。
        /// </summary>
        public CancelingPropertyRow()
        {
        }

        /// <summary>
        /// 设置属性读取时抛出的取消异常。
        /// </summary>
        /// <param name="exception">属性读取时应抛出的原始取消异常。</param>
        public void SetException(OperationCanceledException exception) => _exception = exception;

        /// <summary>
        /// 获取测试值并触发预设读取异常。
        /// </summary>
        public string Value => throw _exception;
    }

    /// <summary>
    /// 记录枚举器创建和推进次数的测试序列。
    /// </summary>
    /// <typeparam name="T">测试序列的元素类型。</typeparam>
    private sealed class CountingEnumerable<T> : IEnumerable<T>
    {
        /// <summary>
        /// 固定保存供测试枚举的数据快照。
        /// </summary>
        private readonly IReadOnlyList<T> _items;

        /// <summary>
        /// 初始化一个 <see cref="CountingEnumerable{T}"/> 类型的实例。
        /// </summary>
        /// <param name="items">待枚举的测试数据。</param>
        public CountingEnumerable(IEnumerable<T> items) => _items = items.ToArray();

        /// <summary>
        /// 获取或设置枚举器创建次数。
        /// </summary>
        public int EnumeratorCount { get; private set; }
        /// <summary>
        /// 获取或设置枚举推进次数。
        /// </summary>
        public int MoveNextCount { get; private set; }

        /// <inheritdoc />
        public IEnumerator<T> GetEnumerator()
        {
            EnumeratorCount++;
            for (var index = 0; index < _items.Count; index++)
            {
                MoveNextCount++;
                yield return _items[index];
            }
            MoveNextCount++;
        }

        /// <inheritdoc />
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
