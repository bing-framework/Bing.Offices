using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Csv;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>CSV 表头歧义、无表头动态列及解析资源边界的回归测试。</summary>
public sealed class CsvImportBindingRegressionTest
{
    /// <summary>不同标题或别名指向同一固定字段时，必须在写入实体前拒绝。</summary>
    [Theory]
    [InlineData(false, "Display,Legacy")]
    [InlineData(true, "Display,Legacy")]
    [InlineData(false, "Display,Name")]
    [InlineData(true, "Legacy,Name")]
    public async Task Import_DuplicateFixedTarget_ShouldReturnInvalidHeader(bool asynchronous, string header)
    {
        var options = new CsvImportOptions<Row>
        {
            MappingConfiguration = new ExcelMappingConfiguration
            {
                Columns = new List<ExcelColumnConfiguration>
                {
                    new() { PropertyName = "Name", Title = "Display", Aliases = new List<string> { "Legacy" } }
                }
            }
        };
        var result = await Import(header + "\r\nfirst,second\r\n", options, asynchronous);
        Assert.Empty(result.Items);
        AssertError(result, CsvImportErrorCode.InvalidHeader, 1, 0, null,
            "CSV 多个表头映射到同一属性: Name", false);
    }

    /// <summary>同一动态键的标题和别名不能同时绑定。</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Import_DuplicateDynamicTarget_ShouldReturnInvalidHeader(bool asynchronous)
    {
        var options = new CsvImportOptions<DynamicRow> { MappingConfiguration = DynamicMapping() };
        var result = await Import("Name,Amount,OldAmount,Tail\r\nfirst,1,2,last\r\n", options, asynchronous);
        Assert.Empty(result.Items);
        AssertError(result, CsvImportErrorCode.InvalidHeader, 1, 0, null,
            "CSV 多个表头映射到同一动态列: amount", false);
    }

    /// <summary>合法的动态别名和多个不同键应完整导入。</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Import_DistinctDynamicAliases_ShouldPreserveAllValues(bool asynchronous)
    {
        var result = await Import("Name,OldAmount,A,Z,Tail\r\nfirst,7,east,west,last\r\n",
            new CsvImportOptions<DynamicRow> { MappingConfiguration = DynamicMapping() }, asynchronous);
        AssertCompleteRow(result);
    }

    /// <summary>无表头的同步异步往返必须遵循固定列位置及动态 Order/Key 顺序。</summary>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task Import_NoHeaderDynamicColumns_ShouldRoundTripInExportOrder(bool asyncExport, bool asyncImport)
    {
        var data = new[]
        {
            new DynamicRow
            {
                Name = "first", Tail = "last",
                Values = new Dictionary<string, object> { ["z"] = "west", ["amount"] = 7, ["a"] = "east" }
            }
        };
        using var output = new MemoryStream();
        var options = new CsvExportOptions<DynamicRow> { IncludeHeader = false, MappingConfiguration = DynamicMapping() };
        var exporter = new CsvEntityExporter();
        if (asyncExport)
            await exporter.ExportAsync(data, output, options);
        else
            exporter.Export(data, output, options);
        Assert.Equal("first,7,east,west,last\r\n", Encoding.UTF8.GetString(output.ToArray()));
        var result = await Import(Encoding.UTF8.GetString(output.ToArray()),
            new CsvImportOptions<DynamicRow> { HasHeader = false, MappingConfiguration = DynamicMapping() }, asyncImport);
        AssertCompleteRow(result);
    }

    /// <summary>无表头模式无法推断未声明的动态列，即使输入为空也应明确拒绝。</summary>
    [Theory]
    [InlineData(false, "first,extra,last\r\n")]
    [InlineData(true, "first,extra,last\r\n")]
    [InlineData(false, "")]
    [InlineData(true, "")]
    public async Task Import_NoHeaderWithoutDynamicDefinitions_ShouldReject(bool asynchronous, string content)
    {
        var result = await Import(content, new CsvImportOptions<DynamicRow> { HasHeader = false }, asynchronous);
        Assert.Empty(result.Items);
        AssertError(result, CsvImportErrorCode.InvalidHeader, 1, 0, null,
            "无表头 CSV 导入必须显式配置动态列。", false);
    }

    /// <summary>展开后的动态列也必须计入无表头映射列数限制。</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Import_NoHeaderDynamicColumns_ShouldEnforceColumnLimit(bool asynchronous)
    {
        var result = await Import("first,7,east,west,last\r\n", new CsvImportOptions<DynamicRow>
        {
            HasHeader = false, MappingConfiguration = DynamicMapping(), MaxColumns = 4
        }, asynchronous);
        Assert.Empty(result.Items);
        AssertError(result, CsvImportErrorCode.ResourceLimit, 0, 0, null, "CSV 映射列数超过最大列数: 4", true);
    }

    /// <summary>无表头动态字段仍执行类型转换并返回原始列坐标。</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Import_NoHeaderDynamicColumns_ShouldReportConversionError(bool asynchronous)
    {
        var result = await Import("first,invalid,east,west,last\r\n", new CsvImportOptions<DynamicRow>
        {
            HasHeader = false, MappingConfiguration = DynamicMapping()
        }, asynchronous);
        Assert.Empty(result.Items);
        var error = Assert.Single(result.Errors);
        Assert.Equal(CsvImportErrorCode.ValueConversion, error.Code);
        Assert.Equal(1, error.RowIndex);
        Assert.Equal(2, error.ColumnIndex);
        Assert.Equal("amount", error.PropertyName);
        Assert.Null(error.FirstRowNumber);
        Assert.False(string.IsNullOrWhiteSpace(error.Message));
        Assert.False(result.IsTruncated);
    }

    /// <summary>字段限制必须覆盖表头、映射字段和未映射字段。</summary>
    [Theory]
    [InlineData(false, "Name,Extra\r\nok,123456789\r\n", 2, 2, null)]
    [InlineData(true, "Name,Extra\r\nok,123456789\r\n", 2, 2, null)]
    [InlineData(false, "Name,ExtraLong\r\nok,ok\r\n", 1, 2, null)]
    [InlineData(true, "Name,ExtraLong\r\nok,ok\r\n", 1, 2, null)]
    [InlineData(false, "Name\r\n123456789\r\n", 2, 1, "Name")]
    [InlineData(true, "Name\r\n123456789\r\n", 2, 1, "Name")]
    public async Task Import_FieldLimit_ShouldCoverEveryField(bool asynchronous, string content,
        int row, int column, string property)
    {
        var result = await Import(content, new CsvImportOptions<Row> { MaxFieldLength = 8 }, asynchronous);
        Assert.Empty(result.Items);
        AssertError(result, CsvImportErrorCode.ResourceLimit, row, column, property,
            "CSV 字段超过最大字符数: 8", true);
    }

    /// <summary>长度按解码字符计算，允许边界处的双引号转义、换行和中文。</summary>
    [Theory]
    [InlineData(false, "\"\"\"\"")]
    [InlineData(true, "\"\"\"\"")]
    [InlineData(false, "a\r\nb")]
    [InlineData(true, "a\r\nb")]
    [InlineData(false, "中文测试")]
    [InlineData(true, "中文测试")]
    public async Task Import_FieldLimit_ShouldMeasureDecodedCharacters(bool asynchronous, string value)
    {
        using var output = new MemoryStream();
        new CsvEntityExporter().Export(new[] { new Row { Name = value } }, output,
            new CsvExportOptions<Row> { IncludeHeader = false });
        var result = await Import(Encoding.UTF8.GetString(output.ToArray()),
            new CsvImportOptions<Row> { HasHeader = false, MaxFieldLength = 4 }, asynchronous);
        Assert.Empty(result.Errors);
        Assert.Equal(value, Assert.Single(result.Items).Name);
        Assert.False(result.IsTruncated);
    }

    /// <summary>解析器应在巨型未映射字段读取完毕前停止，异步入口不得调用同步读取。</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Import_FieldLimit_ShouldStopReadingOversizedUnmappedField(bool asynchronous)
    {
        var bytes = Encoding.UTF8.GetBytes("Name,Extra\r\nok," + new string('x', 32768) + "\r\n");
        using var input = new ReadTrackingStream(bytes);
        var importer = new CsvEntityImporter();
        var options = new CsvImportOptions<Row> { MaxFieldLength = 8 };
        var result = asynchronous ? await importer.ImportAsync(input, options) : importer.Import(input, options);
        Assert.Empty(result.Items);
        AssertError(result, CsvImportErrorCode.ResourceLimit, 2, 0, null,
            "CSV 字段超过最大字符数: 8", true);
        Assert.InRange(input.BytesRead, 1, bytes.Length - 1);
        Assert.Equal(asynchronous ? 0 : input.ReadCalls, input.SyncReadCalls);
        Assert.True(input.CanRead);
    }

    private static async Task<CsvImportResult<T>> Import<T>(string text, CsvImportOptions<T> options, bool asynchronous)
        where T : class, new()
    {
        using var source = new MemoryStream(Encoding.UTF8.GetBytes(text));
        var importer = new CsvEntityImporter();
        var result = asynchronous ? await importer.ImportAsync(source, options) : importer.Import(source, options);
        Assert.True(source.CanRead);
        return result;
    }

    private static void AssertError<T>(CsvImportResult<T> result, CsvImportErrorCode code, int row, int column,
        string property, string message, bool truncated)
    {
        var error = Assert.Single(result.Errors);
        Assert.Equal(code, error.Code);
        Assert.Equal(row, error.RowIndex);
        Assert.Equal(column, error.ColumnIndex);
        Assert.Equal(property, error.PropertyName);
        Assert.Equal(message, error.Message);
        Assert.Null(error.FirstRowNumber);
        Assert.Equal(truncated, result.IsTruncated);
        Assert.Null(result.MaxErrors);
    }

    private static void AssertCompleteRow(CsvImportResult<DynamicRow> result)
    {
        Assert.Empty(result.Errors);
        Assert.False(result.IsTruncated);
        Assert.Null(result.MaxErrors);
        var item = Assert.Single(result.Items);
        Assert.Equal("first", item.Name);
        Assert.Equal("last", item.Tail);
        Assert.Equal(3, item.Values.Count);
        Assert.Equal(7, Assert.IsType<int>(item.Values["amount"]));
        Assert.Equal("east", item.Values["a"]);
        Assert.Equal("west", item.Values["z"]);
    }

    private static ExcelMappingConfiguration DynamicMapping() => new()
    {
        DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
        {
            new() { Key = "z", Title = "Z", DataTypeName = "string", Order = 1 },
            new() { Key = "a", Title = "A", DataTypeName = "string", Order = 1 },
            new() { Key = "amount", Title = "Amount", DataTypeName = "int32", Order = 0,
                Aliases = new List<string> { "OldAmount" } }
        }
    };

    public sealed class Row
    {
        public string Name { get; set; }
    }

    public sealed class DynamicRow
    {
        public string Name { get; set; }
        [DynamicColumn]
        public IDictionary<string, object> Values { get; set; } = new Dictionary<string, object>();
        public string Tail { get; set; }
    }

    private sealed class ReadTrackingStream : MemoryStream
    {
        public ReadTrackingStream(byte[] content) : base(content) { }
        public int BytesRead { get; private set; }
        public int ReadCalls { get; private set; }
        public int SyncReadCalls { get; private set; }
        public override int Read(byte[] buffer, int offset, int count)
        {
            SyncReadCalls++;
            ReadCalls++;
            var read = base.Read(buffer, offset, count);
            BytesRead += read;
            return read;
        }
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ReadCalls++;
            var read = base.Read(buffer, offset, count);
            BytesRead += read;
            return Task.FromResult(read);
        }
    }
}
