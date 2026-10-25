using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Configurations;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>单表六入口的真实格式、结构错误、配置、所有权与异常传播测试。</summary>
public sealed class ExcelSingleSheetImportTest
{
    public static IEnumerable<object[]> Entrypoints()
    {
        foreach (var format in new[] { "xls", "xlsx" })
        foreach (var entry in new[] { "stream", "stream-async", "bytes", "bytes-async", "file", "file-async" })
            yield return new object[] { format, entry };
    }

    /// <summary>每个入口都应读取默认首表，完整保留成功实体和零基源行。</summary>
    [Theory]
    [MemberData(nameof(Entrypoints))]
    public async Task ImportSheet_AllEntrypoints_ShouldReturnCompleteFirstSheet(string format, string entry)
    {
        IExcelImporter importer = new NpoiExcelImporter();
        var bytes = Workbook(format, "Count", "7", "9");
        using var stream = new MemoryStream(bytes, writable: false);
        var path = Path.Combine(Path.GetTempPath(), $"bing-single-sheet-direct-{Guid.NewGuid():N}.bin");
        ExcelSingleSheetImportResult<Row> result;
        try
        {
            if (entry.StartsWith("file", StringComparison.Ordinal))
                File.WriteAllBytes(path, bytes);
            // 直接调用六个公开签名，供 IL 职责覆盖门禁追溯到本测试。
            result = entry switch
            {
                "stream" => importer.ImportSheet<Row>(stream),
                "stream-async" => await importer.ImportSheetAsync<Row>(stream),
                "bytes" => importer.ImportSheetFromBytes<Row>(bytes),
                "bytes-async" => await importer.ImportSheetFromBytesAsync<Row>(bytes),
                "file" => importer.ImportSheetFromFile<Row>(path),
                _ => await importer.ImportSheetFromFileAsync<Row>(path)
            };
            Assert.True(stream.CanRead);
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
        Assert.Equal(new[] { 7, 9 }, result.Items.Select(item => item.Count));
        Assert.Equal(new[] { 1, 2 }, result.SourceRows);
        Assert.Empty(result.Errors);
        Assert.True(result.IsSuccess);
        Assert.True(result.IsComplete);
        Assert.Equal(ImportStopReason.None, result.StopReason);
        Assert.False(result.ErrorsTruncated);
        Assert.Null(result.MaxErrors);
    }

    /// <summary>六入口在两种格式下都应保留缺列错误的全部结构化信息。</summary>
    [Theory]
    [MemberData(nameof(Entrypoints))]
    public async Task ImportSheet_AllEntrypoints_ShouldPreserveInvalidHeader(string format, string entry)
    {
        var result = await Invoke(new NpoiExcelImporter(), Workbook(format, "Wrong", "7"), entry);
        Assert.Empty(result.Items);
        Assert.Empty(result.SourceRows);
        Assert.False(result.IsSuccess);
        Assert.False(result.IsComplete);
        Assert.Equal(ImportStopReason.InvalidHeader, result.StopReason);
        Assert.False(result.ErrorsTruncated);
        Assert.Null(result.MaxErrors);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.InvalidHeader, error.Code);
        Assert.Equal("导入的表格不存在列：Count", error.Message);
        Assert.Equal("Data", error.SheetName);
        Assert.Equal(1, error.RowIndex);
        Assert.Equal(0, error.ColumnIndex);
        Assert.Null(error.PropertyName);
        Assert.Null(error.ColumnKey);
        Assert.Null(error.Header);
        Assert.Null(error.RawValue);
        Assert.Null(error.FirstRowNumber);
    }

    /// <summary>映射、名称选择和资源选项必须进入原管线，失败后下一请求仍可成功。</summary>
    [Theory]
    [InlineData("xls", false)]
    [InlineData("xlsx", false)]
    [InlineData("xls", true)]
    [InlineData("xlsx", true)]
    public async Task ImportSheet_OptionsAndMissingSheet_ShouldPreserveWorkbookDiagnostics(string format, bool asynchronous)
    {
        IExcelImporter importer = new NpoiExcelImporter();
        var bytes = Workbook(format, "数量", "7", "bad", "9");
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(Row.Count), Title = "数量" }
            }
        };
        var limits = new ExcelResourceLimits { MaxErrors = 1 };
        using var source = new MemoryStream(bytes);
        var result = asynchronous
            ? await importer.ImportSheetAsync<Row>(source, ExcelSheetSelector.ByName("Data"), s => s.Mapping(mapping), limits)
            : importer.ImportSheet<Row>(source, ExcelSheetSelector.ByName("Data"), s => s.Mapping(mapping), limits);
        Assert.Equal(new[] { 7 }, result.Items.Select(item => item.Count));
        Assert.Equal(new[] { 1 }, result.SourceRows);
        Assert.True(result.ErrorsTruncated);
        Assert.Equal(1, result.MaxErrors);
        Assert.False(result.IsComplete);
        Assert.Equal(ImportStopReason.ErrorLimit, result.StopReason);
        var error = Assert.Single(result.Errors);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, error.Code);
        Assert.Equal(3, error.RowIndex);
        Assert.Equal(1, error.ColumnIndex);
        Assert.Equal(nameof(Row.Count), error.PropertyName);
        Assert.Equal("数量", error.Header);
        Assert.Equal("bad", error.RawValue);
        Assert.True(source.CanRead);

        source.Position = 0;
        var missing = asynchronous
            ? await importer.ImportSheetAsync<Row>(source, ExcelSheetSelector.ByName("Missing"))
            : importer.ImportSheet<Row>(source, ExcelSheetSelector.ByName("Missing"));
        Assert.Empty(missing.Items);
        Assert.Empty(missing.SourceRows);
        var missingError = Assert.Single(missing.Errors);
        Assert.Equal(ExcelImportErrorCode.InvalidHeader, missingError.Code);
        Assert.Equal("缺少请求的 Sheet: Missing", missingError.Message);
        Assert.Equal("Missing", missingError.SheetName);
        Assert.Equal(0, missingError.RowIndex);
        Assert.Equal(0, missingError.ColumnIndex);
        Assert.Null(missingError.PropertyName);
        Assert.Equal(ImportStopReason.InvalidHeader, missing.StopReason);

        source.Position = 0;
        var recovered = asynchronous
            ? await importer.ImportSheetAsync<Row>(source, configure: s => s.Mapping(mapping))
            : importer.ImportSheet<Row>(source, configure: s => s.Mapping(mapping));
        Assert.Equal(new[] { 7, 9 }, recovered.Items.Select(item => item.Count));
        Assert.Equal(new[] { 1, 3 }, recovered.SourceRows);
        Assert.True(recovered.IsComplete);
        Assert.False(recovered.IsSuccess);
        Assert.False(recovered.ErrorsTruncated);
        Assert.Equal(ImportStopReason.None, recovered.StopReason);
        Assert.Single(recovered.Errors);
    }

    /// <summary>所有便利入口均原样传播取消和 Provider 异常，不把异常包装成成功结果。</summary>
    [Theory]
    [MemberData(nameof(Entrypoints))]
    public async Task ImportSheet_AllEntrypoints_ShouldPreserveCancellationAndProviderException(string format, string entry)
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var token = cancellation.Token;
        var bytes = Workbook(format, "Count", "7");
        var cancelled = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            Invoke(new NpoiExcelImporter(), bytes, entry, token));
        Assert.Equal(token, cancelled.CancellationToken);

        var expected = new IOException("provider failure");
        var provider = new ThrowingImporter(expected);
        var actual = await Assert.ThrowsAsync<IOException>(() => Invoke(provider, bytes, entry));
        Assert.Same(expected, actual);
        Assert.Equal(entry.EndsWith("async", StringComparison.Ordinal), provider.AsyncCalled);
        Assert.Equal(1, provider.SheetCount);
    }

    private static async Task<ExcelSingleSheetImportResult<Row>> Invoke(IExcelImporter importer, byte[] bytes,
        string entry, CancellationToken cancellationToken = default)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        if (entry == "stream" || entry == "stream-async")
        {
            try
            {
                return entry == "stream"
                    ? importer.ImportSheet<Row>(stream, cancellationToken: cancellationToken)
                    : await importer.ImportSheetAsync<Row>(stream, cancellationToken: cancellationToken);
            }
            finally { Assert.True(stream.CanRead); }
        }
        if (entry == "bytes")
            return importer.ImportSheetFromBytes<Row>(bytes, cancellationToken: cancellationToken);
        if (entry == "bytes-async")
            return await importer.ImportSheetFromBytesAsync<Row>(bytes, cancellationToken: cancellationToken);
        var path = Path.Combine(Path.GetTempPath(), $"bing-single-sheet-{Guid.NewGuid():N}.bin");
        File.WriteAllBytes(path, bytes);
        try
        {
            return entry == "file"
                ? importer.ImportSheetFromFile<Row>(path, cancellationToken: cancellationToken)
                : await importer.ImportSheetFromFileAsync<Row>(path, cancellationToken: cancellationToken);
        }
        finally
        {
            // Windows 上独占打开验证文件入口没有遗留源文件句柄。
            using (File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None)) { }
            File.Delete(path);
        }
    }

    private static byte[] Workbook(string format, string header, params string[] values)
    {
        using IWorkbook workbook = format == "xls" ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("Data");
        sheet.CreateRow(0).CreateCell(0).SetCellValue(header);
        for (var index = 0; index < values.Length; index++)
            sheet.CreateRow(index + 1).CreateCell(0).SetCellValue(values[index]);
        var second = workbook.CreateSheet("Other");
        second.CreateRow(0).CreateCell(0).SetCellValue("Count");
        second.CreateRow(1).CreateCell(0).SetCellValue(42);
        using var stream = new MemoryStream();
        workbook.Write(stream, false);
        return stream.ToArray();
    }

    private sealed class Row { public int Count { get; set; } }

    private sealed class ThrowingImporter : IExcelImporter
    {
        private readonly Exception _exception;
        internal ThrowingImporter(Exception exception) => _exception = exception;
        internal bool AsyncCalled { get; private set; }
        internal int SheetCount { get; private set; }
        public ExcelWorkbookImportResult<TWorkbook> Import<TWorkbook>(Stream source,
            ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken = default)
            where TWorkbook : class, new()
        {
            SheetCount = request.SheetCount;
            throw _exception;
        }
        public Task<ExcelWorkbookImportResult<TWorkbook>> ImportAsync<TWorkbook>(Stream source,
            ExcelWorkbookImportRequest<TWorkbook> request, CancellationToken cancellationToken = default)
            where TWorkbook : class, new()
        {
            AsyncCalled = true;
            SheetCount = request.SheetCount;
            return Task.FromException<ExcelWorkbookImportResult<TWorkbook>>(_exception);
        }
    }
}
