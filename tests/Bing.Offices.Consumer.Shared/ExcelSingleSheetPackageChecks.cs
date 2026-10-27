using Bing.Offices;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Exports;

namespace Bing.Offices.Consumer.Shared;

/// <summary>验证通过 NuGet 包公开的单表导入入口和完整结果契约。</summary>
public static class ExcelSingleSheetPackageChecks
{
    private static readonly ImportEntry[] Entries =
    {
        ImportEntry.Stream,
        ImportEntry.StreamAsync,
        ImportEntry.Bytes,
        ImportEntry.BytesAsync,
        ImportEntry.File,
        ImportEntry.FileAsync
    };

    /// <summary>用实际 Excel Provider 验证 XLS/XLSX 的六个单表入口。</summary>
    public static async Task RunAsync(IExcelImporter importer, IExcelExporter exporter, string directory)
    {
        var successCases = 0;
        var failureCases = 0;
        foreach (var format in new[] { ExcelFormat.Xls, ExcelFormat.Xlsx })
        {
            var extension = format == ExcelFormat.Xls ? ".xls" : ".xlsx";
            var validBytes = Export(exporter, format, "Data",
                new[] { new PackageSheetRow { Count = 7 }, new PackageSheetRow { Count = 9 } });
            var invalidBytes = Export(exporter, format, "Data",
                new[] { new MissingColumnRow { Other = 4 } });

            foreach (var entry in Entries)
            {
                var success = await InvokeAsync(importer, validBytes, directory, extension, entry);
                AssertSuccess(success);
                successCases++;

                var failure = await InvokeAsync(importer, invalidBytes, directory, extension, entry);
                AssertInvalidHeader(failure);
                failureCases++;
            }
        }

        Console.WriteLine($"package-single-sheet-ok formats=2 entrypoints=6 success={successCases} failure={failureCases}");
    }

    private static byte[] Export<T>(IExcelExporter exporter, ExcelFormat format, string name, IEnumerable<T> rows)
        where T : class, new()
    {
        var request = ExcelExport.Workbook(builder => builder
            .Format(format)
            .AddSheet(name, rows));
        return exporter.ExportToBytes(request);
    }

    private static async Task<ExcelSingleSheetImportResult<PackageSheetRow>> InvokeAsync(
        IExcelImporter importer, byte[] content, string directory, string extension, ImportEntry entry)
    {
        switch (entry)
        {
            case ImportEntry.Stream:
                using (var source = new MemoryStream(content, writable: false))
                {
                    var result = importer.ImportSheet<PackageSheetRow>(source);
                    Ensure(source.CanRead, "同步流入口关闭了调用方拥有的源流。");
                    return result;
                }
            case ImportEntry.StreamAsync:
                using (var source = new MemoryStream(content, writable: false))
                {
                    var result = await importer.ImportSheetAsync<PackageSheetRow>(source);
                    Ensure(source.CanRead, "异步流入口关闭了调用方拥有的源流。");
                    return result;
                }
            case ImportEntry.Bytes:
                return importer.ImportSheetFromBytes<PackageSheetRow>(content);
            case ImportEntry.BytesAsync:
                return await importer.ImportSheetFromBytesAsync<PackageSheetRow>(content);
            case ImportEntry.File:
            case ImportEntry.FileAsync:
                var path = Path.Combine(directory, $"single-sheet-{Guid.NewGuid():N}{extension}");
                File.WriteAllBytes(path, content);
                try
                {
                    return entry == ImportEntry.File
                        ? importer.ImportSheetFromFile<PackageSheetRow>(path)
                        : await importer.ImportSheetFromFileAsync<PackageSheetRow>(path);
                }
                finally
                {
                    File.Delete(path);
                }
            default:
                throw new ArgumentOutOfRangeException(nameof(entry), entry, "未知单表导入入口。");
        }
    }

    private static void AssertSuccess(ExcelSingleSheetImportResult<PackageSheetRow> result)
    {
        Ensure(result.Items.Select(item => item.Count).SequenceEqual(new[] { 7, 9 }),
            "单表导入未返回完整实体内容。");
        Ensure(result.SourceRows.SequenceEqual(new[] { 1, 2 }), "单表导入未保留完整源行索引。");
        Ensure(result.Errors.Count == 0, "成功导入意外包含错误。");
        Ensure(result.IsSuccess, "成功导入的 IsSuccess 状态不正确。");
        Ensure(result.IsComplete, "成功导入的 IsComplete 状态不正确。");
        Ensure(result.StopReason == ImportStopReason.None, "成功导入的 StopReason 不正确。");
        Ensure(!result.ErrorsTruncated && result.MaxErrors == null, "成功导入的错误截断状态不正确。");
    }

    private static void AssertInvalidHeader(ExcelSingleSheetImportResult<PackageSheetRow> result)
    {
        Ensure(result.Items.Count == 0, "表头错误结果不应包含实体。");
        Ensure(result.SourceRows.Count == 0, "表头错误结果不应包含源行。");
        Ensure(!result.IsSuccess, "表头错误结果不应成功。");
        Ensure(!result.IsComplete, "表头错误结果应标记为不完整。");
        Ensure(result.StopReason == ImportStopReason.InvalidHeader, "表头错误的 StopReason 不正确。");
        Ensure(!result.ErrorsTruncated && result.MaxErrors == null, "表头错误的截断状态不正确。");

        var error = result.Errors.Single();
        Ensure(error.Code == ExcelImportErrorCode.InvalidHeader, "表头错误 Code 不正确。");
        Ensure(error.Message == "导入的表格不存在列：Count", "表头错误 Message 不正确。");
        Ensure(error.SheetName == "Data", "表头错误 SheetName 不正确。");
        Ensure(error.RowIndex == 1 && error.ColumnIndex == 0, "表头错误行列位置不正确。");
        Ensure(error.PropertyName == null && error.ColumnKey == null, "表头错误不应关联目标属性。");
        Ensure(error.Header == null && error.RawValue == null && error.FirstRowNumber == null,
            "表头错误的空字段结构不正确。");
    }

    private static void Ensure(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
    }

    private enum ImportEntry
    {
        Stream,
        StreamAsync,
        Bytes,
        BytesAsync,
        File,
        FileAsync
    }

    private sealed class PackageSheetRow
    {
        public PackageSheetRow() { }

        public int Count { get; set; }
    }

    private sealed class MissingColumnRow
    {
        public MissingColumnRow() { }

        public int Other { get; set; }
    }
}
