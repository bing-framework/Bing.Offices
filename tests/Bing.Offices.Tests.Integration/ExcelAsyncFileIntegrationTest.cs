using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Npoi.Extensions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests.Integration;

/// <summary>
/// 使用真实文件路径验证 Excel 异步导入导出边界。
/// </summary>
public sealed class ExcelAsyncFileIntegrationTest
{
    /// <summary>
    /// 测试 - XLS/XLSX 异步文件导出后应可立即通过异步文件导入，并释放文件句柄。
    /// </summary>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public async Task ExportToFileAsync_ThenImportFromFileAsync_ShouldRoundTripThroughRealPath(
        ExcelFormat format)
    {
        var directory = CreateTemporaryDirectory();
        var path = CreatePath(directory, format);
        var (exporter, importer) = CreateComponents(directory);

        try
        {
            var request = CreateExportRequest(format, "异步路径", 7);
            var stagingBefore = GetExcelStagingFiles(directory);

            await exporter.ExportToFileAsync(request, path);
            var result = await ExcelStreamExtensions.ImportFromFileAsync<AsyncFileWorkbook<AsyncFileRow>>(
                importer, path, CreateImportRequest());

            var item = Assert.Single(result.Workbook.Rows);
            Assert.Empty(result.Errors);
            Assert.Equal("异步路径", item.Name);
            Assert.Equal(7, item.Count);
            AssertFileCanBeOpenedExclusively(path);
            Assert.Empty(GetAtomicTempFiles(path));
            Assert.Equal(stagingBefore.OrderBy(item => item), GetExcelStagingFiles(directory).OrderBy(item => item));
        }
        finally
        {
            DeleteTemporaryDirectory(directory);
        }
    }

    /// <summary>
    /// 测试 - 异步导出到已有 XLS/XLSX 目标时应原子替换旧内容。
    /// </summary>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public async Task ExportToFileAsync_ExistingTarget_ShouldReplaceItAtomically(ExcelFormat format)
    {
        var directory = CreateTemporaryDirectory();
        var path = CreatePath(directory, format);
        var (exporter, importer) = CreateComponents(directory);

        try
        {
            await exporter.ExportToFileAsync(
                CreateExportRequest(format, "旧内容", 1), path);
            var originalBytes = File.ReadAllBytes(path);
            var stagingBefore = GetExcelStagingFiles(directory);

            await exporter.ExportToFileAsync(
                CreateExportRequest(format, "新内容", 2), path);
            var replacedBytes = File.ReadAllBytes(path);
            var result = await ExcelStreamExtensions.ImportFromFileAsync<AsyncFileWorkbook<AsyncFileRow>>(
                importer, path, CreateImportRequest());

            Assert.False(originalBytes.SequenceEqual(replacedBytes));
            var item = Assert.Single(result.Workbook.Rows);
            Assert.Empty(result.Errors);
            Assert.Equal("新内容", item.Name);
            Assert.Equal(2, item.Count);
            AssertFileCanBeOpenedExclusively(path);
            Assert.Empty(GetAtomicTempFiles(path));
            Assert.Equal(stagingBefore.OrderBy(item => item), GetExcelStagingFiles(directory).OrderBy(item => item));
        }
        finally
        {
            DeleteTemporaryDirectory(directory);
        }
    }

    /// <summary>
    /// 测试 - 预取消已有目标的异步导出时应保留原文件并清理临时文件。
    /// </summary>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public async Task ExportToFileAsync_PreCancelled_ShouldPreserveOriginalAndCleanTemps(ExcelFormat format)
    {
        var directory = CreateTemporaryDirectory();
        var path = CreatePath(directory, format);
        var (exporter, _) = CreateComponents(directory);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        try
        {
            var originalBytes = await WriteOriginalTargetAsync(exporter, format, path);
            var stagingBefore = GetExcelStagingFiles(directory);
            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                exporter.ExportToFileAsync(
                    CreateExportRequest(format, "取消后的新内容", 9), path, cancellation.Token));

            Assert.Equal(originalBytes, File.ReadAllBytes(path));
            Assert.Empty(GetAtomicTempFiles(path));
            Assert.Equal(stagingBefore.OrderBy(item => item), GetExcelStagingFiles(directory).OrderBy(item => item));
        }
        finally
        {
            DeleteTemporaryDirectory(directory);
        }
    }

    /// <summary>
    /// 测试 - 导出已开始写入后取消时，应保留已有目标并清理两层临时文件。
    /// </summary>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public async Task ExportToFileAsync_MidExportCancellation_ShouldPreserveOriginalAndCleanTemps(
        ExcelFormat format)
    {
        var directory = CreateTemporaryDirectory();
        var path = CreatePath(directory, format);
        var (exporter, _) = CreateComponents(directory);
        using var cancellation = new CancellationTokenSource();

        try
        {
            var originalBytes = await WriteOriginalTargetAsync(exporter, format, path);
            var stagingBefore = GetExcelStagingFiles(directory);
            var request = ExcelExport.Workbook(builder => builder
                .Format(format)
                .AddSheet("Data", CreateRowsThatCancelAfterFirstRow(cancellation)));

            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                exporter.ExportToFileAsync(request, path, cancellation.Token));

            Assert.Equal(originalBytes, File.ReadAllBytes(path));
            Assert.Empty(GetAtomicTempFiles(path));
            Assert.Equal(stagingBefore.OrderBy(item => item), GetExcelStagingFiles(directory).OrderBy(item => item));
        }
        finally
        {
            DeleteTemporaryDirectory(directory);
        }
    }

    /// <summary>
    /// 测试 - 导出内容异常时应保留已有目标并清理原子提交和 Excel staging 临时文件。
    /// </summary>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public async Task ExportToFileAsync_ExportFailure_ShouldPreserveOriginalAndCleanTemps(ExcelFormat format)
    {
        var directory = CreateTemporaryDirectory();
        var path = CreatePath(directory, format);
        var (exporter, _) = CreateComponents(directory);

        try
        {
            var originalBytes = await WriteOriginalTargetAsync(exporter, format, path);
            var stagingBefore = GetExcelStagingFiles(directory);
            var request = ExcelExport.Workbook(builder => builder
                .Format(format)
                .UseTemplate(new MemoryStream(CreateTemplateBytes(format, "Other")))
                .AddSheet("Data", new[] { new AsyncFileRow { Name = "不会写入", Count = 3 } }));

            await Assert.ThrowsAsync<BingOfficesConfigurationException>(() =>
                exporter.ExportToFileAsync(request, path));

            Assert.Equal(originalBytes, File.ReadAllBytes(path));
            Assert.Empty(GetAtomicTempFiles(path));
            Assert.Equal(stagingBefore.OrderBy(item => item), GetExcelStagingFiles(directory).OrderBy(item => item));
        }
        finally
        {
            DeleteTemporaryDirectory(directory);
        }
    }

    private static (IExcelExporter Exporter, IExcelImporter Importer) CreateComponents(string directory)
    {
        var stagingFactory = new NpoiAsyncStagingFactory(NpoiAsyncStagingStrategy.TempFile,
            directory: directory);
        return (
            new NpoiExcelExporter(new DefaultFileExportCommitter(), stagingFactory),
            new NpoiExcelImporter(stagingFactory));
    }

    private static string CreateTemporaryDirectory()
    {
        var directory = Path.Combine(Path.GetTempPath(), "Bing.Offices.Tests.Integration",
            Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        return directory;
    }

    private static string CreatePath(string directory, ExcelFormat format)
    {
        var extension = format == ExcelFormat.Xls ? "xls" : "xlsx";
        return Path.Combine(directory, $"data.{extension}");
    }

    private static ExcelWorkbookExportRequest CreateExportRequest(ExcelFormat format, string name, int count)
    {
        return ExcelExport.Workbook(builder => builder
            .Format(format)
            .AddSheet("Data", new[] { new AsyncFileRow { Name = name, Count = count } }));
    }

    private static ExcelWorkbookImportRequest<AsyncFileWorkbook<AsyncFileRow>> CreateImportRequest()
    {
        return ExcelImport.Workbook<AsyncFileWorkbook<AsyncFileRow>>(
            builder => builder.Sheet("Data", root => root.Rows));
    }

    private static async Task<byte[]> WriteOriginalTargetAsync(IExcelExporter exporter, ExcelFormat format,
        string path)
    {
        await exporter.ExportToFileAsync(
            CreateExportRequest(format, "原始内容", 1), path);
        return File.ReadAllBytes(path);
    }

    private static IEnumerable<AsyncFileRow> CreateRowsThatCancelAfterFirstRow(
        CancellationTokenSource cancellation)
    {
        yield return new AsyncFileRow { Name = "部分写入", Count = 1 };
        cancellation.Cancel();
    }

    private static byte[] CreateTemplateBytes(ExcelFormat format, string sheetName)
    {
        using IWorkbook workbook = format == ExcelFormat.Xls
            ? new HSSFWorkbook()
            : new XSSFWorkbook();
        workbook.CreateSheet(sheetName);
        using var destination = new MemoryStream();
        workbook.Write(destination, false);
        return destination.ToArray();
    }

    private static void AssertFileCanBeOpenedExclusively(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
        Assert.True(stream.Length > 0);
    }

    private static string[] GetAtomicTempFiles(string path)
    {
        var directory = Path.GetDirectoryName(path);
        var fileName = Path.GetFileName(path);
        return Directory.GetFiles(directory, fileName + ".*.tmp");
    }

    private static string[] GetExcelStagingFiles(string directory)
    {
        return Directory.GetFiles(directory, "bing-offices-excel-async-*.tmp");
    }

    private static void DeleteTemporaryDirectory(string directory)
    {
        if (Directory.Exists(directory))
            Directory.Delete(directory, true);
    }

    private sealed class AsyncFileWorkbook<T> where T : class, new()
    {
        public List<T> Rows { get; } = new();
    }

    private sealed class AsyncFileRow
    {
        public string Name { get; set; }
        public int Count { get; set; }
    }
}
