using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests.Integration;

/// <summary>
/// ClosedXML 真实文件和外围异步 IO 集成测试。
/// </summary>
public sealed class ClosedXmlRealFileIntegrationTest
{
    /// <summary>
    /// 验证异步文件导出提交后可被 ClosedXML 重新打开并导入。
    /// </summary>
    [Fact]
    public async Task ExportToFileAsync_ShouldCommitReadableXlsxAndRoundTrip()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-closedxml-{Guid.NewGuid():N}.xlsx");
        try
        {
            var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
                new[] { new Row { Id = 7, Name = "real-file" } }));
            await new ClosedXmlExcelExporter().ExportToFileAsync(export, path);
            Assert.True(File.Exists(path));
            await using var input = File.OpenRead(path);
            using var workbook = new XLWorkbook(input);
            Assert.Equal("real-file", workbook.Worksheet("Rows").Cell(2, 2).GetString());
            input.Position = 0;
            var request = ExcelImport.Workbook<Root>(builder =>
                builder.Sheet<Row>("Rows", root => root.Rows));
            var result = await new ClosedXmlExcelImporter().ImportAsync(input, request);
            Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
            Assert.Equal(7, Assert.Single(result.Workbook.Rows).Id);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 验证异步导出预取消时不会替换既有目标文件。
    /// </summary>
    [Fact]
    public async Task CancelledAsyncExport_ShouldLeaveExistingFileUntouched()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-closedxml-cancel-{Guid.NewGuid():N}.xlsx");
        var original = new byte[] { 1, 2, 3, 4 };
        await File.WriteAllBytesAsync(path, original);
        try
        {
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
                new[] { new Row { Id = 1, Name = "cancelled" } }));
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                new ClosedXmlExcelExporter().ExportToFileAsync(export, path, cancellation.Token));
            Assert.Equal(original, await File.ReadAllBytesAsync(path));
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 验证同步文件导出能够替换既有文件并生成可读 XLSX。
    /// </summary>
    [Fact]
    public void ExportToFile_ShouldCommitReadableXlsxAndReplaceExistingFile()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-closedxml-sync-{Guid.NewGuid():N}.xlsx");
        try
        {
            File.WriteAllBytes(path, new byte[] { 1, 2, 3 });
            var export = ExcelExport.Workbook(workbook => workbook.AddSheet("Rows",
                new[] { new Row { Id = 8, Name = "sync-file" } }));
            new ClosedXmlExcelExporter().ExportToFile(export, path);

            using var input = File.OpenRead(path);
            using var workbook = new XLWorkbook(input);
            Assert.Equal("sync-file", workbook.Worksheet("Rows").Cell(2, 2).GetString());
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 验证生成失败时保留既有目标文件内容。
    /// </summary>
    [Fact]
    public void ExportToFile_ShouldKeepExistingFileWhenGenerationFails()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-closedxml-failure-{Guid.NewGuid():N}.xlsx");
        var original = new byte[] { 9, 8, 7, 6 };
        try
        {
            File.WriteAllBytes(path, original);
            var export = ExcelExport.Workbook(workbook => workbook
                .Format(ExcelFormat.Xls)
                .AddSheet("Rows", new[] { new Row { Id = 9, Name = "rejected" } }));

            Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
                new ClosedXmlExcelExporter().ExportToFile(export, path));
            Assert.Equal(original, File.ReadAllBytes(path));
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 表示真实文件往返测试的工作簿模型。
    /// </summary>
    private sealed class Root
    {
        /// <summary>
        /// 获取导入的数据行集合。
        /// </summary>
        public List<Row> Rows { get; } = new();
    }

    /// <summary>
    /// 表示真实文件往返测试的数据行。
    /// </summary>
    private sealed class Row
    {
        /// <summary>
        /// 获取或设置行编号。
        /// </summary>
        public int Id { get; set; }

        /// <summary>
        /// 获取或设置行名称。
        /// </summary>
        public string Name { get; set; }
    }
}
