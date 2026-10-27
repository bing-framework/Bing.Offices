using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Bing.Offices.Attributes;
using Bing.Offices.Exports;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Styles;
using Microsoft.Extensions.DependencyInjection;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>低内存 XLSX 导出器的职责测试。</summary>
public class NpoiExcelStreamingExporterTest
{
    /// <summary>跨越多个行窗口后，所有固定列、动态列、样式和日期仍应完整可读。</summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Export_ShouldWriteAllRowsAcrossWindowsAndPreserveWorkbookContent(bool limitedText)
    {
        var primaryRows = new SingleEnumerationEnumerable<StreamingRow>(Enumerable.Range(0, 9)
            .Select(CreateStreamingRow).ToArray());
        var dateRows = new SingleEnumerationEnumerable<DateRow>(new[]
        {
            new DateRow { Name = "中文日期", CreatedOn = new DateTime(2025, 2, 3) }
        });
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("明细", primaryRows, sheet => sheet
                .DynamicColumns(row => row.DynamicValues, CreateDynamicDefinitions())
                .HeaderStyle(new ExcelCellStyle { Bold = true })
                .BodyStyle(new ExcelCellStyle { WrapText = true }))
            .AddSheet("日期", dateRows, sheet => sheet
                .HeaderStyle(new ExcelCellStyle { Bold = true })
                .BodyStyle(new ExcelCellStyle { NumberFormat = "yyyy-mm-dd" })));
        var temporaryRoot = CreateTemporaryDirectory();

        try
        {
            using var destination = new MemoryStream();
            new NpoiExcelStreamingExporter().Export(request, destination,
                new ExcelStreamingExportOptions { RowWindowSize = 2, TempDirectory = temporaryRoot,
                    MaxRowCharacters = limitedText ? 256 : 262144,
                    MaxBufferedCharacters = limitedText ? 256 : 2097152 });

            Assert.True(destination.CanWrite);
            Assert.Equal(1, primaryRows.EnumeratorCount);
            Assert.Equal(1, dateRows.EnumeratorCount);
            Assert.Empty(Directory.GetFileSystemEntries(temporaryRoot));

            using var workbookBytes = new MemoryStream(destination.ToArray());
            using var workbook = new XSSFWorkbook(workbookBytes);
            Assert.Equal(2, workbook.NumberOfSheets);
            Assert.Equal("明细", workbook.GetSheetAt(0).SheetName);
            Assert.Equal("日期", workbook.GetSheetAt(1).SheetName);

            var detail = workbook.GetSheet("明细");
            var headers = detail.GetRow(0);
            Assert.Equal((short)10, headers.LastCellNum);
            var codeColumn = FindColumn(headers, "Code");
            var createdColumn = FindColumn(headers, "CreatedOn");
            var dynamicColumns = Enumerable.Range(0, 8)
                .Select(index => FindColumn(headers, $"动态字段{index}"))
                .ToArray();

            for (var rowIndex = 0; rowIndex < 9; rowIndex++)
            {
                var row = detail.GetRow(rowIndex + 1);
                Assert.NotNull(row);
                Assert.Equal((short)10, row.LastCellNum);
                Assert.Equal($"编号-{rowIndex}", row.GetCell(codeColumn).StringCellValue);
                Assert.Equal(CellType.Numeric, row.GetCell(createdColumn).CellType);
                for (var dynamicIndex = 0; dynamicIndex < dynamicColumns.Length; dynamicIndex++)
                    Assert.Equal($"中文值-{rowIndex}-{dynamicIndex}",
                        row.GetCell(dynamicColumns[dynamicIndex]).StringCellValue);
            }

            Assert.True(workbook.GetFontAt(headers.GetCell(codeColumn).CellStyle.FontIndex).IsBold);
            Assert.True(detail.GetRow(1).GetCell(dynamicColumns[0]).CellStyle.WrapText);

            var dates = workbook.GetSheet("日期");
            var dateColumn = FindColumn(dates.GetRow(0), "CreatedOn");
            var dateCell = dates.GetRow(1).GetCell(dateColumn);
            Assert.True(DateUtil.IsCellDateFormatted(dateCell));
            Assert.Equal(DateUtil.GetExcelDate(new DateTime(2025, 2, 3)), dateCell.NumericCellValue, 8);
            Assert.Equal("中文日期", dates.GetRow(1).GetCell(FindColumn(dates.GetRow(0), "Name")).StringCellValue);
        }
        finally
        {
            Directory.Delete(temporaryRoot, true);
        }
    }

    /// <summary>显式注册应提供可解析的流式导出器。</summary>
    [Fact]
    public void AddBingOfficesNpoiStreaming_ShouldRegisterStreamingExporter()
    {
        var services = new ServiceCollection();

        var returned = services.AddBingOfficesNpoiStreaming();
        using var provider = services.BuildServiceProvider();

        Assert.Same(services, returned);
        Assert.IsType<NpoiExcelStreamingExporter>(provider.GetRequiredService<IExcelStreamingExporter>());
    }

    /// <summary>遇到不支持的模板配置时，应在取得数据枚举器前失败并保持模板流打开。</summary>
    [Fact]
    public void Export_WhenTemplateIsConfigured_ShouldRejectBeforeEnumeratingData()
    {
        using var template = new MemoryStream(new byte[] { 1, 2, 3 });
        var rows = new SingleEnumerationEnumerable<StreamingRow>(new[] { CreateStreamingRow(0) });
        var request = ExcelExport.Workbook(workbook => workbook
            .UseTemplate(template, leaveOpen: true)
            .AddSheet("明细", rows, sheet => sheet.DynamicColumns(row => row.DynamicValues,
                CreateDynamicDefinitions())));
        using var destination = new MemoryStream();

        var exception = Assert.Throws<NotSupportedException>(() =>
            new NpoiExcelStreamingExporter().Export(request, destination));

        Assert.Contains("模板", exception.Message);
        Assert.Equal(0, rows.EnumeratorCount);
        Assert.True(template.CanRead);
        Assert.Empty(destination.ToArray());
    }

    /// <summary>数据源失败时，文件出口应保留旧目标并清理本次调用的临时文件。</summary>
    [Fact]
    public void ExportToFile_WhenDataEnumerationFails_ShouldPreserveTargetAndCleanTemporaryFiles()
    {
        var roots = CreateFileAndSessionDirectories();
        var targetPath = Path.Combine(roots.FileDirectory, "report.xlsx");
        var original = Encoding.UTF8.GetBytes("原有文件内容");
        File.WriteAllBytes(targetPath, original);
        var request = CreateRequest(ThrowAfterFirstRow());

        try
        {
            var exception = Assert.Throws<InvalidOperationException>(() =>
                new NpoiExcelStreamingExporter().ExportToFile(request, targetPath,
                    new ExcelStreamingExportOptions { RowWindowSize = 1, TempDirectory = roots.SessionDirectory }));

            Assert.Equal("数据源失败", exception.Message);
            Assert.Equal(original, File.ReadAllBytes(targetPath));
            Assert.Empty(Directory.GetFileSystemEntries(roots.SessionDirectory));
            Assert.Equal(new[] { targetPath }, Directory.GetFiles(roots.FileDirectory));
        }
        finally
        {
            Directory.Delete(roots.Root, true);
        }
    }

    /// <summary>跨过行窗口后取消时，文件出口应保留旧目标并清理本次调用的临时文件。</summary>
    [Fact]
    public void ExportToFile_WhenCancelledAcrossWindow_ShouldPreserveTargetAndCleanTemporaryFiles()
    {
        var roots = CreateFileAndSessionDirectories();
        var targetPath = Path.Combine(roots.FileDirectory, "report.xlsx");
        var original = Encoding.UTF8.GetBytes("原有文件内容");
        File.WriteAllBytes(targetPath, original);
        using var cancellation = new CancellationTokenSource();
        var request = CreateRequest(CancelAfterRows(cancellation, 3));

        try
        {
            Assert.Throws<OperationCanceledException>(() => new NpoiExcelStreamingExporter()
                .ExportToFile(request, targetPath,
                    new ExcelStreamingExportOptions { RowWindowSize = 2, TempDirectory = roots.SessionDirectory },
                    cancellation.Token));

            Assert.Equal(original, File.ReadAllBytes(targetPath));
            Assert.Empty(Directory.GetFileSystemEntries(roots.SessionDirectory));
            Assert.Equal(new[] { targetPath }, Directory.GetFiles(roots.FileDirectory));
        }
        finally
        {
            Directory.Delete(roots.Root, true);
        }
    }

    private static ExcelWorkbookExportRequest CreateRequest(IEnumerable<StreamingRow> rows) =>
        ExcelExport.Workbook(workbook => workbook.AddSheet("明细", rows,
            sheet => sheet.DynamicColumns(row => row.DynamicValues, CreateDynamicDefinitions())));

    private static ExcelDynamicColumnDefinition[] CreateDynamicDefinitions() => Enumerable.Range(0, 8)
        .Select(index => new ExcelDynamicColumnDefinition
        {
            Key = $"extra-{index}",
            Title = $"动态字段{index}",
            DataType = typeof(string)
        }).ToArray();

    private static StreamingRow CreateStreamingRow(int index) => new()
    {
        Code = $"编号-{index}",
        CreatedOn = new DateTime(2025, 1, 1).AddDays(index),
        DynamicValues = Enumerable.Range(0, 8).ToDictionary(
            column => $"extra-{column}", column => (object)$"中文值-{index}-{column}")
    };

    private static IEnumerable<StreamingRow> ThrowAfterFirstRow()
    {
        yield return CreateStreamingRow(0);
        throw new InvalidOperationException("数据源失败");
    }

    private static IEnumerable<StreamingRow> CancelAfterRows(CancellationTokenSource cancellation, int completedRows)
    {
        for (var index = 0; index < 8; index++)
        {
            if (index == completedRows)
                cancellation.Cancel();
            yield return CreateStreamingRow(index);
        }
    }

    private static int FindColumn(IRow header, string title)
    {
        for (var index = 0; index < header.LastCellNum; index++)
            if (string.Equals(header.GetCell(index)?.StringCellValue, title, StringComparison.Ordinal))
                return index;
        throw new Xunit.Sdk.XunitException($"未找到列标题: {title}");
    }

    private static string CreateTemporaryDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), "Bing.Offices.StreamingTest." + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(path);
        return path;
    }

    private static (string Root, string FileDirectory, string SessionDirectory) CreateFileAndSessionDirectories()
    {
        var root = CreateTemporaryDirectory();
        var fileDirectory = Path.Combine(root, "target");
        var sessionDirectory = Path.Combine(root, "session");
        Directory.CreateDirectory(fileDirectory);
        Directory.CreateDirectory(sessionDirectory);
        return (root, fileDirectory, sessionDirectory);
    }

    private sealed class StreamingRow
    {
        public string Code { get; set; }
        public DateTime CreatedOn { get; set; }

        [DynamicColumn]
        public IDictionary<string, object> DynamicValues { get; set; }
    }

    private sealed class DateRow
    {
        public string Name { get; set; }
        public DateTime CreatedOn { get; set; }
    }

    private sealed class SingleEnumerationEnumerable<T> : IEnumerable<T>
    {
        private readonly IEnumerable<T> _items;

        internal SingleEnumerationEnumerable(IEnumerable<T> items) => _items = items;

        internal int EnumeratorCount { get; private set; }

        public IEnumerator<T> GetEnumerator()
        {
            EnumeratorCount++;
            if (EnumeratorCount > 1)
                throw new InvalidOperationException("数据源被重复枚举。");
            return _items.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
