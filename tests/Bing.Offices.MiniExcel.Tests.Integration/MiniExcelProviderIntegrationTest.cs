using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.MiniExcel.Extensions;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.MiniExcel.Tests.Integration;

/// <summary>
/// 使用真实文件路径验证 MiniExcel Provider 的包消费者边界。
/// </summary>
public sealed class MiniExcelProviderIntegrationTest
{
    /// <summary>
    /// 验证导出到文件后再从文件导入时可往返多个工作表和日期。
    /// </summary>
    [Fact]
    public void ExportToFile_ThenImportFromFile_ShouldRoundTripMultipleSheetsAndDates()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-miniexcel-{Guid.NewGuid():N}.xlsx");
        using var provider = new ServiceCollection()
            .AddBingOfficesMiniExcel()
            .BuildServiceProvider();
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("People", new[] { new IntegrationRow { Name = "同步文件", Count = 3, Date = new DateTime(2026, 9, 17) } })
            .AddSheet("Summary", new[] { new SummaryRow { Label = "rows", Value = 1 } }));
        var importRequest = ExcelImport.Workbook<IntegrationWorkbook>(workbook => workbook
            .Sheet<IntegrationRow>("People", root => root.Rows)
            .Sheet<SummaryRow>("Summary", root => root.Summaries));

        try
        {
            provider.GetRequiredService<IExcelExporter>().ExportToFile(request, path);
            var result = ExcelStreamExtensions.ImportFromFile<IntegrationWorkbook>(
                provider.GetRequiredService<IExcelImporter>(), path, importRequest);

            Assert.Empty(result.Errors);
            var row = Assert.Single(result.Workbook.Rows);
            Assert.Equal("同步文件", row.Name);
            Assert.Equal(3, row.Count);
            Assert.Equal(new DateTime(2026, 9, 17), row.Date);
            var summary = Assert.Single(result.Workbook.Summaries);
            Assert.Equal("rows", summary.Label);
            Assert.Equal(1, summary.Value);
            Assert.True(new FileInfo(path).Length > 0);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 验证动态列导出到文件后可通过真实路径往返。
    /// </summary>
    [Fact]
    public void ExportToFile_DynamicColumns_ShouldRoundTripThroughRealPath()
    {
        var directory = Path.Combine(Path.GetTempPath(), "Bing.Offices.MiniExcel", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "dynamic.xlsx");
        var mapping = new ExcelMappingConfiguration
        {
            Columns =
            {
                new ExcelColumnConfiguration { PropertyName = nameof(DynamicIntegrationRow.Name), Title = "Name" }
            }
        };
        var dynamicColumn = new ExcelDynamicColumnDefinition
        {
            Key = "region",
            Title = "Region",
            DataType = typeof(string)
        };
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new DynamicIntegrationRow { Name = "动态文件", Values = new Dictionary<string, object> { ["region"] = "east" } } },
            sheet => sheet.Mapping(mapping).DynamicColumns(row => row.Values, new[] { dynamicColumn })));
        var importRequest = ExcelImport.Workbook<DynamicIntegrationWorkbook>(workbook =>
            workbook.Sheet<DynamicIntegrationRow>("Data", root => root.Rows,
                sheet => sheet.Mapping(mapping).DynamicColumns(row => row.Values, new[] { dynamicColumn })));
        using var provider = new ServiceCollection().AddBingOfficesMiniExcel().BuildServiceProvider();

        try
        {
            provider.GetRequiredService<IExcelExporter>().ExportToFile(request, path);
            var result = ExcelStreamExtensions.ImportFromFile<DynamicIntegrationWorkbook>(
                provider.GetRequiredService<IExcelImporter>(), path, importRequest);

            Assert.Empty(result.Errors);
            var row = Assert.Single(result.Workbook.Rows);
            Assert.Equal("动态文件", row.Name);
            Assert.Equal("east", row.Values["region"]?.ToString());
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 验证导出到已有文件目标时会替换旧内容并成功写入。
    /// </summary>
    [Fact]
    public void ExportToFile_ExistingTarget_ShouldReplaceAfterSuccessfulWrite()
    {
        var directory = Path.Combine(Path.GetTempPath(), "Bing.Offices.MiniExcel", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "replace.xlsx");
        using var provider = new ServiceCollection().AddBingOfficesMiniExcel().BuildServiceProvider();
        var first = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new IntegrationRow { Name = "旧内容", Count = 1 } }));
        var replacement = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new IntegrationRow { Name = "新内容", Count = 2 } }));
        var importRequest = ExcelImport.Workbook<IntegrationWorkbook>(workbook =>
            workbook.Sheet<IntegrationRow>("Data", root => root.Rows));

        try
        {
            provider.GetRequiredService<IExcelExporter>().ExportToFile(first, path);
            provider.GetRequiredService<IExcelExporter>().ExportToFile(replacement, path);
            var result = ExcelStreamExtensions.ImportFromFile<IntegrationWorkbook>(
                provider.GetRequiredService<IExcelImporter>(), path, importRequest);
            var row = Assert.Single(result.Workbook.Rows);
            Assert.Empty(result.Errors);
            Assert.Equal("新内容", row.Name);
            Assert.Equal(2, row.Count);
            Assert.Empty(Directory.GetFiles(directory, "replace.xlsx.*.tmp"));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 验证文件导出发生非取消失败时会保留现有目标并清理临时文件。
    /// </summary>
    [Fact]
    public void ExportToFile_NonCancellationFailure_ShouldPreserveExistingTargetAndCleanTemporaryFile()
    {
        var directory = Path.Combine(Path.GetTempPath(), "Bing.Offices.MiniExcel", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "failure.xlsx");
        var original = Encoding.UTF8.GetBytes("existing-target");
        File.WriteAllBytes(path, original);
        var converter = new ThrowingExportConverter();
        var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
            new[] { new IntegrationRow { Name = "触发失败", Count = 3 } }, sheet => sheet.Mapping(
                new ExcelMappingConfiguration
                {
                    Columns =
                    {
                        new ExcelColumnConfiguration
                        {
                            PropertyName = nameof(IntegrationRow.Name), Title = "Name",
                            ConverterName = converter.Name
                        }
                    }
                })));
        var exporter = new Bing.Offices.Exports.MiniExcelExcelExporter(
            valueConverters: new IExcelValueConverter[] { converter });

        try
        {
            var exception = Assert.Throws<BingOfficesExportException>(() => exporter.ExportToFile(request, path));
            Assert.Equal(BingOfficesErrorCode.ExportFailed, exception.Code);
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Empty(Directory.GetFiles(directory, "failure.xlsx.*.tmp"));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 验证异步文件导出预先取消时会保留现有目标并清理临时文件。
    /// </summary>
    [Fact]
    public async Task AsyncFileExport_PreCanceled_ShouldPreserveExistingTargetAndCleanTemporaryFile()
    {
        var directory = Path.Combine(Path.GetTempPath(), "Bing.Offices.MiniExcel", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "target.xlsx");
        var original = Encoding.UTF8.GetBytes("before");
        File.WriteAllBytes(path, original);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var provider = new ServiceCollection()
            .AddBingOfficesMiniExcel()
            .BuildServiceProvider();
        try
        {
            var request = ExcelExport.Workbook(workbook => workbook
                .AddSheet("Data", new[] { new IntegrationRow { Name = "cancel", Count = 1 } }));

            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                provider.GetRequiredService<IExcelExporter>()
                    .ExportToFileAsync(request, path, cancellation.Token));

            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Empty(Directory.GetFiles(directory, "target.xlsx.*.tmp"));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }

    }

    /// <summary>
    /// 验证异步文件导出期间取消时会保留现有目标并清理临时文件。
    /// </summary>
    [Fact]
    public async Task AsyncFileExport_MidFlightCancellation_ShouldPreserveExistingTargetAndCleanTemporaryFile()
    {
        var directory = Path.Combine(Path.GetTempPath(), "Bing.Offices.MiniExcel", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "target.xlsx");
        var original = Encoding.UTF8.GetBytes("before-mid-flight");
        File.WriteAllBytes(path, original);
        using var cancellation = new CancellationTokenSource();
        var converter = new CancelAfterFirstExportConverter(cancellation);
        try
        {
            var request = ExcelExport.Workbook(workbook => workbook.AddSheet("Data",
                Enumerable.Range(0, 1000).Select(index => new IntegrationRow
                {
                    Name = $"row-{index}",
                    Count = index,
                    Date = new DateTime(2026, 1, 1).AddDays(index)
                }), sheet => sheet.Mapping(new ExcelMappingConfiguration
                {
                    Columns =
                    {
                        new ExcelColumnConfiguration
                        {
                            PropertyName = nameof(IntegrationRow.Name),
                            Title = "Name",
                            ConverterName = converter.Name
                        }
                    }
                })));

            var exporter = new Bing.Offices.Exports.MiniExcelExcelExporter(
                valueConverters: new IExcelValueConverter[] { converter });
            await Assert.ThrowsAsync<OperationCanceledException>(() =>
                exporter.ExportToFileAsync(request, path, cancellation.Token));

            Assert.True(converter.ExportCalls > 0);
            Assert.Equal(original, File.ReadAllBytes(path));
            Assert.Empty(Directory.GetFiles(directory, "target.xlsx.*.tmp"));
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 验证异步文件导出后再异步导入时可往返真实 XLSX 文件。
    /// </summary>
    [Fact]
    public async Task ExportToFileAsync_ThenImportFromFileAsync_ShouldRoundTripRealXlsx()
    {
        var path = Path.Combine(Path.GetTempPath(), $"bing-offices-miniexcel-{Guid.NewGuid():N}.xlsx");
        using var provider = new ServiceCollection()
            .AddBingOfficesMiniExcel()
            .BuildServiceProvider();
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Data", new[] { new IntegrationRow { Name = "真实文件", Count = 9 } }));
        var importRequest = ExcelImport.Workbook<IntegrationWorkbook>(workbook => workbook
            .Sheet<IntegrationRow>("Data", root => root.Rows));

        try
        {
            await provider.GetRequiredService<IExcelExporter>().ExportToFileAsync(request, path);
            var result = await ExcelStreamExtensions.ImportFromFileAsync<IntegrationWorkbook>(
                provider.GetRequiredService<IExcelImporter>(), path, importRequest);

            var row = Assert.Single(result.Workbook.Rows);
            Assert.Empty(result.Errors);
            Assert.Equal("真实文件", row.Name);
            Assert.Equal(9, row.Count);
            Assert.True(new FileInfo(path).Length > 0);
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }
    }

    /// <summary>
    /// 验证 MiniExcel 大文件导入导出往返。
    /// </summary>
    /// <param name="rowCount">往返测试的数据行数。</param>
    /// <remarks>覆盖 500K 和 1M 行；可通过 Category=Large 筛选运行。</remarks>
    [Trait("Category", "Large")]
    [Theory]
    [InlineData(500_000)]
    [InlineData(1_000_000)]
    public void LargeFileRoundTrip_ShouldUseRealPath(int rowCount)
    {
        var directory = Path.Combine(Path.GetTempPath(), $"Bing.Offices.MiniExcel.Large.{Guid.NewGuid():N}");
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, "large.xlsx");
        using var provider = new ServiceCollection()
            .AddBingOfficesMiniExcel()
            .BuildServiceProvider();
        var rows = Enumerable.Range(0, rowCount).Select(index => new IntegrationRow
        {
            Name = $"large-{index}",
            Count = index,
            Date = new DateTime(2026, 1, 1).AddDays(index % 365)
        });
        var request = ExcelExport.Workbook(builder => builder.AddSheet("Data", rows));
        var importRequest = ExcelImport.Workbook<IntegrationWorkbook>(builder =>
            builder.ResourceLimits(new ExcelResourceLimits
            {
                MaxInputBytes = null,
                MaxZipEntryUncompressedBytes = null,
                MaxZipTotalUncompressedBytes = null,
                MaxXmlCharacters = null,
                MaxSharedStringsBytes = null,
                MaxWorksheetBytes = null,
                MaxTotalWorksheetBytes = null
            }).Sheet<IntegrationRow>("Data", root => root.Rows));

        try
        {
            provider.GetRequiredService<IExcelExporter>().ExportToFile(request, path);
            var result = ExcelStreamExtensions.ImportFromFile<IntegrationWorkbook>(
                provider.GetRequiredService<IExcelImporter>(), path, importRequest);

            Assert.Empty(result.Errors);
            Assert.Equal(rowCount, result.Workbook.Rows.Count);
            Assert.Equal("large-0", result.Workbook.Rows[0].Name);
            Assert.Equal(0, result.Workbook.Rows[0].Count);
            Assert.Equal(new DateTime(2026, 1, 1), result.Workbook.Rows[0].Date);
            Assert.Equal($"large-{rowCount - 1}", result.Workbook.Rows[rowCount - 1].Name);
            Assert.Equal(rowCount - 1, result.Workbook.Rows[rowCount - 1].Count);
            Assert.Equal(new DateTime(2026, 1, 1).AddDays((rowCount - 1) % 365),
                result.Workbook.Rows[rowCount - 1].Date);
        }
        finally
        {
            if (Directory.Exists(directory))
                Directory.Delete(directory, true);
        }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class IntegrationWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<IntegrationRow> Rows { get; } = new();
        /// <summary>
        /// 获取摘要集合。
        /// </summary>
        public List<SummaryRow> Summaries { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class IntegrationRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Count { get; set; }
        /// <summary>
        /// 获取或设置日期。
        /// </summary>
        public DateTime Date { get; set; }
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class SummaryRow
    {
        /// <summary>
        /// 获取或设置标签。
        /// </summary>
        public string Label { get; set; }
        /// <summary>
         /// 获取或设置摘要行数值。
        /// </summary>
        public int Value { get; set; }
    }

    /// <summary>
    /// 表示 Excel 测试使用的工作簿数据模型。
    /// </summary>
    private sealed class DynamicIntegrationWorkbook
    {
        /// <summary>
        /// 获取数据行集合。
        /// </summary>
        public List<DynamicIntegrationRow> Rows { get; } = new();
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class DynamicIntegrationRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置值集合。
        /// </summary>
        public Dictionary<string, object> Values { get; set; } = new();
    }

    /// <summary>
    /// 提供测试场景使用的值转换器。
    /// </summary>
    private sealed class ThrowingExportConverter : INamedExcelValueConverter
    {
        /// <inheritdoc />
        public string Name => "throwing-file-export";
        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = context.Value;
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            throw new InvalidOperationException("集成测试故意触发导出失败。");
        }
    }

    /// <summary>
    /// 提供测试场景使用的值转换器。
    /// </summary>
    private sealed class CancelAfterFirstExportConverter : INamedExcelValueConverter
    {
        /// <summary>
        /// 用于在首次导出后取消操作的令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation;

        /// <summary>
        /// 指示是否已触发首次导出后的取消。
        /// </summary>
        private bool _cancelled;

        /// <summary>
        /// 初始化一个 <see cref="CancelAfterFirstExportConverter" /> 类型的实例。
        /// </summary>
        /// <param name="cancellation">首次导出转换时触发的取消源。</param>
        public CancelAfterFirstExportConverter(CancellationTokenSource cancellation) => _cancellation = cancellation;

        /// <inheritdoc />
        public string Name => "cancel-after-first";
        /// <summary>
        /// 获取或设置导出调用次数。
        /// </summary>
        public int ExportCalls { get; private set; }
        /// <inheritdoc />
        public bool CanConvert(Type propertyType) => propertyType == typeof(string);

        /// <inheritdoc />
        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            value = context.Value;
            return true;
        }

        /// <inheritdoc />
        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            ExportCalls++;
            value = context.Value;
            if (!_cancelled)
            {
                _cancelled = true;
                _cancellation.Cancel();
            }

            return true;
        }
    }
}
