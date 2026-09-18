using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Csv;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Validations;
using Bing.Offices.Extensions;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Exceptions;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// Review Fix 回归测试。
/// </summary>
public sealed class ReviewFixRegressionTest
{
    /// <summary>
    /// 测试 - Excel/CSV exporter 的文件提交失败必须在同一公共边界内观察同一异常实例且只观察一次。
    /// </summary>
    [Fact]
    public void ExporterFileCommitFailure_ShouldObserveSameExceptionOnce()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Data", new[] { new ReviewRow { Name = "A" } }));
        var excelObserver = new RecordingObserver();
        var excelException = new BingOfficesFileCommitException("Excel 提交失败", provider: "Test");
        var excelExporter = new NpoiExcelExporter(
            exceptionObservers: new[] { excelObserver },
            fileExportCommitter: new ThrowingCommitter(excelException));

        var actualExcel = Assert.Throws<BingOfficesFileCommitException>(() =>
            excelExporter.ExportToFile(request, "target.xlsx"));

        Assert.Same(excelException, actualExcel);
        Assert.Single(excelObserver.Exceptions);
        Assert.Same(actualExcel, excelObserver.Exceptions[0]);

        var csvObserver = new RecordingObserver();
        var csvException = new BingOfficesFileCommitException("CSV 提交失败", provider: "Test");
        var csvExporter = new CsvEntityExporter(
            exceptionObservers: new[] { csvObserver },
            fileExportCommitter: new ThrowingCommitter(csvException));

        var actualCsv = Assert.Throws<BingOfficesFileCommitException>(() =>
            csvExporter.ExportToFile(new[] { new ReviewRow { Name = "A" } }, "target.csv"));

        Assert.Same(csvException, actualCsv);
        Assert.Single(csvObserver.Exceptions);
        Assert.Same(actualCsv, csvObserver.Exceptions[0]);
    }
    /// <summary>
    /// 测试 - Core 默认注册应提供计划工厂，且 NPOI 注册不得覆盖调用方预注册的替换实现。
    /// </summary>
    [Fact]
    public void MappingPlanFactory_DiDefaultAndReplacement_ShouldPreserveOwnershipBoundary()
    {
        // Arrange
        var defaultServices = new ServiceCollection();
        defaultServices.AddBingOfficesNpoi();
        using var defaultProvider = defaultServices.BuildServiceProvider();
        var replacement = new ExcelMappingPlanFactory(cacheCapacity: 3);
        var replacementServices = new ServiceCollection();
        replacementServices.AddSingleton<IExcelMappingPlanFactory>(replacement);

        // Act
        replacementServices.AddBingOfficesNpoi();
        using var replacementProvider = replacementServices.BuildServiceProvider();

        // Assert
        Assert.IsType<ExcelMappingPlanFactory>(defaultProvider.GetRequiredService<IExcelMappingPlanFactory>());
        Assert.Same(replacement, replacementProvider.GetRequiredService<IExcelMappingPlanFactory>());
    }

    /// <summary>
    /// 表示测试使用的一行数据模型。
    /// </summary>
    private sealed class ReviewRow
    {
        /// <summary>
        /// 获取或设置名称。
        /// </summary>
        public string Name { get; set; }
    }

    /// <summary>
    /// 记录测试场景中的通知事件。
    /// </summary>
    private sealed class RecordingObserver : IBingOfficesExceptionObserver
    {
        /// <summary>
        /// 获取异常集合。
        /// </summary>
        public List<BingOfficesException> Exceptions { get; } = new();

        /// <inheritdoc />
        public void Observe(BingOfficesException exception) => Exceptions.Add(exception);
    }
    /// <summary>
    /// 提供会抛出提交异常的文件导出提交替身。
    /// </summary>
    private sealed class ThrowingCommitter : IFileExportCommitter
    {
        /// <summary>
        /// 文件提交时应抛出的测试异常。
        /// </summary>
        private readonly BingOfficesFileCommitException _exception;

        /// <summary>
        /// 初始化一个 <see cref="ThrowingCommitter" /> 类型的实例。
        /// </summary>
        /// <param name="exception">提交操作应抛出的异常。</param>
        public ThrowingCommitter(BingOfficesFileCommitException exception) => _exception = exception;

        /// <inheritdoc />
        public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken, string format)
            => throw _exception;

        /// <inheritdoc />
        public Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
            CancellationToken cancellationToken, string format) => Task.FromException(_exception);
    }
}
