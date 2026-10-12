using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.ClosedXml.Entities;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.ClosedXml.Tests;

/// <summary>
/// 验证 Buffered Entity 批次导入的失败工作簿交付顺序和错误预算。
/// </summary>
public sealed class EntityBatchFailureWorkbookTest
{
    /// <summary>
    /// 验证失败工作簿在第一个同步或异步批次回调前已提交。
    /// </summary>
    /// <param name="asynchronous">是否使用异步批次入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FailureWorkbook_ShouldBeCommittedBeforeFirstBatch(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = CreateInvalidWorkbook(1);
        using var provider = CreateProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var path = Path.Combine(directory.Path, "batch-failure.xlsx");
        var callbackCount = 0;
        var options = CreateOptions(directory.Path, path);
        var layout = CreateLayout();

        void AssertCommittedArtifact(ExcelEntityImportBatch<Root> batch)
        {
            callbackCount++;
            Assert.True(File.Exists(path));
            using var workbook = new XLWorkbook(path);
            Assert.True(workbook.TryGetWorksheet("Data", out var data));
            Assert.Equal("Code", data.Cell("A1").GetString());
            Assert.Equal(string.Empty, data.Cell("A2").GetString());
            Assert.True(workbook.TryGetWorksheet("_ImportErrors", out var errors));
            Assert.Equal(2, errors.LastRowUsed().RowNumber());
            Assert.NotEmpty(batch.Errors);
        }

        var summary = await Import(importer, source, layout, options, asynchronous,
            AssertCommittedArtifact);

        Assert.Equal(1, callbackCount);
        Assert.False(summary.IsSuccess);
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证批次回调异常保持原实例传播且失败工作簿已在回调前提交。
    /// </summary>
    /// <param name="asynchronous">是否使用异步批次入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CallbackException_ShouldPropagateOnceAfterCommit(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = CreateInvalidWorkbook(2);
        using var provider = CreateProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var path = Path.Combine(directory.Path, "batch-failure.xlsx");
        var expected = new InvalidOperationException("batch callback failed");
        var callbackCount = 0;
        var options = CreateOptions(directory.Path, path);
        var layout = CreateLayout();

        void Callback(ExcelEntityImportBatch<Root> _)
        {
            callbackCount++;
            Assert.True(File.Exists(path));
            throw expected;
        }

        var actual = await Assert.ThrowsAsync<InvalidOperationException>(() =>
            Import(importer, source, layout, options, asynchronous, Callback));

        Assert.Same(expected, actual);
        Assert.Equal(1, callbackCount);
        Assert.True(File.Exists(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 MaxErrors 同时限制完整结果、失败摘要和已交付批次错误。
    /// </summary>
    /// <param name="asynchronous">是否使用异步批次入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task MaxErrors_ShouldCapSummaryResultAndBatchErrors(bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = CreateInvalidWorkbook(3);
        using var provider = CreateProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var path = Path.Combine(directory.Path, "batch-failure.xlsx");
        var batches = new List<ExcelEntityImportBatch<Root>>();
        var importOptions = new ExcelEntityImportOptions(new ExcelResourceLimits { MaxErrors = 1 },
            ExcelValidationFailureMode.Continue, ExcelImportValidationMode.ConfiguredRules,
            ExcelUnsupportedFeaturePolicy.Fail, FailureOptions(directory.Path, path));
        var batchOptions = new ExcelEntityBatchImportOptions(1, importOptions);

        var summary = await Import(importer, source, CreateLayout(), batchOptions, asynchronous,
            batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.True(summary.ErrorsTruncated);
        Assert.True(summary.Result.ErrorsTruncated);
        Assert.Equal(1, summary.Result.MaxErrors);
        Assert.Equal(1, summary.ErrorCount);
        Assert.Equal(1, summary.Result.Errors.Count);
        Assert.Equal(1, batches.Sum(batch => batch.Errors.Count));
        Assert.Equal(summary.ErrorCount, batches.Sum(batch => batch.Errors.Count));
        Assert.Single(summary.Regions);
        Assert.True(summary.Regions[0].Summary.ErrorsTruncated);

        using var workbook = new XLWorkbook(path);
        var errors = workbook.Worksheet("_ImportErrors");
        Assert.Equal(2, errors.LastRowUsed().RowNumber());
        Assert.Single(summary.Result.Errors);
        Assert.Contains(summary.Result.Errors[0].Message,
            errors.CellsUsed().Select(cell => cell.GetString()));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
    }

    /// <summary>
    /// 创建解析和映射规则一致的 Buffered 批次测试服务。
    /// </summary>
    /// <returns>可解析 ClosedXML 实体批次导入器的服务提供器。</returns>
    private static ServiceProvider CreateProvider() =>
        new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();

    /// <summary>
    /// 创建把明细列表映射到 Data 工作表的布局。
    /// </summary>
    /// <returns>Data 工作表实体布局。</returns>
    private static ExcelEntityLayout<Root> CreateLayout() =>
        ExcelEntity.Layout<Root>(builder => builder.ListRegion("Data", "A1", root => root.Rows));

    /// <summary>
    /// 创建包含指定数量必填错误行的真实 XLSX 工作簿。
    /// </summary>
    /// <param name="invalidRows">空编码错误行数量。</param>
    /// <returns>定位在起始位置的 XLSX 流。</returns>
    private static MemoryStream CreateInvalidWorkbook(int invalidRows)
    {
        var stream = new MemoryStream();
        using (var workbook = new XLWorkbook())
        {
            var sheet = workbook.Worksheets.Add("Data");
            sheet.Cell("A1").Value = "Code";
            sheet.Cell("B1").Value = "Quantity";
            for (var index = 0; index < invalidRows; index++)
            {
                sheet.Cell(index + 2, 1).Value = string.Empty;
                sheet.Cell(index + 2, 2).Value = index + 10;
            }
            workbook.SaveAs(stream);
        }
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 创建向指定文件提交失败工作簿的批次选项。
    /// </summary>
    /// <param name="directory">暂存文件目录。</param>
    /// <param name="path">失败工作簿目标路径。</param>
    /// <returns>配置失败工作簿输出的批次选项。</returns>
    private static ExcelEntityBatchImportOptions CreateOptions(string directory, string path) =>
        new(1, new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            FailureOptions(directory, path)));

    /// <summary>
    /// 创建失败工作簿文件提交选项。
    /// </summary>
    /// <param name="directory">暂存文件目录。</param>
    /// <param name="path">失败工作簿目标路径。</param>
    /// <returns>失败工作簿选项。</returns>
    private static ExcelImportFailureOptions FailureOptions(string directory, string path) =>
        new()
        {
            Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
            DestinationPath = path,
            TemporaryDirectory = directory
        };

    /// <summary>
    /// 按同步或异步入口执行 Buffered 实体批次导入。
    /// </summary>
    /// <param name="importer">实体批次导入器。</param>
    /// <param name="source">真实 XLSX 来源流。</param>
    /// <param name="layout">实体布局。</param>
    /// <param name="options">批次与失败工作簿选项。</param>
    /// <param name="asynchronous">是否执行异步入口。</param>
    /// <param name="callback">接收批次的同步回调。</param>
    /// <returns>完整实体批次摘要。</returns>
    private static Task<ExcelEntityBatchImportSummary<Root>> Import(
        IExcelEntityBatchImporter importer, Stream source, ExcelEntityLayout<Root> layout,
        ExcelEntityBatchImportOptions options, bool asynchronous,
        Action<ExcelEntityImportBatch<Root>> callback) => asynchronous
            ? importer.ImportEntityBatchesAsync(source, layout, options,
                (batch, _) =>
                {
                    callback(batch);
                    return Task.CompletedTask;
                })
            : Task.FromResult(importer.ImportEntityBatches(source, layout, options, callback));

    /// <summary>
    /// 为一个批次失败工作簿测试创建并清理独立目录。
    /// </summary>
    private sealed class TestDirectory : IDisposable
    {
        /// <summary>
        /// 获取本测试使用的目录路径。
        /// </summary>
        public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
            "closedxml-entity-batch-failure-" + Guid.NewGuid().ToString("N"));

        /// <summary>
        /// 初始化一个 <see cref="TestDirectory" /> 类型的实例。
        /// </summary>
        public TestDirectory() => Directory.CreateDirectory(Path);

        /// <inheritdoc />
        public void Dispose() => Directory.Delete(Path, recursive: true);
    }

    /// <summary>
    /// 保存一个实体列表批次导入结果。
    /// </summary>
    private sealed class Root
    {
        /// <summary>
        /// 获取明细行集合。
        /// </summary>
        public List<Row> Rows { get; } = new();
    }

    /// <summary>
    /// 保存必填错误测试行。
    /// </summary>
    private sealed class Row
    {
        /// <summary>
        /// 获取或设置必填编码。
        /// </summary>
        [ExcelRequired]
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }
    }
}
