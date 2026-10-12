using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Bing.Offices.Attributes;
using Bing.Offices.Entities;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Microsoft.Extensions.DependencyInjection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 NPOI Buffered Entity 批次与失败工作簿的交付顺序和错误截断。
/// </summary>
public sealed class EntityBatchFailureWorkbookTest
{
    /// <summary>
    /// 验证失败工作簿在首个批次回调前已经提交。
    /// </summary>
    /// <param name="xls">是否使用 XLS 格式。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void FailureWorkbook_ShouldBeReadyBeforeFirstBatchCallback(bool xls)
    {
        using var directory = new TestDirectory();
        using var source = new MemoryStream(CreateWorkbookBytes(xls), writable: false);
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var path = Path.Combine(directory.Path, xls ? "entity-failure.xls" : "entity-failure.xlsx");
        var importOptions = CreateImportOptions(directory.Path, path,
            new ExcelResourceLimits());
        var batches = new List<ExcelEntityImportBatch<Root>>();
        var callbackCount = 0;

        var summary = importer.ImportEntityBatches(source, CreateLayout(),
            new ExcelEntityBatchImportOptions(1, importOptions), batch =>
            {
                callbackCount++;
                Assert.True(File.Exists(path));
                using var failure = File.OpenRead(path);
                using var workbook = WorkbookFactory.Create(failure);
                Assert.NotNull(workbook.GetSheet("_ImportErrors"));
                batches.Add(batch);
            });

        Assert.False(summary.IsSuccess);
        Assert.NotEmpty(summary.Result.Errors);
        Assert.Equal(2, callbackCount);
        Assert.Equal(2, batches.Count);
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证批次回调异常原样传播且不会重复回调。
    /// </summary>
    [Fact]
    public void CallbackFailure_ShouldPropagateAfterFailureWorkbookIsWritten()
    {
        using var directory = new TestDirectory();
        using var source = new MemoryStream(CreateWorkbookBytes(false), writable: false);
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var path = Path.Combine(directory.Path, "entity-failure.xlsx");
        var options = new ExcelEntityBatchImportOptions(1,
            CreateImportOptions(directory.Path, path, new ExcelResourceLimits()));
        var callbackCount = 0;

        var exception = Assert.Throws<InvalidOperationException>(() =>
            importer.ImportEntityBatches(source, CreateLayout(), options, _ =>
            {
                callbackCount++;
                throw new InvalidOperationException("batch callback failed");
            }));

        Assert.Equal("batch callback failed", exception.Message);
        Assert.Equal(1, callbackCount);
        Assert.True(File.Exists(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证错误数量达到限制时完整结果、批次和截断标记保持一致。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void MaxErrors_ShouldBeReflectedByEntityBatchSummary(bool xls)
    {
        using var source = new MemoryStream(CreateWorkbookBytes(xls, multipleErrors: true),
            writable: false);
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var importOptions = new ExcelEntityImportOptions(
            new ExcelResourceLimits { MaxErrors = 1 }, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail);
        var batches = new List<ExcelEntityImportBatch<Root>>();

        var summary = importer.ImportEntityBatches(source, CreateLayout(),
            new ExcelEntityBatchImportOptions(1, importOptions), batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.True(summary.ErrorsTruncated);
        Assert.Equal(1, summary.ErrorCount);
        Assert.Equal(1, summary.Result.Errors.Count);
        Assert.Equal(1, summary.Result.MaxErrors);
        Assert.Equal(1, batches.Sum(batch => batch.Errors.Count));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 创建实体列表布局。
    /// </summary>
    /// <returns>包含明细区域的实体布局。</returns>
    private static ExcelEntityLayout<Root> CreateLayout() =>
        ExcelEntity.Layout<Root>(builder => builder.ListRegion("明细", "A1", root => root.Rows));

    /// <summary>
    /// 创建带失败工作簿输出的实体导入选项。
    /// </summary>
    /// <param name="directory">失败工作簿暂存目录。</param>
    /// <param name="path">失败工作簿目标路径。</param>
    /// <param name="limits">实体导入资源限制。</param>
    /// <returns>实体导入选项。</returns>
    private static ExcelEntityImportOptions CreateImportOptions(string directory, string path,
        ExcelResourceLimits limits) => new(limits, ExcelValidationFailureMode.Continue,
        ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
        new ExcelImportFailureOptions
        {
            Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
            DestinationPath = path,
            TemporaryDirectory = directory
        });

    /// <summary>
    /// 创建用于实体导入的 NPOI 工作簿字节。
    /// </summary>
    /// <param name="xls">是否生成 XLS 工作簿。</param>
    /// <param name="multipleErrors">是否写入两个错误行。</param>
    /// <returns>完整工作簿字节。</returns>
    private static byte[] CreateWorkbookBytes(bool xls, bool multipleErrors = false)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        CreateRow(sheet, 1, string.Empty, 1);
        CreateRow(sheet, 2, multipleErrors ? string.Empty : "B", 2);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 写入一行实体数据。
    /// </summary>
    /// <param name="sheet">目标工作表。</param>
    /// <param name="rowIndex">零基行号。</param>
    /// <param name="code">编码值。</param>
    /// <param name="quantity">数量值。</param>
    private static void CreateRow(ISheet sheet, int rowIndex, string code, int quantity)
    {
        var row = sheet.CreateRow(rowIndex);
        row.CreateCell(0).SetCellValue(code);
        row.CreateCell(1).SetCellValue(quantity);
    }

    private sealed class Root
    {
        /// <summary>
        /// 获取实体行集合。
        /// </summary>
        public List<Row> Rows { get; } = new();
    }

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

    /// <summary>
    /// 为单个测试创建并清理临时目录。
    /// </summary>
    private sealed class TestDirectory : IDisposable
    {
        /// <summary>
        /// 获取测试临时目录路径。
        /// </summary>
        public string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(),
            "npoi-entity-batch-failure-" + Guid.NewGuid().ToString("N"));

        /// <summary>
        /// 初始化一个 <see cref="TestDirectory" /> 类型的实例。
        /// </summary>
        public TestDirectory() => Directory.CreateDirectory(Path);

        /// <inheritdoc />
        public void Dispose() => Directory.Delete(Path, recursive: true);
    }
}
