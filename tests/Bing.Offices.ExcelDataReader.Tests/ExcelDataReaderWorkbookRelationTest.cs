using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.ExcelDataReader;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.ExcelDataReader.Tests;

/// <summary>
/// ExcelDataReader Workbook 分批关系的职责级测试。
/// </summary>
public sealed class ExcelDataReaderWorkbookRelationTest
{
    /// <summary>
    /// 关系绑定应支持子 Sheet 先于父 Sheet 交付，并在读取完成后填充导航集合。
    /// </summary>
    [Theory]
    [InlineData(ExcelFormat.Xls)]
    [InlineData(ExcelFormat.Xlsx)]
    public void ImportWorkbookBatches_Relations_ShouldBindAfterOutOfOrderSheets(ExcelFormat format)
    {
        using var source = CreateWorkbook(format, includeInvalidParent: false);
        var batches = new List<ExcelWorkbookImportBatch>();
        var request = CreateRequest();

        var summary = new ExcelDataReaderExcelImporter().ImportWorkbookBatches(source, request, batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", batches.SelectMany(batch => batch.Errors)
            .Select(error => error.Message)));
        Assert.Equal(2, summary.SucceededRows);
        Assert.Equal(new[] { "Children", "Parents" }, batches.Select(batch => batch.SheetName));
        var parent = Assert.Single(batches.Single(batch => batch.SheetName == "Parents").Items)
            as RelationParent;
        var child = Assert.Single(batches.Single(batch => batch.SheetName == "Children").Items)
            as RelationChild;
        Assert.Same(child, Assert.Single(parent.Children));
    }

    /// <summary>
    /// 父行失败时子批次仍然交付，关系缺失以最终子 Sheet 诊断批次报告。
    /// </summary>
    [Fact]
    public async Task ImportWorkbookBatchesAsync_Relations_ShouldReportParentFailureAndMissingParent()
    {
        using var source = CreateWorkbook(ExcelFormat.Xlsx, includeInvalidParent: true);
        var batches = new List<ExcelWorkbookImportBatch>();
        var request = CreateRequest();

        var summary = await new ExcelDataReaderExcelImporter().ImportWorkbookBatchesAsync(source, request,
            (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.False(summary.IsSuccess);
        Assert.Equal(1, summary.SucceededRows);
        Assert.Equal(1, summary.FailedRows);
        Assert.Contains(batches, batch => batch.SheetName == "Children" && batch.Items.Count == 1);
        var relationBatch = Assert.Single(batches.Where(batch => batch.SheetName == "Children"
            && batch.FirstRowIndex == 0 && batch.Items.Count == 0));
        Assert.Contains(relationBatch.Errors, error => error.Code == ExcelImportErrorCode.Relationship
            && error.RawValue as string == "P1");
        Assert.Contains(batches.SelectMany(batch => batch.Errors), error =>
            error.Code == ExcelImportErrorCode.Validation && error.SheetName == "Parents");
        Assert.Equal(2, summary.ErrorCount);
    }

    /// <summary>
    /// 创建父子工作表关系导入请求。
    /// </summary>
    /// <returns>配置完成的工作簿导入请求。</returns>
    private static ExcelWorkbookImportRequest<RelationWorkbook> CreateRequest()
        => ExcelImport.Workbook<RelationWorkbook>(workbook => workbook
            .Sheet<RelationChild>("Children", root => root.Children,
                sheet => sheet.BatchSize(1))
            .Sheet<RelationParent>("Parents", root => root.Parents,
                sheet => sheet.BatchSize(1).Validate(ExcelValidationFailureMode.Continue))
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children));

    /// <summary>
    /// 创建包含父子工作表的测试工作簿。
    /// </summary>
    /// <param name="format">测试工作簿格式。</param>
    /// <param name="includeInvalidParent">是否省略父项必填值。</param>
    /// <returns>测试工作簿输入流。</returns>
    private static MemoryStream CreateWorkbook(ExcelFormat format, bool includeInvalidParent)
    {
        using IWorkbook workbook = format == ExcelFormat.Xls
            ? new HSSFWorkbook()
            : new XSSFWorkbook();
        var children = workbook.CreateSheet("Children");
        var childHeader = children.CreateRow(0);
        childHeader.CreateCell(0).SetCellValue("ParentId");
        childHeader.CreateCell(1).SetCellValue("Name");
        var childRow = children.CreateRow(1);
        childRow.CreateCell(0).SetCellValue("P1");
        childRow.CreateCell(1).SetCellValue("Line");

        var parents = workbook.CreateSheet("Parents");
        var parentHeader = parents.CreateRow(0);
        parentHeader.CreateCell(0).SetCellValue("Id");
        parentHeader.CreateCell(1).SetCellValue("Name");
        var parentRow = parents.CreateRow(1);
        parentRow.CreateCell(0).SetCellValue("P1");
        if (!includeInvalidParent)
            parentRow.CreateCell(1).SetCellValue("Parent");

        using var output = new MemoryStream();
        workbook.Write(output, false);
        return new MemoryStream(output.ToArray(), writable: false);
    }

    /// <summary>
    /// 父子关系测试工作簿模型。
    /// </summary>
    private sealed class RelationWorkbook
    {
        /// <summary>
        /// 获取父项集合。
        /// </summary>
        public List<RelationParent> Parents { get; } = new List<RelationParent>();

        /// <summary>
        /// 获取子项集合。
        /// </summary>
        public List<RelationChild> Children { get; } = new List<RelationChild>();
    }

    /// <summary>
    /// 父子关系测试父项。
    /// </summary>
    private sealed class RelationParent
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string Id { get; set; }

        [ExcelRequired]
        /// <summary>
        /// 获取或设置父项名称。
        /// </summary>
        public string Name { get; set; }

        [ExcelIgnore]
        /// <summary>
        /// 获取子项导航集合。
        /// </summary>
        public List<RelationChild> Children { get; } = new List<RelationChild>();
    }

    /// <summary>
    /// 父子关系测试子项。
    /// </summary>
    private sealed class RelationChild
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string ParentId { get; set; }

        /// <summary>
        /// 获取或设置子项名称。
        /// </summary>
        public string Name { get; set; }
    }
}
