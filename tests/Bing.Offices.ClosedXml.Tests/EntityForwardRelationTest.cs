using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Entities;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// ClosedXML Forward 实体关系绑定测试。
/// </summary>
public sealed class EntityForwardRelationTest
{
    /// <summary>
    /// 验证 Forward 完成模式在同步和异步路径绑定父子集合。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldBindRelations(bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(), writable: false);
        var layout = CreateLayout();
        var options = new ExcelEntityBatchImportOptions(1, new ExcelEntityImportOptions(),
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, layout, options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, layout, options, _ => { });

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors));
        var parent = Assert.Single(summary.Result.Entity.Parents);
        var child = Assert.Single(summary.Result.Entity.Children);
        Assert.Same(child, Assert.Single(parent.Children));
        Assert.Equal("P-1", child.ParentId);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证完整结果关系错误使用列表结果中的原始源行坐标。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_Complete_ShouldReuseSourceRows(bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(unmatchedChild: true), writable: false);
        var layout = CreateCountingLayout();
        var options = new ExcelEntityBatchImportOptions(1, new ExcelEntityImportOptions(),
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, layout, options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, layout, options, _ => { });

        Assert.False(summary.IsSuccess);
        var relationError = Assert.Single(summary.Result.Errors);
        Assert.Equal(ExcelImportErrorCode.Relationship, relationError.Code);
        Assert.Equal("子项", relationError.SheetName);
        Assert.Equal(2, relationError.RowIndex);
        Assert.Equal(1, summary.Result.Entity.Parents.EnumerationCount);
        Assert.Equal(1, summary.Result.Entity.Children.EnumerationCount);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证摘要模式只验证关系键，不保留不完整的关系图。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_SummaryOnly_ShouldValidateRelationsWithoutRetainingEntities(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(), writable: false);
        var options = new ExcelEntityBatchImportOptions(1, new ExcelEntityImportOptions(),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors));
        Assert.False(summary.HasCompleteResult);
        Assert.Equal(2, summary.RowsRead);
        Assert.Equal(2, summary.SucceededRows);
        Assert.Equal(0, summary.ErrorCount);
        Assert.Empty(summary.Result.Entity.Parents);
        Assert.Empty(summary.Result.Entity.Children);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Buffered 摘要关系错误会进入低保留失败工作簿且不保留实体结果。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Buffered_SummaryOnly_RelationFailure_ShouldWriteErrorRows(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(unmatchedChild: true), writable: false);
        using var failure = new MemoryStream();
        var importOptions = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            });
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.False(summary.HasCompleteResult);
        Assert.Contains(summary.Result.Errors, error =>
            error.Code == ExcelImportErrorCode.Relationship
            && error.SheetName == "子项"
            && error.RowIndex == 2);
        Assert.Empty(summary.Result.Entity.Parents);
        Assert.Empty(summary.Result.Entity.Children);
        Assert.NotEmpty(failure.ToArray());
        failure.Position = 0;
        using var workbook = new XLWorkbook(failure);
        var childSheet = workbook.Worksheet("子项");
        Assert.Equal("ParentId", childSheet.Cell("A1").GetString());
        Assert.Equal("missing", childSheet.Cell("A2").GetString());
        Assert.NotNull(workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "_ImportErrors"));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Buffered 摘要关系键超限会保留资源错误并写出失败行。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Buffered_SummaryOnly_RelationKeyLimit_ShouldWriteErrorRows(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateRelationKeyLimitWorkbook(), writable: false);
        using var failure = new MemoryStream();
        var importOptions = new ExcelEntityImportOptions(
            new ExcelResourceLimits { MaxTrackedRelationKeys = 1 },
            ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules,
            ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            });
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.False(summary.HasCompleteResult);
        Assert.Contains(summary.Result.Errors, error =>
            error.Code == ExcelImportErrorCode.ResourceLimit);
        Assert.Empty(summary.Result.Entity.Parents);
        Assert.Empty(summary.Result.Entity.Children);
        Assert.NotEmpty(failure.ToArray());
        failure.Position = 0;
        using var workbook = new XLWorkbook(failure);
        Assert.NotNull(workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "父项"));
        Assert.NotNull(workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "_ImportErrors"));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Buffered 摘要导入预取消时不改写已有失败输出且保持输入流打开。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Buffered_SummaryOnly_PreCanceled_ShouldPreserveFailureDestination(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(), writable: false);
        using var failure = new MemoryStream();
        var original = new byte[] { 0x42, 0x49, 0x4E, 0x47 };
        failure.Write(original, 0, original.Length);
        failure.Position = 0;
        var importOptions = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            });
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        if (asynchronous)
        {
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                    (_, _) => Task.CompletedTask, cancellation.Token));
        }
        else
        {
            Assert.Throws<OperationCanceledException>(() =>
                importer.ImportEntityBatches(source, CreateLayout(), options, _ => { },
                    cancellation.Token));
        }

        Assert.Equal(original, failure.ToArray());
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Forward 模式在关系错误与唯一性错误之间共享错误预算。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    /// <param name="summaryOnly">是否使用摘要结果模式。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ImportEntityBatches_Forward_CombinedRelationAndUniqueErrors_ShouldShareMaxErrors(
        bool asynchronous, bool summaryOnly)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateCombinedBudgetWorkbook(), writable: false);
        var batches = new List<ExcelEntityImportBatch<CombinedBudgetDocument>>();
        var importOptions = new ExcelEntityImportOptions(
            new ExcelResourceLimits
            {
                MaxErrors = 2,
                MaxTrackedUniqueValues = 4,
                MaxTrackedRelationKeys = 2
            }, ExcelValidationFailureMode.Continue);
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Forward,
            summaryOnly ? ExcelEntityBatchResultMode.SummaryOnly
                : ExcelEntityBatchResultMode.Complete);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateCombinedBudgetLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateCombinedBudgetLayout(), options,
                batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.True(summary.ErrorsTruncated);
        Assert.Equal(2, summary.ErrorCount);
        Assert.Contains(summary.Result.Errors,
            error => error.Code == ExcelImportErrorCode.Relationship);
        var relationError = Assert.Single(summary.Result.Errors.Where(
            error => error.Code == ExcelImportErrorCode.Relationship));
        Assert.Equal("子项", relationError.SheetName);
        Assert.Equal(2, relationError.RowIndex);
        Assert.Contains(summary.Result.Errors,
            error => error.Code == ExcelImportErrorCode.Validation
                && error.SheetName == "唯一项");
        Assert.NotEmpty(batches);
        if (summaryOnly)
        {
            Assert.False(summary.HasCompleteResult);
            Assert.Empty(summary.Result.Entity.Parents);
            Assert.Empty(summary.Result.Entity.Children);
            Assert.Empty(summary.Result.Entity.UniqueLines);
        }
        else
        {
            Assert.True(summary.HasCompleteResult);
            Assert.Single(summary.Result.Entity.Parents);
            Assert.Single(summary.Result.Entity.Children);
            Assert.Single(summary.Result.Entity.UniqueLines);
        }
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证关系键资源耗尽时 SummaryOnly 优先停止后续区域，Complete 保留两类资源错误。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    /// <param name="summaryOnly">是否使用摘要结果模式。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ImportEntityBatches_Forward_ResourceBudgets_ShouldRespectModeOrder(
        bool asynchronous, bool summaryOnly)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateResourceBudgetWorkbook(), writable: false);
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(
                new ExcelResourceLimits
                {
                    MaxErrors = 5,
                    MaxTrackedUniqueValues = 1,
                    MaxTrackedRelationKeys = 1
                }, ExcelValidationFailureMode.Continue),
            ExcelEntityBatchParsingMode.Forward,
            summaryOnly ? ExcelEntityBatchResultMode.SummaryOnly
                : ExcelEntityBatchResultMode.Complete);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateCombinedBudgetLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateCombinedBudgetLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.True(summary.ErrorsTruncated);
        if (summaryOnly)
        {
            Assert.Equal(1, summary.ErrorCount);
            Assert.Single(summary.Result.Errors);
            Assert.Equal(ExcelImportErrorCode.ResourceLimit, summary.Result.Errors[0].Code);
            Assert.Single(summary.Regions);
            Assert.Equal(2, summary.RowsRead);
            Assert.Empty(summary.Result.Entity.Parents);
            Assert.Empty(summary.Result.Entity.Children);
            Assert.Empty(summary.Result.Entity.UniqueLines);
        }
        else
        {
            Assert.Equal(2, summary.ErrorCount);
            Assert.Equal(2, summary.Result.Errors.Count);
            Assert.All(summary.Result.Errors,
                error => Assert.Equal(ExcelImportErrorCode.ResourceLimit, error.Code));
            Assert.Contains(summary.Result.Errors, error =>
                error.SheetName == "唯一项" && error.RowIndex == 3);
            Assert.Contains(summary.Result.Errors, error =>
                error.SheetName == "父项" && error.RowIndex == 3);
            Assert.Equal(2, summary.Result.Entity.Parents.Count);
            Assert.Single(summary.Result.Entity.Children);
            Assert.Single(summary.Result.Entity.UniqueLines);
        }
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证摘要模式报告缺失父项关系错误且不保留实体集合。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_SummaryOnly_ShouldReportMissingParent(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(unmatchedChild: true), writable: false);
        var options = new ExcelEntityBatchImportOptions(1, new ExcelEntityImportOptions(),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.False(summary.HasCompleteResult);
        Assert.Equal(2, summary.RowsRead);
        Assert.Equal(2, summary.SucceededRows);
        Assert.Contains(summary.Result.Errors,
            error => error.Code == ExcelImportErrorCode.Relationship);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证摘要关系键预算耗尽后停止读取后续列表且不保留实体集合。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_SummaryOnly_ShouldStopWhenRelationKeyBudgetIsExhausted(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateRelationKeyLimitWorkbook(), writable: false);
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(new ExcelResourceLimits { MaxTrackedRelationKeys = 1 }),
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.False(summary.HasCompleteResult);
        Assert.Equal(2, summary.RowsRead);
        Assert.Equal(2, summary.SucceededRows);
        Assert.True(summary.ErrorsTruncated);
        Assert.Contains(summary.Result.Errors,
            error => error.Code == ExcelImportErrorCode.ResourceLimit);
        Assert.Empty(summary.Result.Entity.Parents);
        Assert.Empty(summary.Result.Entity.Children);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 创建父子列表工作簿。
    /// </summary>
    private static byte[] CreateWorkbook(bool unmatchedChild = false)
    {
        using var workbook = new XLWorkbook();
        var parents = workbook.Worksheets.Add("父项");
        parents.Cell(1, 1).Value = "Id";
        parents.Cell(2, 1).Value = "P-1";
        var children = workbook.Worksheets.Add("子项");
        children.Cell(1, 1).Value = "ParentId";
        children.Cell(1, 2).Value = "Name";
        children.Cell(2, 1).Value = unmatchedChild ? "missing" : "P-1";
        children.Cell(2, 2).Value = "detail";
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建触发摘要关系父键索引上限的工作簿。
    /// </summary>
    private static byte[] CreateRelationKeyLimitWorkbook()
    {
        using var workbook = new XLWorkbook();
        var parents = workbook.Worksheets.Add("父项");
        parents.Cell(1, 1).Value = "Id";
        parents.Cell(2, 1).Value = "P-1";
        parents.Cell(3, 1).Value = "P-2";
        var children = workbook.Worksheets.Add("子项");
        children.Cell(1, 1).Value = "ParentId";
        children.Cell(1, 2).Value = "Name";
        children.Cell(2, 1).Value = "P-1";
        children.Cell(2, 2).Value = "detail 1";
        children.Cell(3, 1).Value = "P-2";
        children.Cell(3, 2).Value = "detail 2";
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建同时包含关系错误和唯一性错误的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateCombinedBudgetWorkbook()
    {
        using var workbook = new XLWorkbook();
        var parents = workbook.Worksheets.Add("父项");
        parents.Cell(1, 1).Value = "Id";
        parents.Cell(2, 1).Value = "P-1";
        var children = workbook.Worksheets.Add("子项");
        children.Cell(1, 1).Value = "ParentId";
        children.Cell(1, 2).Value = "Name";
        children.Cell(2, 1).Value = "missing";
        children.Cell(2, 2).Value = "orphan";
        var unique = workbook.Worksheets.Add("唯一项");
        unique.Cell(1, 1).Value = "Code";
        unique.Cell(1, 2).Value = "Quantity";
        unique.Cell(2, 1).Value = "U-1";
        unique.Cell(2, 2).Value = 1;
        unique.Cell(3, 1).Value = "U-1";
        unique.Cell(3, 2).Value = 2;
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建同时触发关系键和唯一值资源上限的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateResourceBudgetWorkbook()
    {
        using var workbook = new XLWorkbook();
        var parents = workbook.Worksheets.Add("父项");
        parents.Cell(1, 1).Value = "Id";
        parents.Cell(2, 1).Value = "P-1";
        parents.Cell(3, 1).Value = "P-2";
        var children = workbook.Worksheets.Add("子项");
        children.Cell(1, 1).Value = "ParentId";
        children.Cell(1, 2).Value = "Name";
        children.Cell(2, 1).Value = "P-1";
        children.Cell(2, 2).Value = "child";
        var unique = workbook.Worksheets.Add("唯一项");
        unique.Cell(1, 1).Value = "Code";
        unique.Cell(1, 2).Value = "Quantity";
        unique.Cell(2, 1).Value = "U-1";
        unique.Cell(2, 2).Value = 1;
        unique.Cell(3, 1).Value = "U-2";
        unique.Cell(3, 2).Value = 2;
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建父子列表布局及关系定义。
    /// </summary>
    private static ExcelEntityLayout<RelationDocument> CreateLayout() =>
        ExcelEntity.Layout<RelationDocument>(builder => builder
            .ListRegion("父项", "A1", root => root.Parents)
            .ListRegion("子项", "A1", root => root.Children)
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));

    /// <summary>
    /// 创建使用可计数集合的父子列表布局。
    /// </summary>
    private static ExcelEntityLayout<CountingRelationDocument> CreateCountingLayout() =>
        ExcelEntity.Layout<CountingRelationDocument>(builder => builder
            .ListRegion("父项", "A1", root => root.Parents)
            .ListRegion("子项", "A1", root => root.Children)
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));

    /// <summary>
    /// 创建同时包含关系错误和唯一性错误的前向布局。
    /// </summary>
    private static ExcelEntityLayout<CombinedBudgetDocument> CreateCombinedBudgetLayout() =>
        ExcelEntity.Layout<CombinedBudgetDocument>(builder => builder
            .ListRegion("父项", "A1", root => root.Parents)
            .ListRegion("子项", "A1", root => root.Children)
            .ListRegion("唯一项", "A1", root => root.UniqueLines)
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));

    /// <summary>
    /// 保存关系测试根模型。
    /// </summary>
    private sealed class RelationDocument
    {
        /// <summary>
        /// 获取或设置父项集合。
        /// </summary>
        public List<RelationParent> Parents { get; set; } = new();

        /// <summary>
        /// 获取或设置子项集合。
        /// </summary>
        public List<RelationChild> Children { get; set; } = new();
    }

    /// <summary>
    /// 保存关系测试父项。
    /// </summary>
    private sealed class CountingRelationDocument
    {
        /// <summary>
        /// 获取父项集合。
        /// </summary>
        public CountingCollection<RelationParent> Parents { get; } = new();

        /// <summary>
        /// 获取子项集合。
        /// </summary>
        public CountingCollection<RelationChild> Children { get; } = new();
    }

    private sealed class CountingCollection<T> : ICollection<T>
    {
        /// <summary>
        /// 存储集合元素。
        /// </summary>
        private readonly List<T> _items = new();

        /// <summary>
        /// 获取已添加项数。
        /// </summary>
        public int Count => _items.Count;

        /// <summary>
        /// 获取集合枚举次数。
        /// </summary>
        public int EnumerationCount { get; private set; }

        /// <inheritdoc />
        public bool IsReadOnly => false;

        /// <inheritdoc />
        public void Add(T item) => _items.Add(item);

        /// <inheritdoc />
        public void Clear() => _items.Clear();

        /// <inheritdoc />
        public bool Contains(T item) => _items.Contains(item);

        /// <inheritdoc />
        public void CopyTo(T[] array, int arrayIndex) => _items.CopyTo(array, arrayIndex);

        /// <inheritdoc />
        public bool Remove(T item) => _items.Remove(item);

        /// <inheritdoc />
        public IEnumerator<T> GetEnumerator()
        {
            EnumerationCount++;
            return _items.GetEnumerator();
        }

        /// <inheritdoc />
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    private sealed class RelationParent
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 获取父项子集合。
        /// </summary>
        public List<RelationChild> Children { get; } = new();
    }

    /// <summary>
    /// 保存关系测试子项。
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

    /// <summary>
    /// 保存组合预算测试根模型。
    /// </summary>
    private sealed class CombinedBudgetDocument
    {
        /// <summary>
        /// 获取父项集合。
        /// </summary>
        public List<CombinedBudgetParent> Parents { get; } = new();

        /// <summary>
        /// 获取子项集合。
        /// </summary>
        public List<CombinedBudgetChild> Children { get; } = new();

        /// <summary>
        /// 获取唯一性测试明细集合。
        /// </summary>
        public List<CombinedBudgetLine> UniqueLines { get; } = new();
    }

    /// <summary>
    /// 保存组合预算测试父项。
    /// </summary>
    private sealed class CombinedBudgetParent
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 获取父项子集合。
        /// </summary>
        public List<CombinedBudgetChild> Children { get; } = new();
    }

    /// <summary>
    /// 保存组合预算测试子项。
    /// </summary>
    private sealed class CombinedBudgetChild
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

    /// <summary>
    /// 保存组合预算测试唯一性明细。
    /// </summary>
    private sealed class CombinedBudgetLine
    {
        /// <summary>
        /// 获取或设置唯一编码。
        /// </summary>
        [ExcelUnique]
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }
    }
}
