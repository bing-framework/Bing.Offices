using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
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
/// NPOI Forward 实体关系绑定测试。
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
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
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
    /// 验证完整结果关系绑定复用每个列表的源行结果，避免逐实体坐标映射。
    /// </summary>
    /// <param name="hssf">是否使用 XLS/HSSF 输入。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ImportEntityBatches_Forward_Complete_ShouldReuseSourceRows(
        bool hssf, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(
            hssf ? CreateHssfWorkbook(unmatchedChild: true) : CreateWorkbook(unmatchedChild: true),
            writable: false);
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
    /// 验证 HSSF Forward 完成模式在同步和异步路径绑定父子集合。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_Hssf_ShouldBindRelations(bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateHssfWorkbook(), writable: false);
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
    /// 验证 XLS/XLSX 前向关系绑定均遵守父键索引资源上限。
    /// </summary>
    /// <param name="hssf">是否使用 XLS/HSSF 输入。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ImportEntityBatches_Forward_RelationKeyLimit_ShouldReturnStructuredResourceError(
        bool hssf, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateRelationKeyLimitWorkbook(hssf), writable: false);
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(new ExcelResourceLimits { MaxTrackedRelationKeys = 1 }),
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.True(summary.ErrorsTruncated);
        Assert.Contains(summary.Result.Errors, error =>
            error.Code == Bing.Offices.Imports.ExcelImportErrorCode.ResourceLimit);
        Assert.Equal(2, summary.Result.Entity.Parents.Count);
        Assert.Empty(summary.Result.Entity.Parents[0].Children);
        Assert.Empty(summary.Result.Entity.Parents[1].Children);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证无关系布局的完整结果不会额外枚举实体集合建立源坐标索引。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_WithoutRelations_ShouldNotEnumerateCompleteCollection(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(), writable: false);
        var layout = ExcelEntity.Layout<NonEnumeratedDocument>(builder => builder
            .ListRegion("父项", "A1", root => root.Lines));
        var options = new ExcelEntityBatchImportOptions(1, new ExcelEntityImportOptions(),
            ExcelEntityBatchParsingMode.Forward);
        Assert.Empty(layout.Relations);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, layout, options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, layout, options, _ => { });

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors));
        Assert.NotNull(summary.Result.Entity.Lines.SingleAdded);
        Assert.Equal("P-1", summary.Result.Entity.Lines.SingleAdded.Id);
        Assert.Equal(0, summary.Result.Entity.Lines.EnumerationCount);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证摘要模式只验证关系键，不保留不完整的关系图。
    /// </summary>
    /// <param name="hssf">是否使用 XLS/HSSF 输入。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ImportEntityBatches_Forward_SummaryOnly_ShouldValidateRelationsWithoutRetainingEntities(
        bool hssf, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(hssf ? CreateHssfWorkbook() : CreateWorkbook(), writable: false);
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
    /// <param name="hssf">是否使用 XLS/HSSF 输入。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ImportEntityBatches_Buffered_SummaryOnly_RelationFailure_ShouldWriteErrorRows(
        bool hssf, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(
            hssf ? CreateHssfWorkbook(unmatchedChild: true) : CreateWorkbook(unmatchedChild: true),
            writable: false);
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
        using IWorkbook workbook = hssf ? new HSSFWorkbook(failure) : new XSSFWorkbook(failure);
        var childSheet = workbook.GetSheet("子项");
        Assert.NotNull(childSheet);
        Assert.Equal("ParentId", childSheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("missing", childSheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.NotNull(workbook.GetSheet("_ImportErrors"));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Buffered 摘要关系键超限会保留资源错误并写出失败行。
    /// </summary>
    /// <param name="hssf">是否使用 XLS/HSSF 输入。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ImportEntityBatches_Buffered_SummaryOnly_RelationKeyLimit_ShouldWriteErrorRows(
        bool hssf, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateRelationKeyLimitWorkbook(hssf), writable: false);
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
        using IWorkbook workbook = hssf ? new HSSFWorkbook(failure) : new XSSFWorkbook(failure);
        Assert.NotNull(workbook.GetSheet("父项"));
        Assert.NotNull(workbook.GetSheet("_ImportErrors"));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Buffered 摘要导入预取消时不改写已有失败输出且保持输入流打开。
    /// </summary>
    /// <param name="hssf">是否使用 XLS/HSSF 输入。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ImportEntityBatches_Buffered_SummaryOnly_PreCanceled_ShouldPreserveFailureDestination(
        bool hssf, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(
            hssf ? CreateHssfWorkbook() : CreateWorkbook(), writable: false);
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
    /// 验证摘要模式报告缺失父项关系错误且不保留实体集合。
    /// </summary>
    /// <param name="hssf">是否使用 XLS/HSSF 输入。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ImportEntityBatches_Forward_SummaryOnly_ShouldReportMissingParent(
        bool hssf, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(
            hssf ? CreateHssfWorkbook(unmatchedChild: true) : CreateWorkbook(unmatchedChild: true),
            writable: false);
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
    /// <param name="hssf">是否使用 XLS/HSSF 输入。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ImportEntityBatches_Forward_SummaryOnly_ShouldStopWhenRelationKeyBudgetIsExhausted(
        bool hssf, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateRelationKeyLimitWorkbook(hssf), writable: false);
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
    /// 验证 Forward 模式在关系错误与唯一性错误之间共享错误预算。
    /// </summary>
    /// <param name="hssf">是否使用 XLS/HSSF 输入。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    /// <param name="summaryOnly">是否使用摘要结果模式。</param>
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    public async Task ImportEntityBatches_Forward_CombinedRelationAndUniqueErrors_ShouldShareMaxErrors(
        bool hssf, bool asynchronous, bool summaryOnly)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateCombinedBudgetWorkbook(hssf), writable: false);
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
    /// <param name="hssf">是否使用 XLS/HSSF 输入。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    /// <param name="summaryOnly">是否使用摘要结果模式。</param>
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    public async Task ImportEntityBatches_Forward_ResourceBudgets_ShouldRespectModeOrder(
        bool hssf, bool asynchronous, bool summaryOnly)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateResourceBudgetWorkbook(hssf), writable: false);
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
            Assert.Equal(2, summary.Result.Entity.Parents.Count);
            Assert.Single(summary.Result.Entity.Children);
            Assert.Single(summary.Result.Entity.UniqueLines);
        }
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证非索引集合继续使用关系对象暂存路径。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldSupportNonIndexedRelationCollections()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(), writable: false);
        var layout = ExcelEntity.Layout<NonIndexedRelationDocument>(builder => builder
            .ListRegion("父项", "A1", root => root.Parents)
            .ListRegion("子项", "A1", root => root.Children)
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));
        var options = new ExcelEntityBatchImportOptions(1, new ExcelEntityImportOptions(),
            ExcelEntityBatchParsingMode.Forward);

        var summary = importer.ImportEntityBatches(source, layout, options, _ => { });

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors));
        var parent = Assert.Single(summary.Result.Entity.Parents);
        var child = Assert.Single(summary.Result.Entity.Children);
        Assert.Same(child, Assert.Single(parent.Children));
    }

    /// <summary>
    /// 创建父子列表工作簿。
    /// </summary>
    private static byte[] CreateWorkbook(bool unmatchedChild = false)
    {
        using var workbook = new XSSFWorkbook();
        var parents = workbook.CreateSheet("父项");
        parents.CreateRow(0).CreateCell(0).SetCellValue("Id");
        parents.CreateRow(1).CreateCell(0).SetCellValue("P-1");
        var children = workbook.CreateSheet("子项");
        children.CreateRow(0).CreateCell(0).SetCellValue("ParentId");
        children.GetRow(0).CreateCell(1).SetCellValue("Name");
        children.CreateRow(1).CreateCell(0).SetCellValue(unmatchedChild ? "missing" : "P-1");
        children.GetRow(1).CreateCell(1).SetCellValue("detail");
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建 HSSF 父子列表工作簿。
    /// </summary>
    private static byte[] CreateHssfWorkbook(bool unmatchedChild = false)
    {
        using var workbook = new HSSFWorkbook();
        var parents = workbook.CreateSheet("父项");
        parents.CreateRow(0).CreateCell(0).SetCellValue("Id");
        parents.CreateRow(1).CreateCell(0).SetCellValue("P-1");
        var children = workbook.CreateSheet("子项");
        children.CreateRow(0).CreateCell(0).SetCellValue("ParentId");
        children.GetRow(0).CreateCell(1).SetCellValue("Name");
        children.CreateRow(1).CreateCell(0).SetCellValue(unmatchedChild ? "missing" : "P-1");
        children.GetRow(1).CreateCell(1).SetCellValue("detail");
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建触发关系父键索引上限的 XLS 或 XLSX 工作簿。
    /// </summary>
    /// <param name="hssf">是否创建 XLS/HSSF 工作簿。</param>
    /// <returns>工作簿字节内容。</returns>
    private static byte[] CreateRelationKeyLimitWorkbook(bool hssf)
    {
        using IWorkbook workbook = hssf ? new HSSFWorkbook() : new XSSFWorkbook();
        var parents = workbook.CreateSheet("父项");
        parents.CreateRow(0).CreateCell(0).SetCellValue("Id");
        parents.CreateRow(1).CreateCell(0).SetCellValue("P-1");
        parents.CreateRow(2).CreateCell(0).SetCellValue("P-2");
        var children = workbook.CreateSheet("子项");
        children.CreateRow(0).CreateCell(0).SetCellValue("ParentId");
        children.GetRow(0).CreateCell(1).SetCellValue("Name");
        children.CreateRow(1).CreateCell(0).SetCellValue("P-1");
        children.GetRow(1).CreateCell(1).SetCellValue("detail 1");
        children.CreateRow(2).CreateCell(0).SetCellValue("P-2");
        children.GetRow(2).CreateCell(1).SetCellValue("detail 2");
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建同时包含关系错误和唯一性错误的 XLS 或 XLSX 工作簿。
    /// </summary>
    /// <param name="hssf">是否创建 XLS/HSSF 工作簿。</param>
    /// <returns>工作簿字节内容。</returns>
    private static byte[] CreateCombinedBudgetWorkbook(bool hssf)
    {
        using IWorkbook workbook = hssf ? new HSSFWorkbook() : new XSSFWorkbook();
        var parents = workbook.CreateSheet("父项");
        parents.CreateRow(0).CreateCell(0).SetCellValue("Id");
        parents.CreateRow(1).CreateCell(0).SetCellValue("P-1");
        var children = workbook.CreateSheet("子项");
        children.CreateRow(0).CreateCell(0).SetCellValue("ParentId");
        children.GetRow(0).CreateCell(1).SetCellValue("Name");
        children.CreateRow(1).CreateCell(0).SetCellValue("missing");
        children.GetRow(1).CreateCell(1).SetCellValue("orphan");
        var unique = workbook.CreateSheet("唯一项");
        unique.CreateRow(0).CreateCell(0).SetCellValue("Code");
        unique.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        unique.CreateRow(1).CreateCell(0).SetCellValue("U-1");
        unique.GetRow(1).CreateCell(1).SetCellValue(1);
        unique.CreateRow(2).CreateCell(0).SetCellValue("U-1");
        unique.GetRow(2).CreateCell(1).SetCellValue(2);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建同时触发关系键和唯一值资源上限的 XLS 或 XLSX 工作簿。
    /// </summary>
    /// <param name="hssf">是否创建 XLS/HSSF 工作簿。</param>
    /// <returns>工作簿字节内容。</returns>
    private static byte[] CreateResourceBudgetWorkbook(bool hssf)
    {
        using IWorkbook workbook = hssf ? new HSSFWorkbook() : new XSSFWorkbook();
        var parents = workbook.CreateSheet("父项");
        parents.CreateRow(0).CreateCell(0).SetCellValue("Id");
        parents.CreateRow(1).CreateCell(0).SetCellValue("P-1");
        parents.CreateRow(2).CreateCell(0).SetCellValue("P-2");
        var children = workbook.CreateSheet("子项");
        children.CreateRow(0).CreateCell(0).SetCellValue("ParentId");
        children.GetRow(0).CreateCell(1).SetCellValue("Name");
        children.CreateRow(1).CreateCell(0).SetCellValue("P-1");
        children.GetRow(1).CreateCell(1).SetCellValue("child");
        var unique = workbook.CreateSheet("唯一项");
        unique.CreateRow(0).CreateCell(0).SetCellValue("Code");
        unique.GetRow(0).CreateCell(1).SetCellValue("Quantity");
        unique.CreateRow(1).CreateCell(0).SetCellValue("U-1");
        unique.GetRow(1).CreateCell(1).SetCellValue(1);
        unique.CreateRow(2).CreateCell(0).SetCellValue("U-2");
        unique.GetRow(2).CreateCell(1).SetCellValue(2);
        using var output = new MemoryStream();
        workbook.Write(output, true);
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
    /// 保存可计数枚举次数的关系测试根模型。
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

    /// <summary>
    /// 保存集合写入及枚举次数。
    /// </summary>
    private sealed class CountingCollection<T> : ICollection<T>
    {
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

    /// <summary>
    /// 保存关系测试父项。
    /// </summary>
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

    /// <summary>
    /// 保存使用非索引集合的关系测试根模型。
    /// </summary>
    private sealed class NonIndexedRelationDocument
    {
        /// <summary>
        /// 获取父项集合。
        /// </summary>
        public ICollection<RelationParent> Parents { get; set; } = new HashSet<RelationParent>();

        /// <summary>
        /// 获取子项集合。
        /// </summary>
        public ICollection<RelationChild> Children { get; set; } = new HashSet<RelationChild>();
    }

    /// <summary>
    /// 保存无关系源坐标测试根模型。
    /// </summary>
    private sealed class NonEnumeratedDocument
    {
        /// <summary>
        /// 获取不会被额外枚举的明细集合。
        /// </summary>
        public NonEnumeratingCollection<NonEnumeratedLine> Lines { get; } = new();
    }

    /// <summary>
    /// 保存无关系源坐标测试明细。
    /// </summary>
    private sealed class NonEnumeratedLine
    {
        /// <summary>
        /// 获取或设置明细标识。
        /// </summary>
        public string Id { get; set; }
    }

    /// <summary>
    /// 实现集合写入但拒绝无关枚举的测试集合。
    /// </summary>
    private sealed class NonEnumeratingCollection<T> : ICollection<T>
    {
        /// <summary>
        /// 保存已写入的测试元素。
        /// </summary>
        private readonly List<T> _items = new();

        /// <summary>
        /// 获取已写入元素数量。
        /// </summary>
        public int Count => _items.Count;

        /// <summary>
        /// 获取最近写入的元素。
        /// </summary>
        public T SingleAdded => _items.Count == 1 ? _items[0] : default;

        /// <summary>
        /// 获取集合被枚举的次数。
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
            throw new InvalidOperationException("测试集合不允许无关枚举。");
        }

        /// <inheritdoc />
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
