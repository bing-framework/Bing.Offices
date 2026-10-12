using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.Conversions;
using Bing.Offices.Entities;
using Bing.Offices.Exports;
using Bing.Offices.Exceptions;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Extensions;
using Bing.Offices.Npoi.Exports;
using Microsoft.Extensions.DependencyInjection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// NPOI 实体布局分批导入测试。
/// </summary>
[Collection("Excel Async staging")]
public sealed class EntityLayoutBatchImportTest
{
    /// <summary>
    /// 验证 Buffered 批次入口在导入前拒绝空回调。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    /// <param name="useTemplate">是否使用模板入口。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task BufferedBatch_NullCallback_ShouldNotReadOrCommit(bool asynchronous,
        bool useTemplate)
    {
        using var source = new TrackingStream(CreateWorkbook(false, invalidQuantity: true));
        using var templateStream = new TrackingStream(CreateWorkbook(false));
        using var failureTarget = new MemoryStream();
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var importOptions = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failureTarget
            });
        var options = new ExcelEntityBatchImportOptions(1, importOptions);
        var template = new ExcelEntityTemplateOptions(templateStream, leaveOpen: true);

        ArgumentNullException exception;
        if (asynchronous && useTemplate)
        {
            exception = await Assert.ThrowsAsync<ArgumentNullException>(() =>
                importer.ImportForTemplateBatchesAsync(source, CreateLayout(), template, options,
                    (Func<ExcelEntityImportBatch<BatchDocument>, CancellationToken, Task>)null));
        }
        else if (asynchronous)
        {
            exception = await Assert.ThrowsAsync<ArgumentNullException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                    (Func<ExcelEntityImportBatch<BatchDocument>, CancellationToken, Task>)null));
        }
        else if (useTemplate)
        {
            exception = Assert.Throws<ArgumentNullException>(() =>
                importer.ImportForTemplateBatches(source, CreateLayout(), template, options,
                    (Action<ExcelEntityImportBatch<BatchDocument>>)null));
        }
        else
        {
            exception = Assert.Throws<ArgumentNullException>(() =>
                importer.ImportEntityBatches(source, CreateLayout(), options,
                    (Action<ExcelEntityImportBatch<BatchDocument>>)null));
        }

        Assert.Equal("onBatch", exception.ParamName);
        Assert.Equal(0, source.ReadCount);
        Assert.Equal(0, templateStream.ReadCount);
        Assert.Empty(failureTarget.ToArray());
    }

    /// <summary>
    /// 枚举 XLS、XLSX 与同步、异步实体批次入口组合。
    /// </summary>
    /// <returns>测试参数集合。</returns>
    public static IEnumerable<object[]> FormatsAndCalls()
    {
        foreach (var xls in new[] { false, true })
        foreach (var asynchronous in new[] { false, true })
            yield return new object[] { xls, asynchronous };
    }

    /// <summary>
    /// 验证前向实体批次对未知失败工作簿模式在计划阶段拒绝请求。
    /// </summary>
    [Fact]
    public void ForwardFailureWorkbook_UnknownMode_ShouldBeRejectedBeforeReading()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var failure = new MemoryStream();
        var importOptions = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = (ExcelImportFailureWorkbookMode)999,
                Destination = failure
            });
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Forward);

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            importer.ImportEntityBatches(new MemoryStream(CreateWorkbook(false)), CreateLayout(),
                options, _ => { }));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Empty(failure.ToArray());
    }

    /// <summary>
    /// 验证 HSSF 与 XSSF 均按列表区域和批次大小交付实体。
    /// </summary>
    /// <param name="xls">是否生成 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImportEntityBatches_ShouldProjectRegions(bool xls)
    {
        var layout = CreateLayout();
        var source = new TrackingStream(CreateWorkbook(xls));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();

        var summary = importer.ImportEntityBatches(source, layout,
            new ExcelEntityBatchImportOptions(1), batches.Add);

        Assert.True(source.CanRead);
        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(2, summary.Regions.Count);
        Assert.Equal(5, summary.RowsRead);
        Assert.Equal(5, summary.SucceededRows);
        Assert.Equal(0, summary.FailedRows);
        Assert.Equal(5, batches.Count);
        Assert.Equal(new[] { "Lines", "Lines", "Lines", "Taxes", "Taxes" },
            batches.Select(batch => batch.PropertyName));
        Assert.All(batches, batch => Assert.False(batch.IsDiagnostic));
        Assert.Equal(new[] { 2, 3, 4 }, batches.Take(3).Select(batch => batch.FirstRowIndex));
        Assert.Equal("采购单", summary.Result.Entity.Title);
        Assert.Equal(42, summary.Result.Entity.Number);
        Assert.Equal(3, summary.Result.Entity.Lines.Count);
        Assert.Equal(2, summary.Result.Entity.Taxes.Count);
    }

    /// <summary>
    /// 验证实体批次投影遇到乱序源行和错误时仍按物理行稳定交付。
    /// </summary>
    [Fact]
    public void EntityBatchProjection_ShouldFallbackForUnorderedRows()
    {
        var firstError = new ExcelImportError(ExcelImportErrorCode.ValueConversion,
            "第一行错误", "明细", 4, 1, nameof(BatchLine.Code));
        var secondError = new ExcelImportError(ExcelImportErrorCode.Validation,
            "第二行错误", "明细", 3, 1, nameof(BatchLine.Code));
        var document = new BatchDocument
        {
            Lines = new List<BatchLine>
            {
                new() { Code = "第二行" },
                new() { Code = "第一行" }
            }
        };
        var layout = CreateLayout();
        var result = new ExcelEntityImportResult<BatchDocument>(document,
            new[] { firstError, secondError },
            new[]
            {
                new ExcelSheetImportResult("明细", typeof(BatchLine), new[] { 2, 1 },
                    new[] { firstError, secondError }),
                new ExcelSheetImportResult("税额", typeof(BatchTax), Array.Empty<int>(),
                    Array.Empty<ExcelImportError>())
            });
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();

        var summary = ExcelEntityBatchProjection.Deliver(result, layout, 1, batches.Add,
            CancellationToken.None);

        Assert.False(summary.IsSuccess);
        Assert.Equal(new[] { 2, 3, 4 }, batches.Select(batch => batch.FirstRowIndex));
        Assert.Equal(new[] { 0, 1, 1 }, batches.Select(batch => batch.Errors.Count));
        Assert.Equal(new[] { "第二行错误", "第一行错误" },
            batches.SelectMany(batch => batch.Errors).Select(error => error.Message));
    }

    /// <summary>
    /// 验证实体集合预检和区域复制期间均能及时响应取消。
    /// </summary>
    /// <param name="asynchronous">是否使用异步批次入口。</param>
    /// <param name="cancelOnEnumeration">触发取消的枚举序号，零表示调用前已取消。</param>
    [Theory]
    [InlineData(false, 0)]
    [InlineData(false, 1)]
    [InlineData(false, 2)]
    [InlineData(true, 0)]
    [InlineData(true, 1)]
    [InlineData(true, 2)]
    public async Task EntityBatchProjection_CancellationDuringPreflightOrCopy_ShouldStopEnumeration(
        bool asynchronous, int cancelOnEnumeration)
    {
        using var cancellationTokenSource = new CancellationTokenSource();
        var lines = new CancelOnEnumerationEnumerable<BatchLine>(new[]
        {
            new BatchLine { Code = "A" },
            new BatchLine { Code = "B" }
        }, cancellationTokenSource, cancelOnEnumeration);
        if (cancelOnEnumeration == 0)
            cancellationTokenSource.Cancel();

        var document = new ProjectionCancellationDocument { Lines = lines };
        var layout = new ExcelEntityLayoutBuilder<ProjectionCancellationDocument>()
            .ListRegion("明细", "A1", value => value.Lines)
            .Build();
        var result = new ExcelEntityImportResult<ProjectionCancellationDocument>(document,
            Array.Empty<ExcelImportError>(), new[]
            {
                new ExcelSheetImportResult("明细", typeof(BatchLine), new[] { 1, 2 },
                    Array.Empty<ExcelImportError>())
            });
        var callbacks = 0;

        if (asynchronous)
        {
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                ExcelEntityBatchProjection.DeliverAsync(result, layout, 2,
                    (_, _) =>
                    {
                        callbacks++;
                        return Task.CompletedTask;
                    }, cancellationTokenSource.Token));
        }
        else
        {
            Assert.ThrowsAny<OperationCanceledException>(() =>
                ExcelEntityBatchProjection.Deliver(result, layout, 2, _ => callbacks++,
                    cancellationTokenSource.Token));
        }

        Assert.Equal(0, callbacks);
        Assert.Equal(cancelOnEnumeration, lines.EnumerationCount);
        Assert.Equal(cancelOnEnumeration == 0 ? 0 : 1, lines.CancelledEnumerationItemsRead);
    }

    /// <summary>
    /// 验证异步批次按顺序等待回调，并保留模板输入流所有权。
    /// </summary>
    [Fact]
    public async Task ImportForTemplateBatchesAsync_ShouldAwaitCallbacksAndLeaveTemplateOpen()
    {
        var layout = CreateLayout();
        var bytes = CreateWorkbook(false);
        using var source = new TrackingStream(bytes);
        using var templateStream = new TrackingStream(bytes);
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();
        var lastBatchNumber = 0;

        var summary = await importer.ImportForTemplateBatchesAsync(source, layout,
            new ExcelEntityTemplateOptions(templateStream, leaveOpen: true),
            new ExcelEntityBatchImportOptions(2), async (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                await Task.Yield();
                Assert.Equal(lastBatchNumber + 1, batch.BatchNumber);
                lastBatchNumber = batch.BatchNumber;
                batches.Add(batch);
            });

        Assert.True(source.CanRead);
        Assert.True(templateStream.CanRead);
        Assert.True(summary.IsSuccess);
        Assert.Equal(3, batches.Count);
        Assert.Equal(3, lastBatchNumber);
        Assert.Equal(2, batches[0].Items.Count);
        Assert.Equal(1, batches[1].Items.Count);
        Assert.Equal(2, batches[2].Items.Count);
    }

    /// <summary>
    /// 验证固定单元格转换错误以实体级诊断批次交付，同时保留列表批次。
    /// </summary>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImportEntityBatches_ShouldDeliverDiagnosticBatch(bool xls)
    {
        var layout = CreateLayout();
        var source = new TrackingStream(CreateWorkbook(xls, invalidNumber: true));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();

        var summary = importer.ImportEntityBatches(source, layout,
            new ExcelEntityBatchImportOptions(2), batches.Add);

        Assert.False(summary.IsSuccess);
        var diagnostic = Assert.Single(batches.Where(batch => batch.IsDiagnostic));
        Assert.Equal(-1, diagnostic.RegionIndex);
        Assert.Empty(diagnostic.Items);
        Assert.NotEmpty(diagnostic.Errors);
        Assert.Equal(3, batches.Count(batch => !batch.IsDiagnostic));
        Assert.NotEmpty(summary.GlobalErrors);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证区域转换失败按原始行号交付并计入失败摘要。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImportEntityBatches_ShouldDeliverRegionFailureAtSourceRow(bool xls)
    {
        var source = new TrackingStream(CreateWorkbook(xls, invalidQuantity: true));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateLayout(),
            new ExcelEntityBatchImportOptions(1), batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(5, summary.RowsRead);
        Assert.Equal(4, summary.SucceededRows);
        Assert.Equal(1, summary.FailedRows);
        var failed = Assert.Single(batches.Where(batch => batch.PropertyName == nameof(BatchDocument.Lines)
            && batch.Errors.Count > 0));
        Assert.Equal(3, failed.FirstRowIndex);
        Assert.Empty(failed.Items);
        var error = Assert.Single(failed.Errors);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, error.Code);
        Assert.Equal(3, error.RowIndex);
        Assert.Equal(nameof(BatchLine.Quantity), error.PropertyName);
    }

    /// <summary>
    /// 验证实体失败工作簿保留同一工作表的固定字段和其他列表表头。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EntityFailureWorkbook_ErrorRowsOnly_ShouldRetainEntityStructureRows(bool xls)
    {
        var layout = ExcelEntity.Layout<BatchDocument>(builder => builder
            .Cell("明细", "G1", document => document.Number)
            .ListRegion("明细", "A1", document => document.Lines,
                region => region.End("B4"))
            .ListRegion("明细", "D5", document => document.Taxes,
                region => region.End("E7")));
        using var input = new MemoryStream(CreateMultiRegionFailureWorkbook(xls), writable: false);
        using var failure = new MemoryStream();
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            });

        var result = new NpoiExcelImporter().ImportEntity(input, layout, options);

        Assert.False(result.IsSuccess);
        using var output = WorkbookFactory.Create(new MemoryStream(failure.ToArray(), writable: false));
        var sheet = output.GetSheet("明细");
        Assert.Equal("Code", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal(42, sheet.GetRow(0).GetCell(6).NumericCellValue);
        Assert.Equal("B", sheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("Name", sheet.GetRow(2).GetCell(3).StringCellValue);
        Assert.NotNull(output.GetSheet("_ImportErrors"));
    }

    /// <summary>
    /// 验证实体失败工作簿保留没有错误的单据头和其他列表工作表。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EntityFailureWorkbook_ErrorRowsOnly_ShouldRetainContextSheets(bool xls)
    {
        using var input = new MemoryStream(CreateWorkbook(xls, invalidQuantity: true), writable: false);
        using var failure = new MemoryStream();
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            });

        var result = new NpoiExcelImporter().ImportEntity(input, CreateLayout(), options);

        Assert.False(result.IsSuccess);
        using var output = WorkbookFactory.Create(new MemoryStream(failure.ToArray(), writable: false));
        Assert.Equal(42, output.GetSheet("单据").GetRow(0).GetCell(1).NumericCellValue);
        Assert.Equal("采购单", output.GetSheet("单据").GetRow(1).GetCell(1).StringCellValue);
        Assert.Equal("Code", output.GetSheet("明细").GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Name", output.GetSheet("税额").GetRow(0).GetCell(0).StringCellValue);
    }

    /// <summary>
    /// 验证实体失败工作簿保留分页小计和最终尾部行。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EntityFailureWorkbook_ErrorRowsOnly_ShouldRetainPaginationAndFooterRows(bool xls)
    {
        var layout = ExcelEntity.Layout<BatchDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .PageBreak(2)
                .PageSubtotal("PAGE_TOTAL", footer => footer.Cell("B1", "subtotal"))
                .Footer("END", footer => footer.Cell("B1", "total"))));
        using var input = new MemoryStream(CreatePaginationFailureWorkbook(xls), writable: false);
        using var failure = new MemoryStream();
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            });

        var result = new NpoiExcelImporter().ImportEntity(input, layout, options);

        Assert.False(result.IsSuccess);
        using var output = WorkbookFactory.Create(new MemoryStream(failure.ToArray(), writable: false));
        var sheet = output.GetSheet("明细");
        Assert.Equal("Code", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("B", sheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("PAGE_TOTAL", sheet.GetRow(2).GetCell(0).StringCellValue);
        Assert.Equal("subtotal", sheet.GetRow(2).GetCell(1).StringCellValue);
        Assert.Equal("END", sheet.GetRow(3).GetCell(0).StringCellValue);
        Assert.Equal("total", sheet.GetRow(3).GetCell(1).StringCellValue);
        Assert.Contains(sheet.MergedRegions, merge => merge.FormatAsString() == "C1:D1");
        var copiedHeader = sheet.GetRow(0).GetCell(2);
        Assert.True(output.GetFontAt(copiedHeader.CellStyle.FontIndex).IsBold);
    }

    /// <summary>
    /// 验证实体失败工作簿保留显式动态列的表头和错误值。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void EntityFailureWorkbook_ErrorRowsOnly_ShouldRetainDynamicColumnBoundary(bool xls)
    {
        using var input = new MemoryStream(CreateExplicitDynamicFailureWorkbook(xls), writable: false);
        using var failure = new MemoryStream();
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            });

        var result = new NpoiExcelImporter().ImportEntity(input,
            CreateExplicitDynamicForwardLayout(), options);

        Assert.False(result.IsSuccess);
        using var output = WorkbookFactory.Create(new MemoryStream(failure.ToArray(), writable: false));
        var sheet = output.GetSheet("明细");
        Assert.Equal("Code", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("Zone", sheet.GetRow(0).GetCell(1).StringCellValue);
        Assert.Equal("Score", sheet.GetRow(0).GetCell(2).StringCellValue);
        Assert.Equal("A", sheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("North", sheet.GetRow(1).GetCell(1).StringCellValue);
        Assert.Equal("wrong", sheet.GetRow(1).GetCell(2).StringCellValue);
    }

    /// <summary>
    /// 验证空列表区域和没有列表区域时均不交付空批次。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImportEntityBatches_ShouldSkipEmptyAndMissingRegions(bool xls)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var emptySource = new TrackingStream(CreateWorkbook(xls, emptyLists: true));
        var emptyBatches = new List<ExcelEntityImportBatch<BatchDocument>>();

        var emptySummary = importer.ImportEntityBatches(emptySource, CreateLayout(),
            new ExcelEntityBatchImportOptions(1), emptyBatches.Add);

        Assert.True(emptySummary.IsSuccess);
        Assert.Equal(2, emptySummary.Regions.Count);
        Assert.Equal(0, emptySummary.RowsRead);
        Assert.Empty(emptyBatches);
        Assert.Empty(emptySummary.Result.Entity.Lines);
        Assert.Empty(emptySummary.Result.Entity.Taxes);

        using var source = new TrackingStream(CreateWorkbook(xls));
        var noRegionLayout = ExcelEntity.Layout<BatchDocument>(builder => builder
            .Cell("单据", "B2", document => document.Number));
        var noRegionBatches = new List<ExcelEntityImportBatch<BatchDocument>>();

        var noRegionSummary = importer.ImportEntityBatches(source, noRegionLayout,
            new ExcelEntityBatchImportOptions(1), noRegionBatches.Add);

        Assert.True(noRegionSummary.IsSuccess);
        Assert.Empty(noRegionSummary.Regions);
        Assert.Equal(0, noRegionSummary.RowsRead);
        Assert.Empty(noRegionBatches);
    }

    /// <summary>
    /// 验证实体布局批次保留动态列值和跨区域关系绑定结果。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImportEntityBatches_ShouldPreserveDynamicColumnsAndRelations(bool xls)
    {
        var childMapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(RelationBatchChild.ParentId), Title = "ParentId" },
                new() { PropertyName = nameof(RelationBatchChild.Name), Title = "Name" }
            },
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "zone", Title = "Zone", DataTypeName = "string" }
            }
        };
        var layout = ExcelEntity.Layout<RelationBatchDocument>(builder => builder
            .ListRegion("父项", "A1", root => root.Parents)
            .ListRegion("子项", "A1", root => root.Children,
                region => region.Mapping(childMapping))
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));
        var sourceEntity = new RelationBatchDocument
        {
            Parents = new List<RelationBatchParent> { new() { Id = "P-1" } },
            Children = new List<RelationBatchChild>
            {
                new()
                {
                    ParentId = "p-1", Name = "detail",
                    CustomFields = new Dictionary<string, object> { ["zone"] = "North" }
                }
            }
        };
        using var template = CreateTemplateWorkbook(xls);
        using var exported = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(sourceEntity, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), exported);
        using var source = new TrackingStream(exported.ToArray());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<RelationBatchDocument>>();

        var summary = importer.ImportEntityBatches(source, layout,
            new ExcelEntityBatchImportOptions(1), batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(2, summary.Regions.Count);
        Assert.Equal(2, summary.RowsRead);
        var childBatch = Assert.Single(batches.Where(batch => batch.PropertyName == nameof(RelationBatchDocument.Children)));
        var child = Assert.IsType<RelationBatchChild>(Assert.Single(childBatch.Items));
        Assert.Equal("North", child.CustomFields["zone"]);
        Assert.Equal("detail", child.Name);
        var parent = Assert.Single(summary.Result.Entity.Parents);
        Assert.Same(child, Assert.Single(parent.Children));
    }

    /// <summary>
    /// 验证关系错误追加时保留已有转换错误并遵守全局错误上限。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImportEntityBatches_ShouldAppendRelationErrorsWithinGlobalLimit(bool xls)
    {
        var layout = ExcelEntity.Layout<RelationErrorDocument>(builder => builder
            .ListRegion("父项", "A1", root => root.Parents)
            .ListRegion("子项", "A1", root => root.Children, region => region.Mapping(
                new ExcelMappingConfiguration
                {
                    Columns = new List<ExcelColumnConfiguration>
                    {
                        new() { PropertyName = nameof(RelationErrorChild.ParentId), Title = "ParentId" },
                        new() { PropertyName = nameof(RelationErrorChild.Quantity), Title = "Quantity" }
                    }
                }))
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));
        using var source = new TrackingStream(CreateRelationErrorWorkbook(xls));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var importOptions = new ExcelEntityImportOptions(
            new ExcelResourceLimits { MaxErrors = 2 }, ExcelValidationFailureMode.Continue);
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Forward);

        var summary = importer.ImportEntityBatches(source, layout,
            options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.True(summary.ErrorsTruncated);
        Assert.Equal(2, summary.ErrorCount);
        Assert.Equal(2, summary.Result.Errors.Count);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, summary.Result.Errors[0].Code);
        Assert.Equal("Quantity", summary.Result.Errors[0].PropertyName);
        Assert.Equal(3, summary.Result.Errors[0].RowIndex);
        Assert.Equal(ExcelImportErrorCode.Relationship, summary.Result.Errors[1].Code);
        Assert.Equal("子项", summary.Result.Errors[1].SheetName);
        Assert.Equal(2, summary.Result.Errors[1].RowIndex);
        Assert.Equal("Key", summary.Result.Errors[1].ColumnKey);
        Assert.Equal("missing", summary.Result.Errors[1].RawValue);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证页小计和最终尾部行被跳过后仍报告明细原始物理行号。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImportEntityBatches_ShouldKeepSourceRowsAfterPageSubtotalAndFooter(bool xls)
    {
        var layout = ExcelEntity.Layout<BatchDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .PageBreak(2)
                .PageSubtotal("PAGE_TOTAL", footer => footer.Cell("B1", "subtotal"))
                .Footer("END", footer => footer.Cell("B1", "total"))));
        var entity = new BatchDocument
        {
            Lines = new List<BatchLine>
            {
                new() { Code = "A", Quantity = 2 },
                new() { Code = "B", Quantity = 3 },
                new() { Code = "C", Quantity = 4 }
            }
        };
        using var template = CreateTemplateWorkbook(xls);
        using var exported = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(entity, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), exported);
        using var source = new TrackingStream(exported.ToArray());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();

        var summary = importer.ImportEntityBatches(source, layout,
            new ExcelEntityBatchImportOptions(1), batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { 2, 3, 5 }, batches.Select(batch => batch.FirstRowIndex));
        Assert.Equal(new[] { "A", "B", "C" }, batches.Select(batch =>
            Assert.IsType<BatchLine>(Assert.Single(batch.Items)).Code));
        Assert.Equal(3, summary.RowsRead);
        Assert.Equal(3, summary.SucceededRows);
    }

    /// <summary>
    /// 验证批次回调异常会原样传播并停止后续交付。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_ShouldStopAfterCallbackException()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateWorkbook(false));
        var callbackCount = 0;

        var exception = Assert.Throws<InvalidOperationException>(() => importer.ImportEntityBatches(source,
            CreateLayout(), new ExcelEntityBatchImportOptions(1), _ =>
            {
                callbackCount++;
                throw new InvalidOperationException("callback failed");
            }));

        Assert.Equal("callback failed", exception.Message);
        Assert.Equal(1, callbackCount);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证异步批次回调取消后不再交付后续批次。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_ShouldStopAfterCallbackCancellation()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateWorkbook(false));
        using var cancellation = new CancellationTokenSource();
        var callbackCount = 0;

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => importer.ImportEntityBatchesAsync(source,
            CreateLayout(), new ExcelEntityBatchImportOptions(1), (batch, token) =>
            {
                callbackCount++;
                cancellation.Cancel();
                return Task.CompletedTask;
            }, cancellation.Token));

        Assert.Equal(1, callbackCount);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证只读集合可由 Provider 填充并继续投影为批次。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImportEntityBatches_ShouldFillGetterOnlyCollection(bool xls)
    {
        var layout = ExcelEntity.Layout<GetterOnlyBatchDocument>(builder =>
            builder.ListRegion("明细", "A1", document => document.Lines));
        var entity = new GetterOnlyBatchDocument();
        entity.Lines.Add(new BatchLine { Code = "A", Quantity = 2 });
        entity.Lines.Add(new BatchLine { Code = "B", Quantity = 3 });
        using var template = CreateTemplateWorkbook(xls);
        using var exported = new MemoryStream();
        new NpoiExcelExporter().ExportForTemplate(entity, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), exported);
        using var source = new TrackingStream(exported.ToArray());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var summary = importer.ImportEntityBatches(source, layout,
            new ExcelEntityBatchImportOptions(1), batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(2, summary.RowsRead);
        Assert.Equal(new[] { "A", "B" }, summary.Result.Entity.Lines.Select(line => line.Code));
        Assert.Equal(new[] { "A", "B" }, batches.Select(batch =>
            Assert.IsType<BatchLine>(Assert.Single(batch.Items)).Code));
    }

    /// <summary>
    /// 验证预取消不会进入实体导入或批次回调。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_ShouldHonorPreCancellation()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateWorkbook(false));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => importer.ImportEntityBatches(source,
            CreateLayout(), new ExcelEntityBatchImportOptions(), _ => { }, cancellation.Token));
    }

    /// <summary>
    /// 验证 XLSX 前向模式按读取顺序同步交付单列表批次。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldDeliverDuringRead()
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();
        var callbackCount = 0;

        var summary = importer.ImportEntityBatches(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward), batch =>
            {
                callbackCount++;
                batches.Add(batch);
            });

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(2, callbackCount);
        Assert.Equal(new[] { "A", "B" }, batches.Select(batch =>
            Assert.IsType<BatchLine>(Assert.Single(batch.Items)).Code));
        Assert.Equal(new[] { 2, 3 }, batches.Select(batch => batch.FirstRowIndex));
        Assert.Equal(new[] { 2, 3 }, summary.Result.Sheets[0].SourceRows.Select(row => row + 1));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证已交付的批次集合在后续同步或异步读取后保持稳定。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task ImportEntityBatches_Forward_ShouldKeepDeliveredBatchItemsStable(bool xls,
        bool asynchronous)
    {
        using var source = new TrackingStream(CreateForwardWorkbook(xls));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();
        IReadOnlyList<object> firstItems = null;
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);
        void CaptureBatch(ExcelEntityImportBatch<GetterOnlyBatchDocument> batch)
        {
            firstItems ??= batch.Items;
            batches.Add(batch);
        }

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateForwardLayout(), options,
                (batch, _) =>
                {
                    CaptureBatch(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateForwardLayout(), options, CaptureBatch);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A", "B" }, batches.Select(batch =>
            Assert.IsType<BatchLine>(Assert.Single(batch.Items)).Code));
        Assert.Same(batches[0].Items, firstItems);
        Assert.Equal("A", Assert.IsType<BatchLine>(Assert.Single(firstItems)).Code);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证后续错误行不会修改已交付批次的错误集合。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task ImportEntityBatches_Forward_ShouldKeepDeliveredBatchErrorsStable(bool xls,
        bool asynchronous)
    {
        using var source = new TrackingStream(CreateForwardWorkbook(xls, invalidQuantity: true,
            invalidLastQuantity: true));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();
        IReadOnlyList<ExcelImportError> firstErrors = null;
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);
        void CaptureBatch(ExcelEntityImportBatch<GetterOnlyBatchDocument> batch)
        {
            if (batch.Errors.Count > 0)
                firstErrors ??= batch.Errors;
            batches.Add(batch);
        }

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateForwardLayout(), options,
                (batch, _) =>
                {
                    CaptureBatch(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateForwardLayout(), options, CaptureBatch);

        Assert.False(summary.IsSuccess);
        Assert.Equal(2, batches.Count(batch => batch.Errors.Count > 0));
        Assert.NotNull(firstErrors);
        Assert.Single(firstErrors);
        Assert.Equal("Quantity", firstErrors[0].PropertyName);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向模式读取固定单元格并在列表批次中保留实体字段。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldReadFixedCells(bool asynchronous)
    {
        using var source = new TrackingStream(CreateFixedForwardWorkbook());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<FixedForwardDocument>>();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateFixedForwardLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateFixedForwardLayout(), options,
                batches.Add);

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal("PO-001", summary.Result.Entity.Title);
        Assert.Equal(42, summary.Result.Entity.Number);
        Assert.Equal("A", Assert.Single(summary.Result.Entity.Lines).Code);
        Assert.Single(batches);
        Assert.False(batches[0].IsDiagnostic);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证固定单元格转换错误以实体级诊断批次交付，同时继续读取列表。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReportFixedCellFailureAsDiagnostic()
    {
        using var source = new TrackingStream(CreateFixedForwardWorkbook(invalidNumber: true));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<FixedForwardDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateFixedForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.False(summary.IsSuccess);
        var diagnostic = Assert.Single(batches.Where(batch => batch.IsDiagnostic));
        Assert.Equal(-1, diagnostic.RegionIndex);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, Assert.Single(diagnostic.Errors).Code);
        Assert.Equal("Number", Assert.Single(diagnostic.Errors).PropertyName);
        Assert.Equal("A", Assert.Single(summary.Result.Entity.Lines).Code);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向模式解析工作簿级命名锚点固定单元格。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldReadNamedFixedCells(bool asynchronous)
    {
        using var source = new TrackingStream(CreateNamedFixedForwardWorkbook());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateNamedFixedForwardLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateNamedFixedForwardLayout(), options,
                batches.Add);

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(42, summary.Result.Entity.Number);
        Assert.Equal("采购单", summary.Result.Entity.Title);
        Assert.Equal(3, summary.Result.Entity.Lines.Count);
        Assert.DoesNotContain(batches, batch => batch.IsDiagnostic);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向模式在命名锚点缺失或引用区域时于读取前拒绝请求。
    /// </summary>
    [Theory]
    [InlineData(true, "缺少命名锚点")]
    [InlineData(false, "必须引用单个单元格")]
    public void ImportEntityBatches_Forward_ShouldRejectInvalidNamedFixedCells(bool missing,
        string expectedMessage)
    {
        using var source = new TrackingStream(CreateNamedFixedForwardWorkbook(
            includeNumber: !missing, areaNumber: !missing));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateNamedFixedForwardLayout(),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Contains(expectedMessage, exception.Message, StringComparison.Ordinal);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向模式按命名锚点读取偏移后的列表区域。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldReadNamedListRegion(bool asynchronous)
    {
        using var source = new TrackingStream(CreateNamedListForwardWorkbook());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateNamedListForwardLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateNamedListForwardLayout(), options,
                batches.Add);

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A-1", "B-2" },
            summary.Result.Entity.Lines.Select(line => line.Code).ToArray());
        Assert.Equal(new[] { 2, 3 },
            summary.Result.Entity.Lines.Select(line => line.Quantity).ToArray());
        Assert.Equal(2, batches.Count);
        Assert.All(batches, batch => Assert.False(batch.IsDiagnostic));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向模式在列表命名锚点缺失或引用区域时于读取前拒绝请求。
    /// </summary>
    [Theory]
    [InlineData(true, "缺少命名锚点")]
    [InlineData(false, "必须引用单个单元格")]
    public void ImportEntityBatches_Forward_ShouldRejectInvalidNamedListRegion(bool missing,
        string expectedMessage)
    {
        using var source = new TrackingStream(CreateNamedListForwardWorkbook(
            includeAnchor: !missing, areaAnchor: !missing));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateNamedListForwardLayout(),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Contains(expectedMessage, exception.Message, StringComparison.Ordinal);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向模式读取单个隐式动态字典并按定义转换值。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReadImplicitDynamicDictionary()
    {
        using var source = new TrackingStream(CreateDynamicForwardWorkbook());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<RelationBatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateDynamicForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        var child = Assert.Single(summary.Result.Entity.Children);
        Assert.Equal("P-1", child.ParentId);
        Assert.Equal("detail", child.Name);
        Assert.Equal("North", child.CustomFields["zone"]);
        Assert.Equal(12.5m, child.CustomFields["score"]);
        Assert.Single(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 NPOI XLSX 前向模式异步读取单个隐式动态字典。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldReadImplicitDynamicDictionary()
    {
        using var source = new TrackingStream(CreateDynamicForwardWorkbook());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<RelationBatchDocument>>();

        var summary = await importer.ImportEntityBatchesAsync(source, CreateDynamicForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal("North", Assert.Single(summary.Result.Entity.Children).CustomFields["zone"]);
        Assert.Equal(12.5m, summary.Result.Entity.Children[0].CustomFields["score"]);
        Assert.Single(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向模式按目标属性将多个隐式动态字典分发。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReadMultipleImplicitDynamicDictionaries()
    {
        using var source = new TrackingStream(CreateMultipleImplicitDynamicForwardWorkbook());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<MultipleImplicitDynamicDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateMultipleImplicitDynamicForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward), batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        var line = Assert.Single(summary.Result.Entity.Lines);
        Assert.Equal("North", line.Attributes["zone"]);
        Assert.Equal("Extended", line.Extensions["note"]);
        Assert.Single(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 NPOI XLSX 前向异步路径按目标属性分发多个隐式动态字典。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldReadMultipleImplicitDynamicDictionaries()
    {
        using var source = new TrackingStream(CreateMultipleImplicitDynamicForwardWorkbook());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<MultipleImplicitDynamicDocument>>();

        var summary = await importer.ImportEntityBatchesAsync(source,
            CreateMultipleImplicitDynamicForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        var line = Assert.Single(summary.Result.Entity.Lines);
        Assert.Equal("North", line.Attributes["zone"]);
        Assert.Equal("Extended", line.Extensions["note"]);
        Assert.Single(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 NPOI 完整实体布局同步路径可导出并读回多个隐式动态字典。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Npoi_EntityLayout_Buffered_ShouldRoundTripMultipleImplicitDynamicDictionaries(bool useHssf)
    {
        var layout = CreateMultipleImplicitDynamicForwardLayout();
        var source = new MultipleImplicitDynamicDocument();
        source.Lines.Add(new MultipleImplicitDynamicLine
        {
            Code = "A-1",
            Attributes = new Dictionary<string, object> { ["zone"] = "North" },
            Extensions = new Dictionary<string, object> { ["note"] = "Extended" }
        });

        using var output = new MemoryStream();
        using var template = CreateTemplateWorkbook(useHssf);
        new NpoiExcelExporter().ExportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        var bytes = output.ToArray();
        using (var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false)))
        {
            Assert.Equal(useHssf, workbook is HSSFWorkbook);
            var sheet = workbook.GetSheet("明细");
            Assert.Equal(new[] { "Code", "Zone", "Note" },
                Enumerable.Range(0, 3).Select(index => sheet.GetRow(0).GetCell(index).StringCellValue));
            Assert.Equal("North", sheet.GetRow(1).GetCell(1).StringCellValue);
            Assert.Equal("Extended", sheet.GetRow(1).GetCell(2).StringCellValue);
        }

        var result = new NpoiExcelImporter().ImportEntity(new MemoryStream(bytes, writable: false), layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var line = Assert.Single(result.Entity.Lines);
        Assert.Equal("North", line.Attributes["zone"]);
        Assert.Equal("Extended", line.Extensions["note"]);
    }

    /// <summary>
    /// 验证 NPOI 完整实体布局异步路径可导出并读回多个隐式动态字典。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Npoi_EntityLayout_BufferedAsync_ShouldRoundTripMultipleImplicitDynamicDictionaries(
        bool useHssf)
    {
        var layout = CreateMultipleImplicitDynamicForwardLayout();
        var source = new MultipleImplicitDynamicDocument();
        source.Lines.Add(new MultipleImplicitDynamicLine
        {
            Code = "A-1",
            Attributes = new Dictionary<string, object> { ["zone"] = "North" },
            Extensions = new Dictionary<string, object> { ["note"] = "Extended" }
        });

        using var output = new MemoryStream();
        using var template = CreateTemplateWorkbook(useHssf);
        await new NpoiExcelExporter().ExportForTemplateAsync(source, layout,
            new ExcelEntityTemplateOptions(template, leaveOpen: true), output);
        var bytes = output.ToArray();
        using (var workbook = WorkbookFactory.Create(new MemoryStream(bytes, writable: false)))
            Assert.Equal(useHssf, workbook is HSSFWorkbook);
        var result = await new NpoiExcelImporter().ImportEntityAsync(
            new MemoryStream(bytes, writable: false), layout);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var line = Assert.Single(result.Entity.Lines);
        Assert.Equal("North", line.Attributes["zone"]);
        Assert.Equal("Extended", line.Extensions["note"]);
    }

    /// <summary>
    /// 验证 XLSX 前向模式读取单个显式动态列组并按定义转换值。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReadExplicitDynamicGroup()
    {
        using var source = new TrackingStream(CreateExplicitDynamicForwardWorkbook());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<ExplicitDynamicDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateExplicitDynamicForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward), batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        var line = Assert.Single(summary.Result.Entity.Lines);
        Assert.Equal("A-1", line.Code);
        Assert.Equal("North", line.Fields["zoneValue"]);
        Assert.Equal(12.5m, line.Fields["scoreValue"]);
        Assert.Single(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 NPOI XLSX 前向模式异步读取单个显式动态列组。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldReadExplicitDynamicGroup()
    {
        using var source = new TrackingStream(CreateExplicitDynamicForwardWorkbook());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<ExplicitDynamicDocument>>();

        var summary = await importer.ImportEntityBatchesAsync(source, CreateExplicitDynamicForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        var line = Assert.Single(summary.Result.Entity.Lines);
        Assert.Equal("North", line.Fields["zoneValue"]);
        Assert.Equal(12.5m, line.Fields["scoreValue"]);
        Assert.Single(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 NPOI Forward 路径会为显式动态组创建空字典。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldCreateNullExplicitDynamicDictionary()
    {
        using var source = new TrackingStream(CreateExplicitDynamicForwardWorkbook());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<NullDynamicDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateNullDynamicForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward), batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        var line = Assert.Single(summary.Result.Entity.Lines);
        Assert.Equal(new[] { "scoreValue", "zoneValue" }, line.Fields.Keys.OrderBy(key => key));
        Assert.Equal("North", line.Fields["zoneValue"]);
        Assert.Equal(12.5m, line.Fields["scoreValue"]);
        Assert.Single(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 NPOI Forward 异步路径创建空字典并忽略未声明动态表头。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldCreateNullExplicitDynamicDictionaryAndIgnoreUnknown()
    {
        using var source = new TrackingStream(CreateExplicitDynamicForwardWorkbook(unknownColumn: true));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<NullDynamicDocument>>();

        var summary = await importer.ImportEntityBatchesAsync(source,
            CreateNullDynamicForwardLayout(ExcelUnknownDynamicValuePolicy.Ignore),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        var line = Assert.Single(summary.Result.Entity.Lines);
        Assert.Equal("North", line.Fields["zoneValue"]);
        Assert.Equal(12.5m, line.Fields["scoreValue"]);
        Assert.Equal(2, line.Fields.Count);
        Assert.Single(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 NPOI Forward 模式同步预检模板后读取源工作簿。
    /// </summary>
    [Fact]
    public void ImportForTemplateBatches_Forward_ShouldValidateTemplateAndReadSource()
    {
        using var source = new TrackingStream(CreateExplicitDynamicForwardWorkbook());
        using var template = new TrackingStream(CreateExplicitDynamicForwardWorkbook());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<ExplicitDynamicDocument>>();

        var summary = importer.ImportForTemplateBatches(source,
            CreateExplicitDynamicForwardLayout(),
            new ExcelEntityTemplateOptions(template, leaveOpen: true),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        var line = Assert.Single(summary.Result.Entity.Lines);
        Assert.Equal("A-1", line.Code);
        Assert.Equal("North", line.Fields["zoneValue"]);
        Assert.Single(batches);
        Assert.True(source.CanRead);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证 NPOI Forward 模式异步预检模板后读取源工作簿。
    /// </summary>
    [Fact]
    public async Task ImportForTemplateBatchesAsync_Forward_ShouldValidateTemplateAndReadSource()
    {
        using var source = new TrackingStream(CreateExplicitDynamicForwardWorkbook());
        using var template = new TrackingStream(CreateExplicitDynamicForwardWorkbook());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<ExplicitDynamicDocument>>();

        var summary = await importer.ImportForTemplateBatchesAsync(source,
            CreateExplicitDynamicForwardLayout(),
            new ExcelEntityTemplateOptions(template, leaveOpen: true),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal("North", Assert.Single(summary.Result.Entity.Lines).Fields["zoneValue"]);
        Assert.Single(batches);
        Assert.True(source.CanRead);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证 NPOI Forward 模式在模板缺少目标工作表时返回配置错误。
    /// </summary>
    [Fact]
    public void ImportForTemplateBatches_Forward_ShouldRejectTemplateWithoutTargetSheet()
    {
        using var source = new TrackingStream(CreateExplicitDynamicForwardWorkbook());
        using var template = new TrackingStream(CreateForwardTemplateWorkbook("其他"));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportForTemplateBatches(source, CreateExplicitDynamicForwardLayout(),
                new ExcelEntityTemplateOptions(template, leaveOpen: true),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward), _ => { }));

        Assert.Contains("Sheet", exception.Message);
        Assert.True(source.CanRead);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证 NPOI Forward 路径按未知动态列策略拒绝显式组中的未声明表头。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldRejectUnknownExplicitDynamicColumn()
    {
        using var source = new TrackingStream(CreateExplicitDynamicForwardWorkbook(unknownColumn: true));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();

        var exception = Assert.Throws<NpoiSheetStructureException>(() => importer.ImportEntityBatches(
            source, CreateExplicitDynamicForwardLayout(ExcelUnknownDynamicValuePolicy.Fail),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward), _ => { }));

        Assert.Contains("未知动态列", exception.Message);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 NPOI XLSX 前向模式读取多个显式动态列组并分别写回字典。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReadMultipleExplicitDynamicGroups()
    {
        using var source = new TrackingStream(CreateMultipleExplicitDynamicForwardWorkbook());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<MultipleExplicitDynamicDocument>>();

        var summary = importer.ImportEntityBatches(
            source, CreateMultipleExplicitDynamicForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        var line = Assert.Single(summary.Result.Entity.Lines);
        Assert.Equal("A-1", line.Code);
        Assert.Equal("North", line.Fields["zoneValue"]);
        Assert.Equal("Extended", line.OtherFields["otherValue"]);
        Assert.Single(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 NPOI XLSX 前向模式异步读取多个显式动态列组并分别写回字典。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldReadMultipleExplicitDynamicGroups()
    {
        using var source = new TrackingStream(CreateMultipleExplicitDynamicForwardWorkbook());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<MultipleExplicitDynamicDocument>>();

        var summary = await importer.ImportEntityBatchesAsync(source,
            CreateMultipleExplicitDynamicForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        var line = Assert.Single(summary.Result.Entity.Lines);
        Assert.Equal("North", line.Fields["zoneValue"]);
        Assert.Equal("Extended", line.OtherFields["otherValue"]);
        Assert.Single(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 NPOI Forward 按未知动态列策略拒绝未声明表头。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldRejectUnknownDynamicColumn()
    {
        using var source = new TrackingStream(CreateDynamicForwardWorkbook(unknownColumn: true));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();

        var exception = Assert.Throws<NpoiSheetStructureException>(() => importer.ImportEntityBatches(
            source, CreateDynamicForwardLayout(ExcelUnknownDynamicValuePolicy.Fail),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward), _ => { }));

        Assert.Contains("未知动态列", exception.Message);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向摘要模式仅保留批次结果和统计信息。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_SummaryOnly_ShouldAvoidRetainingEntityRows()
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward,
                ExcelEntityBatchResultMode.SummaryOnly), batches.Add);

        Assert.True(summary.IsSuccess);
        Assert.False(summary.HasCompleteResult);
        Assert.Equal(2, summary.RowsRead);
        Assert.Equal(2, summary.SucceededRows);
        Assert.Empty(summary.Result.Entity.Lines);
        Assert.Empty(summary.Result.Sheets[0].SourceRows);
        Assert.Equal(new[] { "A", "B" }, batches.Select(batch =>
            Assert.IsType<BatchLine>(Assert.Single(batch.Items)).Code));
    }

    /// <summary>
    /// 验证 NPOI Forward 摘要模式不会访问根实体的结果集合。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_SummaryOnly_ShouldSkipResultCollection(
        bool asynchronous)
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<SummaryOnlyNoCollectionDocument>>();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateSummaryOnlyNoCollectionLayout(),
                options, (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateSummaryOnlyNoCollectionLayout(), options,
                batches.Add);

        Assert.True(summary.IsSuccess);
        Assert.False(summary.HasCompleteResult);
        Assert.Equal(2, summary.SucceededRows);
        Assert.Equal(new[] { "A", "B" }, batches.Select(batch =>
            Assert.IsType<BatchLine>(Assert.Single(batch.Items)).Code));
    }

    /// <summary>
    /// 验证 NPOI Buffered 摘要模式复用有界 Forward 路径。
    /// </summary>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ImportEntityBatches_SummaryOnly_Buffered_ShouldUseForwardPath(
        bool xls, bool asynchronous)
    {
        using var source = new TrackingStream(CreateForwardWorkbook(xls));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);
        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateForwardLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateForwardLayout(), options, batches.Add);

        Assert.True(summary.IsSuccess);
        Assert.False(summary.HasCompleteResult);
        Assert.Equal(2, summary.SucceededRows);
        Assert.Equal(new[] { "A", "B" }, batches.SelectMany(batch => batch.Items)
            .Cast<BatchLine>().Select(line => line.Code));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 NPOI Buffered 摘要模式可从不可定位输入读取并清理暂存。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    /// <param name="asynchronous">是否使用异步实体批次入口。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task ImportEntityBatches_SummaryOnly_Buffered_ShouldReadNonSeekableInput(
        bool xls, bool asynchronous)
    {
        var before = ForwardTemporaryFiles();
        using var source = new NonSeekableTrackingStream(CreateForwardWorkbook(xls));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateForwardLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateForwardLayout(), options, batches.Add);

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.False(summary.HasCompleteResult);
        Assert.Equal(2, summary.RowsRead);
        Assert.Equal(2, summary.SucceededRows);
        Assert.Equal(0, summary.FailedRows);
        Assert.Empty(summary.Result.Entity.Lines);
        Assert.Equal(new[] { "A", "B" }, batches.SelectMany(batch => batch.Items)
            .Cast<BatchLine>().Select(line => line.Code));
        Assert.True(source.CanRead);
        Assert.Equal(before, ForwardTemporaryFiles());
    }

    /// <summary>
    /// 验证 NPOI Buffered 摘要模式在不可定位输入超过字节上限时停止读取。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    /// <param name="asynchronous">是否使用异步实体批次入口。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task ImportEntityBatches_SummaryOnly_Buffered_ShouldRejectNonSeekableInputAboveLimit(
        bool xls, bool asynchronous)
    {
        var bytes = CreateForwardWorkbook(xls);
        var before = ForwardTemporaryFiles();
        using var source = new NonSeekableTrackingStream(bytes);
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var importOptions = new ExcelEntityImportOptions(new ExcelResourceLimits
        {
            MaxInputBytes = bytes.Length - 1
        });
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

        if (asynchronous)
        {
            var exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateForwardLayout(), options,
                    (_, _) => Task.CompletedTask));
            Assert.Equal(BingOfficesStage.Open, exception.Stage);
        }
        else
        {
            var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
                importer.ImportEntityBatches(source, CreateForwardLayout(), options, _ => { }));
            Assert.Equal(BingOfficesStage.Open, exception.Stage);
        }

        Assert.True(source.CanRead);
        Assert.Equal(before, ForwardTemporaryFiles());
    }

    /// <summary>
    /// 验证 NPOI Buffered 摘要模式在不可定位输入暂存期间取消时清理暂存。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    /// <param name="asynchronous">是否使用异步实体批次入口。</param>
    [Theory]
    [MemberData(nameof(FormatsAndCalls))]
    public async Task ImportEntityBatches_SummaryOnly_Buffered_ShouldCleanupAfterNonSeekableBufferCancellation(
        bool xls, bool asynchronous)
    {
        var before = ForwardTemporaryFiles();
        using var cancellation = new CancellationTokenSource();
        using var source = new CancellationOnReadStream(CreateForwardWorkbook(xls), cancellation, 1);
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var callbackCount = 0;
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

        if (asynchronous)
        {
            var exception = await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateForwardLayout(), options,
                    (_, _) =>
                    {
                        callbackCount++;
                        return Task.CompletedTask;
                    }, cancellation.Token));
            Assert.Equal(cancellation.Token, exception.CancellationToken);
        }
        else
        {
            var exception = Assert.ThrowsAny<OperationCanceledException>(() =>
                importer.ImportEntityBatches(source, CreateForwardLayout(), options, _ =>
                {
                    callbackCount++;
                }, cancellation.Token));
            Assert.Equal(cancellation.Token, exception.CancellationToken);
        }

        Assert.True(cancellation.IsCancellationRequested);
        Assert.Equal(0, callbackCount);
        Assert.True(source.CanRead);
        Assert.Equal(before, ForwardTemporaryFiles());
    }

    /// <summary>
    /// 验证摘要前向解析允许仅影响导出的计算列和逻辑分页。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void ImportEntityBatches_SummaryOnly_ShouldAllowCalculatedColumnsAndLogicalPageBreak(
        bool xls)
    {
        using var source = new TrackingStream(CreateForwardWorkbook(xls));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();
        var layout = ExcelEntity.Layout<GetterOnlyBatchDocument>(builder => builder
            .PrintTitleRows("明细", "1:1")
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .PageBreak(2)
                .CalculatedColumn("Total", "合计", context => context.Item.Quantity * 2)));

        var summary = importer.ImportEntityBatches(source, layout,
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Buffered,
                ExcelEntityBatchResultMode.SummaryOnly), batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.False(summary.HasCompleteResult);
        Assert.Equal(2, summary.SucceededRows);
        Assert.Equal(new[] { "A", "B" }, batches.SelectMany(batch => batch.Items)
            .Cast<BatchLine>().Select(line => line.Code));
    }

    /// <summary>
    /// 验证 NPOI Forward 唯一值索引达到上限时返回结构化资源错误。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_UniqueResourceLimit_ShouldBeReported(
        bool asynchronous)
    {
        using var source = new TrackingStream(CreateWorkbook(false));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<UniqueForwardDocument>>();
        var importOptions = new ExcelEntityImportOptions(
            new ExcelResourceLimits { MaxTrackedUniqueValues = 1 },
            ExcelValidationFailureMode.Continue);
        var options = new ExcelEntityBatchImportOptions(2, importOptions,
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateUniqueForwardLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateUniqueForwardLayout(), options,
                batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(3, summary.RowsRead);
        Assert.Equal(1, summary.SucceededRows);
        Assert.Contains(summary.Result.Errors, error =>
            error.Code == ExcelImportErrorCode.ResourceLimit
            && error.RowIndex == 3
            && error.PropertyName == nameof(UniqueForwardLine.Code));
    }

    /// <summary>
    /// 验证完整实体导入的多个列表区域共享唯一值资源上限。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public async Task ImportEntity_MultipleRegions_ShouldShareUniqueResourceLimit(
        bool xls, bool asynchronous)
    {
        using var source = new MemoryStream(CreateMultipleSheetUniqueForwardWorkbook(xls), writable: false);
        var options = new ExcelEntityImportOptions(
            new ExcelResourceLimits { MaxTrackedUniqueValues = 1 },
            ExcelValidationFailureMode.Continue);
        var importer = new NpoiExcelImporter();

        var result = asynchronous
            ? await importer.ImportEntityAsync(source, CreateMultipleSheetUniqueForwardLayout(), options)
            : importer.ImportEntity(source, CreateMultipleSheetUniqueForwardLayout(), options);

        Assert.False(result.IsSuccess);
        Assert.Single(result.Entity.Lines);
        Assert.Empty(result.Entity.OtherLines);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit
            && error.SheetName == "税额"
            && error.PropertyName == nameof(UniqueForwardLine.Code));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证完整实体导入的多个列表区域共享数据行资源上限。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public async Task ImportEntity_MultipleRegions_ShouldShareMaxRows(
        bool xls, bool asynchronous)
    {
        using var source = new MemoryStream(CreateMultipleSheetUniqueForwardWorkbook(xls), writable: false);
        var options = new ExcelEntityImportOptions(
            new ExcelResourceLimits { MaxRows = 1 },
            ExcelValidationFailureMode.Continue);
        var importer = new NpoiExcelImporter();

        var result = asynchronous
            ? await importer.ImportEntityAsync(source, CreateMultipleSheetUniqueForwardLayout(), options)
            : importer.ImportEntity(source, CreateMultipleSheetUniqueForwardLayout(), options);

        Assert.False(result.IsSuccess);
        Assert.Single(result.Entity.Lines);
        Assert.Empty(result.Entity.OtherLines);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit
            && error.SheetName == "税额");
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证实体行数超限探测行映射回原始工作表坐标。
    /// </summary>
    /// <param name="xls">是否使用 HSSF 工作簿。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(true, false)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(false, true)]
    public async Task ImportEntity_ShouldMapMaxRowsOverflowProbeToSourceRow(
        bool xls, bool asynchronous)
    {
        using var source = new MemoryStream(CreateForwardWorkbook(xls, invalidLastQuantity: true), writable: false);
        var options = new ExcelEntityImportOptions(
            new ExcelResourceLimits { MaxRows = 1 },
            ExcelValidationFailureMode.Continue);
        var importer = new NpoiExcelImporter();

        var result = asynchronous
            ? await importer.ImportEntityAsync(source, CreateForwardLayout(), options)
            : importer.ImportEntity(source, CreateForwardLayout(), options);

        Assert.False(result.IsSuccess);
        Assert.Single(result.Entity.Lines);
        var limitError = Assert.Single(result.Errors.Where(error =>
            error.Code == ExcelImportErrorCode.ResourceLimit));
        Assert.Equal("明细", limitError.SheetName);
        Assert.Equal(3, limitError.RowIndex);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 NPOI Forward 异步摘要模式不累积完整实体列表。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_SummaryOnly_ShouldAvoidRetainingEntityRows()
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var summary = await importer.ImportEntityBatchesAsync(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(2, null, ExcelEntityBatchParsingMode.Forward,
                ExcelEntityBatchResultMode.SummaryOnly),
            (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.True(summary.IsSuccess);
        Assert.False(summary.HasCompleteResult);
        Assert.Equal(2, summary.SucceededRows);
        Assert.Empty(summary.Result.Entity.Lines);
        Assert.Empty(summary.Result.Sheets[0].SourceRows);
        Assert.Equal(2, batches.SelectMany(batch => batch.Items).Count());
    }

    /// <summary>
    /// 验证 XLSX 前向摘要模式不保留最终结果中的固定字段实体。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_SummaryOnly_ShouldReleaseFixedFields(
        bool asynchronous)
    {
        using var source = new TrackingStream(CreateWorkbook(false));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();
        var options = new ExcelEntityBatchImportOptions(2, null,
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateLayout(), options, batches.Add);

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.False(summary.HasCompleteResult);
        Assert.Null(summary.Result.Entity.Title);
        Assert.Equal(0, summary.Result.Entity.Number);
        Assert.Empty(summary.Result.Entity.Lines);
        Assert.Empty(summary.Result.Entity.Taxes);
        Assert.NotEmpty(batches);
        Assert.Equal("采购单", batches[0].Entity.Title);
        Assert.Equal(42, batches[0].Entity.Number);
    }

    /// <summary>
    /// 验证 XLSX 前向摘要模式在同步和异步失败输出中只保留错误行。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_SummaryOnly_ShouldWriteOnlyErrorRowsWorkbook(
        bool asynchronous)
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false, invalidQuantity: true));
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var importOptions = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            });
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateForwardLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateForwardLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.False(summary.HasCompleteResult);
        Assert.Equal(2, summary.RowsRead);
        Assert.Equal(1, summary.SucceededRows);
        Assert.Equal(1, summary.FailedRows);
        Assert.Empty(summary.Result.Entity.Lines);
        Assert.Empty(summary.Result.Sheets[0].SourceRows);
        Assert.NotEmpty(failure.ToArray());
        failure.Position = 0;
        using var workbook = new XSSFWorkbook(failure);
        var sheet = workbook.GetSheet("明细");
        Assert.NotNull(sheet);
        Assert.Equal(1, sheet.LastRowNum);
        Assert.Equal("Code", sheet.GetRow(0).GetCell(0).StringCellValue);
        Assert.Equal("B", sheet.GetRow(1).GetCell(0).StringCellValue);
        Assert.Equal("wrong", sheet.GetRow(1).GetCell(1).StringCellValue);
        Assert.Equal(18 * 256, sheet.GetColumnWidth(0));
        Assert.Equal(12 * 256, sheet.GetColumnWidth(1));
        Assert.True(sheet.IsColumnHidden(1));
        var errorSheet = workbook.GetSheet("_ImportErrors");
        Assert.NotNull(errorSheet);
        Assert.True(errorSheet.LastRowNum >= 1);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 NPOI 低保留失败工作簿写出阶段的资源异常保留 Provider 上下文。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_SummaryOnly_FailureWriterBudget_ShouldIdentifyProvider(
        bool asynchronous)
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false, invalidQuantity: true));
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var importOptions = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure,
                MaxEstimatedTargetObjects = 1
            });
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        if (asynchronous)
        {
            var exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateForwardLayout(), options,
                    (_, _) => Task.CompletedTask));
            Assert.Equal("NPOI", exception.Provider);
        }
        else
        {
            var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
                importer.ImportEntityBatches(source, CreateForwardLayout(), options, _ => { }));
            Assert.Equal("NPOI", exception.Provider);
        }

        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 NPOI 低保留失败输出的资源异常保留 Provider 上下文。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_SummaryOnly_FailureBudget_ShouldIdentifyProvider(
        bool asynchronous)
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false, invalidQuantity: true));
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var importOptions = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure,
                MaxCopiedCells = 1
            });
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        if (asynchronous)
        {
            var exception = await Assert.ThrowsAsync<BingOfficesResourceLimitException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateForwardLayout(), options,
                    (_, _) => Task.CompletedTask));
            Assert.Equal("NPOI", exception.Provider);
        }
        else
        {
            var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
                importer.ImportEntityBatches(source, CreateForwardLayout(), options, _ => { }));
            Assert.Equal("NPOI", exception.Provider);
        }

        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向模式异步等待每个批次回调并保留源流。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldAwaitCallbacks()
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var summary = await importer.ImportEntityBatchesAsync(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            async (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                await Task.Yield();
                batches.Add(batch);
            });

        Assert.True(summary.IsSuccess);
        Assert.Equal(new[] { "A", "B" }, batches.Select(batch =>
            Assert.IsType<BatchLine>(Assert.Single(batch.Items)).Code));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向模式在开始读取前响应预取消。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldHonorPreCancellation()
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var callbackCount = 0;

        Assert.Throws<OperationCanceledException>(() => importer.ImportEntityBatches(source,
            CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            _ => callbackCount++, cancellation.Token));

        Assert.Equal(0, callbackCount);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向模式在批次回调取消后停止后续读取。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldStopAfterCallbackCancellation()
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var cancellation = new CancellationTokenSource();
        var callbackCount = 0;

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => importer.ImportEntityBatchesAsync(source,
            CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            (batch, cancellationToken) =>
            {
                callbackCount++;
                cancellation.Cancel();
                return Task.CompletedTask;
            }, cancellation.Token));

        Assert.Equal(1, callbackCount);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向模式遵守最大数据行资源限制并保留结构化错误。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldHonorMaxRows()
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();
        var importOptions = new ExcelEntityImportOptions(new ExcelResourceLimits { MaxRows = 1 });

        var summary = importer.ImportEntityBatches(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, importOptions, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(1, summary.RowsRead);
        Assert.Equal(1, summary.SucceededRows);
        var error = Assert.Single(summary.Result.Errors);
        Assert.Equal(ExcelImportErrorCode.ResourceLimit, error.Code);
        Assert.Equal(3, error.RowIndex);
        Assert.Equal(new[] { "A" }, summary.Result.Entity.Lines.Select(line => line.Code));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向模式在表头超过列限制时返回配置错误。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldHonorMaxColumns()
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false, extraColumn: true));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var importOptions = new ExcelEntityImportOptions(
            new ExcelResourceLimits { MaxColumnsPerSheet = 2 });

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            importer.ImportEntityBatches(source, CreateForwardLayout(),
                new ExcelEntityBatchImportOptions(1, importOptions, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Contains("最大使用列", exception.Message);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向模式在读取前执行输入字节资源限制。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldHonorMaxInputBytesBeforeRead()
    {
        var bytes = CreateForwardWorkbook(false);
        using var source = new TrackingStream(bytes);
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var importOptions = new ExcelEntityImportOptions(new ExcelResourceLimits
        {
            MaxInputBytes = bytes.Length - 1
        });

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            importer.ImportEntityBatches(source, CreateForwardLayout(),
                new ExcelEntityBatchImportOptions(1, importOptions, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Contains("最大字节数", exception.Message);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向模式可从不可定位流读取并保持调用方流打开。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReadNonSeekableInput()
    {
        using var source = new NonSeekableTrackingStream(CreateForwardWorkbook(false));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(2, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(2, summary.RowsRead);
        Assert.Single(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证超过复制块大小的大批次输入可暂存并在摘要模式下完成前向导入。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_LargeNonSeekableInput_ShouldBufferWithinBudget()
    {
        var bytes = CreateLargeForwardWorkbook(6000);
        Assert.True(bytes.Length > 81920, $"测试工作簿应超过复制块大小，实际为 {bytes.Length} 字节。");
        using var source = new NonSeekableTrackingStream(bytes);
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = 0;
        var options = new ExcelEntityBatchImportOptions(257,
            new ExcelEntityImportOptions(new ExcelResourceLimits
            {
                MaxInputBytes = bytes.Length
            }), ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = importer.ImportEntityBatches(source, CreateForwardLayout(), options,
            _ => batches++);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(6000, summary.RowsRead);
        Assert.Equal(6000, summary.SucceededRows);
        Assert.Equal(0, summary.FailedRows);
        Assert.False(summary.HasCompleteResult);
        Assert.Empty(summary.Result.Entity.Lines);
        Assert.Equal((6000 + 256) / 257, batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证大批次不可定位输入超过字节上限时在缓冲阶段拒绝且保持源流打开。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_LargeNonSeekableInput_ShouldRejectBudgetOverflow()
    {
        var bytes = CreateLargeForwardWorkbook(6000);
        using var source = new NonSeekableTrackingStream(bytes);
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(257,
            new ExcelEntityImportOptions(new ExcelResourceLimits
            {
                MaxInputBytes = bytes.Length - 1
            }), ExcelEntityBatchParsingMode.Forward);

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            importer.ImportEntityBatches(source, CreateForwardLayout(), options, _ => { }));

        Assert.Equal(BingOfficesStage.Open, exception.Stage);
        Assert.Contains("最大字节数", exception.Message);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证异步大批次输入在缓冲阶段取消后不交付批次并保持源流打开。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_LargeNonSeekableInput_ShouldCleanupAfterBufferCancellation()
    {
        var bytes = CreateLargeForwardWorkbook(6000);
        using var cancellation = new CancellationTokenSource();
        using var source = new CancellationOnReadStream(bytes, cancellation, cancelAfterReads: 1);
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var callbackCount = 0;
        var options = new ExcelEntityBatchImportOptions(257, null,
            ExcelEntityBatchParsingMode.Forward);

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            importer.ImportEntityBatchesAsync(source, CreateForwardLayout(), options,
                (_, _) =>
                {
                    callbackCount++;
                    return Task.CompletedTask;
                }, cancellation.Token));

        Assert.Equal(0, callbackCount);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向模式读取自定义日期格式、1904 日期系统、数值格式索引和 Unicode 文本。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImportEntityBatches_Forward_ShouldReadStylesDatesAndUnicode(bool date1904)
    {
        var converter = new ForwardStyleCaptureConverter();
        using var source = new TrackingStream(CreateStyledForwardWorkbook(date1904));
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .AddSingleton<IExcelValueConverter>(converter)
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var layout = CreateStyledForwardLayout(converter.Name);
        var batches = new List<ExcelEntityImportBatch<StyledForwardDocument>>();

        var summary = importer.ImportEntityBatches(source, layout,
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        var line = Assert.Single(summary.Result.Entity.Lines);
        Assert.Equal("商品😀", line.Code);
        Assert.Equal(new DateTime(2024, 2, 29), line.OccurredAt);
        Assert.Equal(1234.5m, line.Amount);
        Assert.Equal("中文\n换行", line.Description);
        Assert.Single(batches);
        Assert.Equal(ExcelCellKind.Number, converter.Kind);
        Assert.True(converter.FormatIndex > 0);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向模式沿用实体转换错误和源行坐标。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReportConversionFailure()
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false, invalidQuantity: true));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward), batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(2, summary.RowsRead);
        Assert.Equal(1, summary.SucceededRows);
        Assert.Equal(1, summary.FailedRows);
        var failed = Assert.Single(batches.Where(batch => batch.Errors.Count > 0));
        var error = Assert.Single(failed.Errors);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, error.Code);
        Assert.Equal(3, error.RowIndex);
        Assert.Equal(nameof(BatchLine.Quantity), error.PropertyName);
    }

    /// <summary>
    /// 验证 XLSX 前向同步批次可生成错误行失败工作簿。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldWriteErrorRowsWorkbook()
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false, invalidQuantity: true));
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            });
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, options, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.NotEmpty(summary.Result.Errors);
        Assert.NotEmpty(failure.ToArray());
        failure.Position = 0;
        using var workbook = new XSSFWorkbook(failure);
        Assert.NotNull(workbook.GetSheet("明细"));
        var errorSheet = workbook.GetSheet("_ImportErrors");
        Assert.NotNull(errorSheet);
        Assert.True(errorSheet.LastRowNum >= 1);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向同步批次可在原始单元格写入错误批注。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldAnnotateOriginalWorkbook()
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false, invalidQuantity: true));
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = failure
            });

        var summary = importer.ImportEntityBatches(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, options, ExcelEntityBatchParsingMode.Forward),
            _ => { });

        Assert.False(summary.IsSuccess);
        Assert.NotEmpty(failure.ToArray());
        failure.Position = 0;
        using var workbook = new XSSFWorkbook(failure);
        var sheet = workbook.GetSheet("明细");
        Assert.NotNull(sheet);
        Assert.NotNull(sheet.GetRow(2).GetCell(1).CellComment);
        Assert.NotNull(workbook.GetSheet("_ImportErrors"));
        Assert.Equal("wrong", sheet.GetRow(2).GetCell(1).StringCellValue);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向异步批次可生成错误行失败工作簿。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldWriteErrorRowsWorkbook()
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false, invalidQuantity: true));
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            });
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var summary = await importer.ImportEntityBatchesAsync(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, options, ExcelEntityBatchParsingMode.Forward),
            (batch, _) =>
            {
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.False(summary.IsSuccess);
        Assert.NotEmpty(summary.Result.Errors);
        Assert.NotEmpty(failure.ToArray());
        failure.Position = 0;
        using var workbook = new XSSFWorkbook(failure);
        Assert.NotNull(workbook.GetSheet("明细"));
        var errorSheet = workbook.GetSheet("_ImportErrors");
        Assert.NotNull(errorSheet);
        Assert.True(errorSheet.LastRowNum >= 1);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX 前向异步批次可在原始单元格写入错误批注。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldAnnotateOriginalWorkbook()
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false, invalidQuantity: true));
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = failure
            });

        var summary = await importer.ImportEntityBatchesAsync(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, options, ExcelEntityBatchParsingMode.Forward),
            (_, _) => Task.CompletedTask);

        Assert.False(summary.IsSuccess);
        Assert.NotEmpty(failure.ToArray());
        failure.Position = 0;
        using var workbook = new XSSFWorkbook(failure);
        Assert.NotNull(workbook.GetSheet("明细").GetRow(2).GetCell(1).CellComment);
        Assert.NotNull(workbook.GetSheet("_ImportErrors"));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSX Forward 模板批次在同步和异步入口均生成错误行工作簿。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportForTemplateBatches_Forward_ShouldWriteErrorRowsWorkbook(bool asynchronous)
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false, invalidQuantity: true));
        using var templateStream = new TrackingStream(CreateForwardWorkbook(false));
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            });
        var batchOptions = new ExcelEntityBatchImportOptions(1, options,
            ExcelEntityBatchParsingMode.Forward);
        var template = new ExcelEntityTemplateOptions(templateStream, leaveOpen: true);

        var summary = asynchronous
            ? await importer.ImportForTemplateBatchesAsync(source, CreateForwardLayout(), template,
                batchOptions, (_, _) => Task.CompletedTask)
            : importer.ImportForTemplateBatches(source, CreateForwardLayout(), template,
                batchOptions, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.NotEmpty(summary.Result.Errors);
        Assert.NotEmpty(failure.ToArray());
        Assert.True(source.CanRead);
        Assert.True(templateStream.CanRead);
    }

    /// <summary>
    /// 验证 XLSX Forward 模板批次可在原始工作簿写入错误批注。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportForTemplateBatches_Forward_ShouldAnnotateOriginalWorkbook(bool asynchronous)
    {
        using var source = new TrackingStream(CreateForwardWorkbook(false, invalidQuantity: true));
        using var templateStream = new TrackingStream(CreateForwardWorkbook(false));
        using var failure = new MemoryStream();
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = failure
            });
        var batchOptions = new ExcelEntityBatchImportOptions(1, options,
            ExcelEntityBatchParsingMode.Forward);
        var template = new ExcelEntityTemplateOptions(templateStream, leaveOpen: true);

        var summary = asynchronous
            ? await importer.ImportForTemplateBatchesAsync(source, CreateForwardLayout(), template,
                batchOptions, (_, _) => Task.CompletedTask)
            : importer.ImportForTemplateBatches(source, CreateForwardLayout(), template,
                batchOptions, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.NotEmpty(failure.ToArray());
        failure.Position = 0;
        using var workbook = new XSSFWorkbook(failure);
        Assert.NotNull(workbook.GetSheet("明细").GetRow(2).GetCell(1).CellComment);
        Assert.NotNull(workbook.GetSheet("_ImportErrors"));
        Assert.True(source.CanRead);
        Assert.True(templateStream.CanRead);
    }

    /// <summary>
    /// 验证前向模式支持受限 HSSF，并继续拒绝复杂实体布局。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldSupportHssfAndRejectComplexLayout()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();

        using var xlsSource = new TrackingStream(CreateForwardWorkbook(true));
        var xlsBatches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();
        var xlsSummary = importer.ImportEntityBatches(xlsSource, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            xlsBatches.Add);
        Assert.True(xlsSummary.IsSuccess,
            string.Join(";", xlsSummary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A", "B" }, xlsBatches.Select(batch =>
            Assert.IsType<BatchLine>(Assert.Single(batch.Items)).Code));
        Assert.Equal(new[] { 2, 3 }, xlsSummary.Result.Sheets[0].SourceRows.Select(row => row + 1));
        Assert.True(xlsSource.CanRead);

        using var complexSource = new TrackingStream(CreateForwardWorkbook(false));
        var complexException = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(complexSource, CreateComplexForwardLayout(),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward), _ => { }));
        Assert.Contains("合并", complexException.Message);
        Assert.True(complexSource.CanRead);
    }

    /// <summary>
    /// 验证 HSSF 前向模式异步读取批次并保持源流可用。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldReadHssf()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook(true));
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var summary = await importer.ImportEntityBatchesAsync(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(2, batches.Count);
        Assert.Equal(new[] { "A", "B" }, batches.Select(batch =>
            Assert.IsType<BatchLine>(Assert.Single(batch.Items)).Code));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 路径读取与列表同表的固定单元格。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_Hssf_ShouldReadFixedCells(bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateFixedForwardHssfWorkbook());
        var batches = new List<ExcelEntityImportBatch<FixedForwardDocument>>();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateFixedForwardHssfLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateFixedForwardHssfLayout(), options,
                batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal("PO-001", summary.Result.Entity.Title);
        Assert.Equal(42, summary.Result.Entity.Number);
        Assert.Equal("A", Assert.Single(summary.Result.Entity.Lines).Code);
        Assert.DoesNotContain(batches, batch => batch.IsDiagnostic);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 路径读取位于其他工作表的固定单元格。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_Hssf_ShouldReadCrossSheetFixedCells(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateCrossSheetFixedForwardHssfWorkbook());
        var batches = new List<ExcelEntityImportBatch<FixedForwardDocument>>();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateCrossSheetFixedForwardHssfLayout(),
                options, (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateCrossSheetFixedForwardHssfLayout(), options,
                batches.Add);

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal("PO-CROSS", summary.Result.Entity.Title);
        Assert.Equal(84, summary.Result.Entity.Number);
        Assert.Equal("A", Assert.Single(summary.Result.Entity.Lines).Code);
        Assert.DoesNotContain(batches, batch => batch.IsDiagnostic);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 路径读取工作簿级命名锚点固定字段。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_Hssf_ShouldReadNamedFixedCells(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateNamedFixedHssfForwardWorkbook());
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateNamedFixedForwardLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateNamedFixedForwardLayout(), options,
                batches.Add);

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(42, summary.Result.Entity.Number);
        Assert.Equal("采购单", summary.Result.Entity.Title);
        Assert.Equal(3, summary.Result.Entity.Lines.Count);
        Assert.DoesNotContain(batches, batch => batch.IsDiagnostic);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 路径按命名锚点读取偏移后的列表区域。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    /// <param name="localName">是否使用列表工作表局部名称。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ImportEntityBatches_Forward_Hssf_ShouldReadNamedListRegion(
        bool asynchronous, bool localName)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateNamedListHssfForwardWorkbook(localName: localName));
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateNamedListForwardLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateNamedListForwardLayout(), options,
                batches.Add);

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "A-1", "B-2" },
            summary.Result.Entity.Lines.Select(line => line.Code).ToArray());
        Assert.Equal(new[] { 2, 3 },
            summary.Result.Entity.Lines.Select(line => line.Quantity).ToArray());
        Assert.Equal(2, batches.Count);
        Assert.All(batches, batch => Assert.False(batch.IsDiagnostic));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 模式预检 HSSF 模板后读取源工作簿。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    /// <param name="templateXls">是否使用 HSSF 模板。</param>
    [Theory]
    [InlineData(false, true)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(true, false)]
    public async Task ImportForTemplateBatches_Forward_Hssf_ShouldValidateTemplateAndReadSource(
        bool asynchronous, bool templateXls)
    {
        using var source = new TrackingStream(CreateForwardWorkbook(true));
        using var template = CreateTemplateWorkbook(templateXls);
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);
        var templateOptions = new ExcelEntityTemplateOptions(template, leaveOpen: true);

        var summary = asynchronous
            ? await importer.ImportForTemplateBatchesAsync(source, CreateForwardLayout(),
                templateOptions, options, (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportForTemplateBatches(source, CreateForwardLayout(), templateOptions,
                options, batches.Add);

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(2, batches.Count);
        Assert.Equal(new[] { "A", "B" }, batches.Select(batch =>
            Assert.IsType<BatchLine>(Assert.Single(batch.Items)).Code));
        Assert.True(source.CanRead);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 模板在未保留时由导入器释放，源流仍由调用方保留。
    /// </summary>
    [Fact]
    public void ImportForTemplateBatches_Forward_Hssf_ShouldDisposeTemplateWhenLeaveOpenIsFalse()
    {
        using var source = new TrackingStream(CreateForwardWorkbook(true));
        using var template = CreateTemplateWorkbook(true);
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = importer.ImportForTemplateBatches(source, CreateForwardLayout(),
            new ExcelEntityTemplateOptions(template), options, _ => { });

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.False(template.CanRead);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 模板缺少目标工作表时返回配置错误并保留调用方流。
    /// </summary>
    [Fact]
    public void ImportForTemplateBatches_Forward_Hssf_ShouldRejectTemplateWithoutTargetSheet()
    {
        using var source = new TrackingStream(CreateForwardWorkbook(true));
        using var template = new TrackingStream(CreateForwardTemplateWorkbook("其他", xls: true));
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportForTemplateBatches(source, CreateForwardLayout(),
                new ExcelEntityTemplateOptions(template, leaveOpen: true),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Contains("Sheet", exception.Message);
        Assert.True(source.CanRead);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 在固定字段命名锚点缺失或引用区域时拒绝请求。
    /// </summary>
    [Theory]
    [InlineData(true, false, "缺少命名锚点")]
    [InlineData(false, true, "必须引用单个单元格")]
    public void ImportEntityBatches_Forward_Hssf_ShouldRejectInvalidNamedFixedCells(
        bool missing, bool area, string expectedMessage)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateNamedFixedHssfForwardWorkbook(
            includeNumber: !missing, areaNumber: area));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateNamedFixedForwardLayout(),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Contains(expectedMessage, exception.Message, StringComparison.Ordinal);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 在列表命名锚点缺失或引用区域时拒绝请求。
    /// </summary>
    [Theory]
    [InlineData(true, false, "缺少命名锚点")]
    [InlineData(false, true, "必须引用单个单元格")]
    public void ImportEntityBatches_Forward_Hssf_ShouldRejectInvalidNamedListRegion(
        bool missing, bool area, string expectedMessage)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateNamedListHssfForwardWorkbook(
            includeAnchor: !missing, areaAnchor: area));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateNamedListForwardLayout(),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Contains(expectedMessage, exception.Message, StringComparison.Ordinal);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 路径读取同一工作表中的多个非重叠列表区域。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_Hssf_ShouldReadMultipleRegionsOnSameSheet(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateSameSheetForwardHssfWorkbook());
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateSameSheetForwardLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateSameSheetForwardLayout(), options,
                batches.Add);

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(2, summary.Regions.Count);
        Assert.Equal(new[] { 1, 2, 3, 4 }, batches.Select(batch => batch.BatchNumber));
        Assert.All(batches, batch => Assert.Equal("明细", batch.SheetName));
        Assert.Equal(2, summary.Result.Entity.Lines.Count);
        Assert.Equal(2, summary.Result.Entity.Taxes.Count);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 模式在多个列表区域之间共享最大数据行预算。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_Hssf_MultipleRegions_ShouldShareMaxRows(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateSameSheetForwardHssfWorkbook());
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();
        var importOptions = new ExcelEntityImportOptions(new ExcelResourceLimits { MaxRows = 1 });
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateSameSheetForwardLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateSameSheetForwardLayout(), options,
                batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(1, summary.RowsRead);
        Assert.Equal(1, summary.SucceededRows);
        Assert.Equal(1, summary.Result.Entity.Lines.Count);
        Assert.Empty(summary.Result.Entity.Taxes);
        Assert.Single(summary.Result.Errors.Where(error =>
            error.Code == ExcelImportErrorCode.ResourceLimit));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 路径读取跨工作表的多个列表区域。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_Hssf_ShouldReadMultipleSheets(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateWorkbook(true));
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateMultipleSheetForwardLayout(),
                options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateMultipleSheetForwardLayout(), options,
                batches.Add);

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(2, summary.Regions.Count);
        Assert.Equal(5, batches.Count);
        Assert.All(batches.Take(3), batch => Assert.Equal("明细", batch.SheetName));
        Assert.All(batches.Skip(3), batch => Assert.Equal("税额", batch.SheetName));
        Assert.Equal(3, summary.Result.Entity.Lines.Count);
        Assert.Equal(2, summary.Result.Entity.Taxes.Count);
        Assert.Equal(Enumerable.Range(1, 5), batches.Select(batch => batch.BatchNumber));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF Forward 固定单元格转换失败时交付实体级诊断批次。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_Hssf_ShouldReportFixedCellFailureAsDiagnostic()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateFixedForwardHssfWorkbook(invalidNumber: true));
        var batches = new List<ExcelEntityImportBatch<FixedForwardDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateFixedForwardHssfLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.False(summary.IsSuccess);
        var diagnostic = Assert.Single(batches.Where(batch => batch.IsDiagnostic));
        Assert.Equal(-1, diagnostic.RegionIndex);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, Assert.Single(diagnostic.Errors).Code);
        Assert.Equal(nameof(FixedForwardDocument.Number), Assert.Single(diagnostic.Errors).PropertyName);
        Assert.Equal("A", Assert.Single(summary.Result.Entity.Lines).Code);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF 前向模式能从不可定位流读取并保持调用方流打开。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReadNonSeekableHssf()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new NonSeekableTrackingStream(CreateForwardWorkbook(true));
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(2, batches.Count);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF 前向模式保留日期格式、数字格式和 Unicode 文本。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReadHssfStylesAndUnicode()
    {
        var converter = new ForwardStyleCaptureConverter();
        using var source = new TrackingStream(CreateStyledForwardHssfWorkbook());
        using var provider = new ServiceCollection()
            .AddBingOfficesNpoi()
            .AddSingleton<IExcelValueConverter>(converter)
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<StyledForwardDocument>>();

        var summary = importer.ImportEntityBatches(source,
            CreateStyledForwardLayout(converter.Name),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        var line = Assert.Single(summary.Result.Entity.Lines);
        Assert.Equal("商品😀", line.Code);
        Assert.Equal(new DateTime(2024, 2, 29), line.OccurredAt);
        Assert.Equal(1234.5m, line.Amount);
        Assert.Equal("中文\n换行", line.Description);
        Assert.Equal(ExcelCellKind.Number, converter.Kind);
        Assert.True(converter.FormatIndex > 0);
        Assert.Single(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF 前向模式按目标工作表名称读取非首个工作表。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldSelectHssfTargetSheetByName()
    {
        using var workbook = new HSSFWorkbook();
        var other = workbook.CreateSheet("其他");
        CreateHeader(other, "Ignored");
        other.CreateRow(1).CreateCell(0).SetCellValue("ignore");
        var target = workbook.CreateSheet("明细");
        CreateHeader(target, "Code", "Quantity");
        CreateLine(target, 1, "A", 2);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        using var source = new TrackingStream(output.ToArray());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.True(summary.IsSuccess,
            string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal("A", Assert.Single(summary.Result.Entity.Lines).Code);
        Assert.Equal("明细", Assert.Single(batches).SheetName);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF 前向模式在解析公式单元格时返回结构化不支持错误。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldRejectHssfFormula()
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        CreateHeader(sheet, "Code", "Quantity");
        var row = sheet.CreateRow(1);
        row.CreateCell(0).SetCellValue("A");
        row.CreateCell(1).SetCellFormula("1+1");
        using var output = new MemoryStream();
        workbook.Write(output, true);
        using var source = new TrackingStream(output.ToArray());
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();

        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            importer.ImportEntityBatches(source, CreateForwardLayout(),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Contains("公式", exception.Message, StringComparison.Ordinal);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 HSSF 前向模式按配置生成同格式失败工作簿。
    /// </summary>
    /// <param name="mode">失败工作簿输出模式。</param>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(ExcelImportFailureWorkbookMode.ErrorRowsOnly, false)]
    [InlineData(ExcelImportFailureWorkbookMode.ErrorRowsOnly, true)]
    [InlineData(ExcelImportFailureWorkbookMode.AnnotatedOriginal, false)]
    [InlineData(ExcelImportFailureWorkbookMode.AnnotatedOriginal, true)]
    public async Task ImportEntityBatches_Forward_HssfFailureWorkbook_ShouldPreserveXls(
        ExcelImportFailureWorkbookMode mode, bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook(true, invalidQuantity: true));
        using var destination = new MemoryStream();
        var importOptions = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = mode,
                Destination = destination
            });

        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Forward);
        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateForwardLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateForwardLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.NotEmpty(summary.Result.Errors);
        Assert.True(destination.Length > 0);
        Assert.True(source.CanRead);

        destination.Position = 0;
        using var workbook = WorkbookFactory.Create(destination);
        Assert.IsType<HSSFWorkbook>(workbook);
        var sheet = workbook.GetSheet("明细");
        Assert.NotNull(sheet);
        Assert.NotNull(workbook.GetSheet("_ImportErrors"));
        if (mode == ExcelImportFailureWorkbookMode.AnnotatedOriginal)
        {
            var error = Assert.Single(summary.Result.Errors);
            var errorCell = sheet.GetRow(error.RowIndex - 1)?.GetCell(error.ColumnIndex - 1);
            Assert.NotNull(errorCell?.CellComment);
        }
        else
        {
            Assert.Equal("Code", sheet.GetRow(0).GetCell(0).StringCellValue);
            Assert.Equal("wrong", sheet.GetRow(1).GetCell(1).StringCellValue);
        }
    }

    /// <summary>
    /// 验证 Forward 模式按不同工作表读取多个列表区域并合并批次摘要。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReadMultipleSheets()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateMultipleSheetForwardWorkbook());
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateMultipleSheetForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(2, summary.Regions.Count);
        Assert.Equal(new[] { "明细", "税额" }, batches.Select(batch => batch.SheetName));
        Assert.Equal(new[] { 1, 2 }, batches.Select(batch => batch.BatchNumber));
        Assert.All(batches, batch => Assert.Same(summary.Result.Entity, batch.Entity));
        Assert.Single(summary.Result.Entity.Lines);
        Assert.Single(summary.Result.Entity.Taxes);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Forward 模式异步读取多个工作表并保持全局批次编号。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldReadMultipleSheets()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateMultipleSheetForwardWorkbook());
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();

        var summary = await importer.ImportEntityBatchesAsync(source,
            CreateMultipleSheetForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            (batch, _) =>
            {
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(2, summary.Regions.Count);
        Assert.Equal(new[] { 1, 2 }, batches.Select(batch => batch.BatchNumber));
        Assert.All(batches, batch => Assert.Same(summary.Result.Entity, batch.Entity));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Forward 模式在多个列表区域之间共享最大数据行预算。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_MultipleRegions_ShouldShareMaxRows()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateMultipleSheetForwardWorkbook());
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();
        var importOptions = new ExcelEntityImportOptions(new ExcelResourceLimits { MaxRows = 1 });

        var summary = importer.ImportEntityBatches(source, CreateMultipleSheetForwardLayout(),
            new ExcelEntityBatchImportOptions(1, importOptions, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(1, summary.RowsRead);
        Assert.Equal(1, summary.SucceededRows);
        Assert.Equal(1, summary.Result.Entity.Lines.Count);
        Assert.Empty(summary.Result.Entity.Taxes);
        Assert.Single(summary.Result.Errors.Where(error =>
            error.Code == ExcelImportErrorCode.ResourceLimit));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Forward 异步模式在多个列表区域之间共享最大数据行预算。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_MultipleRegions_ShouldShareMaxRows()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateMultipleSheetForwardWorkbook());
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();
        var importOptions = new ExcelEntityImportOptions(new ExcelResourceLimits { MaxRows = 1 });

        var summary = await importer.ImportEntityBatchesAsync(source,
            CreateMultipleSheetForwardLayout(),
            new ExcelEntityBatchImportOptions(1, importOptions, ExcelEntityBatchParsingMode.Forward),
            (batch, _) =>
            {
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.False(summary.IsSuccess);
        Assert.Equal(1, summary.RowsRead);
        Assert.Equal(1, summary.SucceededRows);
        Assert.Equal(1, summary.Result.Entity.Lines.Count);
        Assert.Empty(summary.Result.Entity.Taxes);
        Assert.Single(summary.Result.Errors.Where(error =>
            error.Code == ExcelImportErrorCode.ResourceLimit));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Forward 模式在多个列表区域之间共享最大错误预算。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    /// <param name="xls">是否使用 HSSF/XLS 工作簿。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ImportEntityBatches_Forward_MultipleRegions_ShouldShareMaxErrors(
        bool asynchronous, bool xls)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateMultipleSheetForwardErrorWorkbook(xls));
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();
        var importOptions = new ExcelEntityImportOptions(new ExcelResourceLimits { MaxErrors = 1 });
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateMultipleSheetForwardLayout(),
                options, (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateMultipleSheetForwardLayout(), options,
                batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.True(summary.ErrorsTruncated);
        Assert.Equal(1, summary.ErrorCount);
        var error = Assert.Single(summary.Result.Errors);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, error.Code);
        Assert.Equal("明细", error.SheetName);
        Assert.Empty(summary.Result.Entity.Lines);
        Assert.Empty(summary.Result.Entity.Taxes);
        Assert.Single(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Forward 模式将固定单元格错误计入列表区域共享错误预算。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_FixedCellAndRegions_ShouldShareMaxErrors(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateWorkbook(false, invalidNumber: true));
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();
        var importOptions = new ExcelEntityImportOptions(new ExcelResourceLimits { MaxErrors = 1 });
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateLayout(), options, batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.True(summary.ErrorsTruncated);
        var error = Assert.Single(summary.Result.Errors);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, error.Code);
        Assert.Equal("单据", error.SheetName);
        Assert.Empty(summary.Result.Entity.Lines);
        Assert.Empty(summary.Result.Entity.Taxes);
        Assert.Single(batches);
        Assert.True(batches[0].IsDiagnostic);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 NPOI Forward 模式在多个列表区域之间共享唯一值预算。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    /// <param name="xls">是否使用 HSSF/XLS 工作簿。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ImportEntityBatches_Forward_MultipleRegions_ShouldShareMaxUniqueValues(
        bool asynchronous, bool xls)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateMultipleSheetUniqueForwardWorkbook(xls));
        var batches = new List<ExcelEntityImportBatch<UniqueForwardDocument>>();
        var importOptions = new ExcelEntityImportOptions(
            new ExcelResourceLimits { MaxTrackedUniqueValues = 1 },
            ExcelValidationFailureMode.Continue);
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateMultipleSheetUniqueForwardLayout(),
                options, (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateMultipleSheetUniqueForwardLayout(), options,
                batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(2, summary.RowsRead);
        Assert.Equal(1, summary.SucceededRows);
        var error = Assert.Single(summary.Result.Errors.Where(item =>
            item.Code == ExcelImportErrorCode.ResourceLimit));
        Assert.Equal("税额", error.SheetName);
        Assert.Equal(nameof(UniqueForwardLine.Code), error.PropertyName);
        Assert.Single(summary.Result.Entity.Lines);
        Assert.Empty(summary.Result.Entity.OtherLines);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Forward 模式在唯一性错误达到全局错误上限后停止后续列表区域。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    /// <param name="xls">是否使用 HSSF/XLS 工作簿。</param>
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ImportEntityBatches_Forward_MultipleRegions_ShouldShareMaxErrorsWithUniqueValidation(
        bool asynchronous, bool xls)
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateMultipleSheetUniqueErrorForwardWorkbook(xls));
        var batches = new List<ExcelEntityImportBatch<UniqueForwardDocument>>();
        var importOptions = new ExcelEntityImportOptions(
            new ExcelResourceLimits { MaxErrors = 1, MaxTrackedUniqueValues = 10 },
            ExcelValidationFailureMode.Continue);
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateMultipleSheetUniqueForwardLayout(),
                options, (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateMultipleSheetUniqueForwardLayout(), options,
                batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.True(summary.ErrorsTruncated);
        Assert.Equal(2, summary.RowsRead);
        Assert.Equal(1, summary.SucceededRows);
        var error = Assert.Single(summary.Result.Errors);
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal("明细", error.SheetName);
        Assert.Equal(nameof(UniqueForwardLine.Code), error.PropertyName);
        Assert.Single(summary.Result.Entity.Lines);
        Assert.Empty(summary.Result.Entity.OtherLines);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Forward 模式读取同一工作表中的多个非重叠列表区域。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReadMultipleRegionsOnSameSheet()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateSameSheetForwardWorkbook());
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateSameSheetForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(2, summary.Regions.Count);
        Assert.Equal(new[] { 1, 2, 3, 4 }, batches.Select(batch => batch.BatchNumber));
        Assert.All(batches, batch => Assert.Equal("明细", batch.SheetName));
        Assert.All(batches, batch => Assert.Same(summary.Result.Entity, batch.Entity));
        Assert.Equal(2, summary.Result.Entity.Lines.Count);
        Assert.Equal(2, summary.Result.Entity.Taxes.Count);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Forward 模式异步读取同一工作表中的多个非重叠列表区域。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldReadMultipleRegionsOnSameSheet()
    {
        using var provider = new ServiceCollection().AddBingOfficesNpoi().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateSameSheetForwardWorkbook());
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();

        var summary = await importer.ImportEntityBatchesAsync(source, CreateSameSheetForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            (batch, _) =>
            {
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(2, summary.Regions.Count);
        Assert.Equal(new[] { 1, 2, 3, 4 }, batches.Select(batch => batch.BatchNumber));
        Assert.Equal(2, summary.Result.Entity.Lines.Count);
        Assert.Equal(2, summary.Result.Entity.Taxes.Count);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 创建包含固定单元格和两个列表区域的实体布局。
    /// </summary>
    private static ExcelEntityLayout<BatchDocument> CreateLayout() =>
        ExcelEntity.Layout<BatchDocument>(builder => builder
            .Cell("单据", "B2", document => document.Number)
            .Cell("单据", "B3", document => document.Title)
            .ListRegion("明细", "A1", document => document.Lines)
            .ListRegion("税额", "A1", document => document.Taxes));

    /// <summary>
    /// 创建前向模式使用的单列表实体布局。
    /// </summary>
    private static ExcelEntityLayout<GetterOnlyBatchDocument> CreateForwardLayout() =>
        ExcelEntity.Layout<GetterOnlyBatchDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines));

    /// <summary>
    /// 创建摘要模式不需要访问结果集合的前向布局。
    /// </summary>
    private static ExcelEntityLayout<SummaryOnlyNoCollectionDocument>
        CreateSummaryOnlyNoCollectionLayout() =>
        ExcelEntity.Layout<SummaryOnlyNoCollectionDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines));

    /// <summary>
    /// 创建包含唯一编码列的 Forward 实体布局。
    /// </summary>
    private static ExcelEntityLayout<UniqueForwardDocument> CreateUniqueForwardLayout() =>
        ExcelEntity.Layout<UniqueForwardDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines));

    /// <summary>
    /// 创建包含两个唯一列列表区域的 Forward 实体布局。
    /// </summary>
    private static ExcelEntityLayout<UniqueForwardDocument> CreateMultipleSheetUniqueForwardLayout() =>
        ExcelEntity.Layout<UniqueForwardDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines)
            .ListRegion("税额", "A1", document => document.OtherLines));

    /// <summary>
    /// 创建包含直接固定单元格和列表区域的前向实体布局。
    /// </summary>
    private static ExcelEntityLayout<FixedForwardDocument> CreateFixedForwardLayout() =>
        ExcelEntity.Layout<FixedForwardDocument>(builder => builder
            .Cell("明细", "D1", document => document.Title)
            .Cell("明细", "D2", document => document.Number)
            .ListRegion("明细", "A4", document => document.Lines));

    /// <summary>
    /// 创建 HSSF Forward 使用的同表固定单元格布局。
    /// </summary>
    private static ExcelEntityLayout<FixedForwardDocument> CreateFixedForwardHssfLayout() =>
        CreateFixedForwardLayout();

    /// <summary>
    /// 创建 HSSF Forward 使用跨工作表固定单元格的布局。
    /// </summary>
    private static ExcelEntityLayout<FixedForwardDocument> CreateCrossSheetFixedForwardHssfLayout() =>
        ExcelEntity.Layout<FixedForwardDocument>(builder => builder
            .Cell("单据", "B1", document => document.Title)
            .Cell("单据", "B2", document => document.Number)
            .ListRegion("明细", "A1", document => document.Lines));

    /// <summary>
    /// 创建使用工作簿级命名锚点固定字段的前向实体布局。
    /// </summary>
    private static ExcelEntityLayout<BatchDocument> CreateNamedFixedForwardLayout() =>
        ExcelEntity.Layout<BatchDocument>(builder => builder
            .CellNamed("单据", "OrderNumber", document => document.Number)
            .CellNamed("单据", "OrderTitle", document => document.Title)
            .ListRegion("明细", "A1", document => document.Lines));

    /// <summary>
    /// 创建使用命名锚点起始位置的前向列表布局。
    /// </summary>
    private static ExcelEntityLayout<BatchDocument> CreateNamedListForwardLayout() =>
        ExcelEntity.Layout<BatchDocument>(builder => builder
            .ListRegionNamed("明细", "LinesStart", document => document.Lines));

    /// <summary>
    /// 创建用于保留复杂布局拒绝边界的 Forward 布局。
    /// </summary>
    private static ExcelEntityLayout<GetterOnlyBatchDocument> CreateComplexForwardLayout() =>
        ExcelEntity.Layout<GetterOnlyBatchDocument>(builder => builder
            .Merge("元数据", "A1:B1")
            .ListRegion("明细", "A1", document => document.Lines));

    /// <summary>
    /// 创建不同工作表承载多个列表区域的 Forward 布局。
    /// </summary>
    private static ExcelEntityLayout<BatchDocument> CreateMultipleSheetForwardLayout() =>
        ExcelEntity.Layout<BatchDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines)
            .ListRegion("税额", "A1", document => document.Taxes));

    /// <summary>
    /// 创建同一工作表包含多个列表区域的 Forward 边界布局。
    /// </summary>
    private static ExcelEntityLayout<BatchDocument> CreateSameSheetForwardLayout()
    {
        var taxMapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(BatchTax.Name), Title = "Code" },
                new() { PropertyName = nameof(BatchTax.Amount), Title = "Quantity" }
            }
        };
        return ExcelEntity.Layout<BatchDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines,
                region => region.End("B3"))
            .ListRegion("明细", "D1", document => document.Taxes,
                region => region.End("E3").Mapping(taxMapping)));
    }

    /// <summary>
    /// 创建包含单个隐式动态字典的前向实体布局。
    /// </summary>
    private static ExcelEntityLayout<RelationBatchDocument> CreateDynamicForwardLayout(
        ExcelUnknownDynamicValuePolicy unknownPolicy = ExcelUnknownDynamicValuePolicy.Ignore)
    {
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(RelationBatchChild.ParentId), Title = "ParentId" },
                new() { PropertyName = nameof(RelationBatchChild.Name), Title = "Name" }
            },
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new() { Key = "zone", Title = "Zone", DataTypeName = "string", ColumnIndex = 2 },
                new() { Key = "score", Title = "Score", DataTypeName = "decimal", ColumnIndex = 3 }
            }
        };
        return ExcelEntity.Layout<RelationBatchDocument>(builder => builder
            .ListRegion("子项", "A1", document => document.Children,
                region => region.Mapping(mapping).UnknownDynamicValues(unknownPolicy)));
    }

    /// <summary>
    /// 创建包含两个隐式动态字典且明确列归属的前向实体布局。
    /// </summary>
    private static ExcelEntityLayout<MultipleImplicitDynamicDocument> CreateMultipleImplicitDynamicForwardLayout()
    {
        var mapping = new ExcelMappingConfiguration
        {
            Columns = new List<ExcelColumnConfiguration>
            {
                new() { PropertyName = nameof(MultipleImplicitDynamicLine.Code), Title = "Code" }
            },
            DynamicColumns = new List<ExcelMappingDynamicColumnConfiguration>
            {
                new()
                {
                    Key = "zone", Title = "Zone", DataTypeName = "string", ColumnIndex = 1,
                    TargetPropertyName = nameof(MultipleImplicitDynamicLine.Attributes)
                },
                new()
                {
                    Key = "note", Title = "Note", DataTypeName = "string", ColumnIndex = 2,
                    TargetPropertyName = nameof(MultipleImplicitDynamicLine.Extensions)
                }
            }
        };
        return ExcelEntity.Layout<MultipleImplicitDynamicDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines,
                region => region.Mapping(mapping)));
    }

    /// <summary>
    /// 创建包含单个显式动态列组的前向实体布局。
    /// </summary>
    private static ExcelEntityLayout<ExplicitDynamicDocument> CreateExplicitDynamicForwardLayout(
        ExcelUnknownDynamicValuePolicy unknownPolicy = ExcelUnknownDynamicValuePolicy.Ignore)
    {
        var definitions = new[]
        {
            new ExcelDynamicColumnDefinition
            {
                Key = "zoneValue",
                Title = "Zone",
                DataType = typeof(string),
                PhysicalColumnIndex = 1
            },
            new ExcelDynamicColumnDefinition
            {
                Key = "scoreValue",
                Title = "Score",
                DataType = typeof(decimal),
                PhysicalColumnIndex = 2
            }
        };
        return ExcelEntity.Layout<ExplicitDynamicDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines,
                region => region.DynamicColumnGroup("attributes", item => item.Fields, definitions)
                    .UnknownDynamicValues(unknownPolicy)));
    }

    /// <summary>
    /// 创建目标字典初始为空的显式动态组布局。
    /// </summary>
    private static ExcelEntityLayout<NullDynamicDocument> CreateNullDynamicForwardLayout(
        ExcelUnknownDynamicValuePolicy unknownPolicy = ExcelUnknownDynamicValuePolicy.Ignore)
    {
        var definitions = new[]
        {
            new ExcelDynamicColumnDefinition
            {
                Key = "zoneValue",
                Title = "Zone",
                DataType = typeof(string),
                PhysicalColumnIndex = 1
            },
            new ExcelDynamicColumnDefinition
            {
                Key = "scoreValue",
                Title = "Score",
                DataType = typeof(decimal),
                PhysicalColumnIndex = 2
            }
        };
        return ExcelEntity.Layout<NullDynamicDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines,
                region => region.DynamicColumnGroup("attributes", item => item.Fields, definitions)
                    .UnknownDynamicValues(unknownPolicy)));
    }

    /// <summary>
    /// 创建包含多个显式动态列组的前向实体布局。
    /// </summary>
    private static ExcelEntityLayout<MultipleExplicitDynamicDocument> CreateMultipleExplicitDynamicForwardLayout()
    {
        var first = new[]
        {
            new ExcelDynamicColumnDefinition
            {
                Key = "zoneValue",
                Title = "Zone",
                DataType = typeof(string),
                PhysicalColumnIndex = 1
            }
        };
        var second = new[]
        {
            new ExcelDynamicColumnDefinition
            {
                Key = "otherValue",
                Title = "Other",
                DataType = typeof(string),
                PhysicalColumnIndex = 2
            }
        };
        return ExcelEntity.Layout<MultipleExplicitDynamicDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines,
                region => region.DynamicColumnGroup("attributes", item => item.Fields, first)
                    .DynamicColumnGroup("extensions", item => item.OtherFields, second)));
    }

    /// <summary>
    /// 创建带数值格式捕获映射的日期实体布局。
    /// </summary>
    private static ExcelEntityLayout<StyledForwardDocument> CreateStyledForwardLayout(string converterName)
    {
        var mapping = new ExcelMappingConfiguration();
        mapping.Columns.Add(new ExcelColumnConfiguration
        {
            PropertyName = nameof(StyledForwardLine.Amount),
            ConverterName = converterName
        });
        return ExcelEntity.Layout<StyledForwardDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines,
                region => region.Mapping(mapping)));
    }

    /// <summary>
    /// 创建测试工作簿字节。
    /// </summary>
    /// <param name="xls">是否生成 HSSF 工作簿。</param>
    /// <param name="invalidNumber">是否把固定数字单元格写成无效文本。</param>
    private static byte[] CreateWorkbook(bool xls, bool invalidNumber = false, bool emptyLists = false,
        bool invalidQuantity = false)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var header = workbook.CreateSheet("单据");
        header.CreateRow(1).CreateCell(1).SetCellValue(invalidNumber ? "错误" : 42);
        header.CreateRow(2).CreateCell(1).SetCellValue("采购单");

        var lines = workbook.CreateSheet("明细");
        CreateHeader(lines, "Code", "Quantity");
        if (!emptyLists)
        {
            CreateLine(lines, 1, "A", 2);
            CreateLine(lines, 2, "B", invalidQuantity ? "wrong" : "3");
            CreateLine(lines, 3, "C", 4);
        }

        var taxes = workbook.CreateSheet("税额");
        CreateHeader(taxes, "Name", "Amount");
        if (!emptyLists)
        {
            CreateTax(taxes, 1, "增值税", 10.5);
            CreateTax(taxes, 2, "附加税", 2.5);
        }

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建同时包含转换错误和关系错误的工作簿。
    /// </summary>
    /// <param name="xls">是否生成 HSSF 工作簿。</param>
    /// <returns>工作簿字节。</returns>
    private static byte[] CreateRelationErrorWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var parents = workbook.CreateSheet("父项");
        CreateHeader(parents, "Id");
        var children = workbook.CreateSheet("子项");
        CreateHeader(children, "ParentId", "Quantity");
        var relationError = children.CreateRow(1);
        relationError.CreateCell(0).SetCellValue("missing");
        relationError.CreateCell(1).SetCellValue(1);
        var conversionError = children.CreateRow(2);
        conversionError.CreateCell(0).SetCellValue("missing");
        conversionError.CreateCell(1).SetCellValue("wrong");
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建不同工作表承载多个列表区域的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateMultipleSheetForwardWorkbook()
    {
        using IWorkbook workbook = new XSSFWorkbook();
        var lines = workbook.CreateSheet("明细");
        CreateHeader(lines, "Code", "Quantity");
        CreateLine(lines, 1, "A-1", 2);
        var taxes = workbook.CreateSheet("税额");
        CreateHeader(taxes, "Name", "Amount");
        CreateTax(taxes, 1, "增值税", 1.25);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建两个列表区域均包含转换错误的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateMultipleSheetForwardErrorWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var lines = workbook.CreateSheet("明细");
        CreateHeader(lines, "Code", "Quantity");
        CreateLine(lines, 1, "A-1", "wrong");
        var taxes = workbook.CreateSheet("税额");
        CreateHeader(taxes, "Name", "Amount");
        CreateTax(taxes, 1, "增值税", 1.25);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建两个列表区域均包含唯一编码的 Forward 工作簿。
    /// </summary>
    private static byte[] CreateMultipleSheetUniqueForwardWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var lines = workbook.CreateSheet("明细");
        CreateHeader(lines, "Code", "Quantity");
        CreateLine(lines, 1, "A-1", 2);
        var taxes = workbook.CreateSheet("税额");
        CreateHeader(taxes, "Code", "Quantity");
        CreateLine(taxes, 1, "T-1", 1);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建第一列表包含重复编码、第二列表有效的 Forward 工作簿。
    /// </summary>
    private static byte[] CreateMultipleSheetUniqueErrorForwardWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var lines = workbook.CreateSheet("明细");
        CreateHeader(lines, "Code", "Quantity");
        CreateLine(lines, 1, "A-1", 2);
        CreateLine(lines, 2, "A-1", 3);
        var taxes = workbook.CreateSheet("税额");
        CreateHeader(taxes, "Code", "Quantity");
        CreateLine(taxes, 1, "T-1", 1);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建同一工作表承载两个非重叠列表区域的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateSameSheetForwardWorkbook()
    {
        using IWorkbook workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        header.CreateCell(3).SetCellValue("Code");
        header.CreateCell(4).SetCellValue("Quantity");
        var first = sheet.CreateRow(1);
        first.CreateCell(0).SetCellValue("A-1");
        first.CreateCell(1).SetCellValue(2);
        first.CreateCell(3).SetCellValue("T-1");
        first.CreateCell(4).SetCellValue(1.25);
        var second = sheet.CreateRow(2);
        second.CreateCell(0).SetCellValue("B-2");
        second.CreateCell(1).SetCellValue(3);
        second.CreateCell(3).SetCellValue("T-2");
        second.CreateCell(4).SetCellValue(2.5);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建同一工作表承载两个非重叠列表区域的 HSSF 工作簿。
    /// </summary>
    private static byte[] CreateSameSheetForwardHssfWorkbook()
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        header.CreateCell(3).SetCellValue("Name");
        header.CreateCell(4).SetCellValue("Amount");
        var first = sheet.CreateRow(1);
        first.CreateCell(0).SetCellValue("A-1");
        first.CreateCell(1).SetCellValue(2);
        first.CreateCell(3).SetCellValue("增值税");
        first.CreateCell(4).SetCellValue(1.25);
        var second = sheet.CreateRow(2);
        second.CreateCell(0).SetCellValue("B-2");
        second.CreateCell(1).SetCellValue(3);
        second.CreateCell(3).SetCellValue("附加税");
        second.CreateCell(4).SetCellValue(2.5);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建同一工作表包含固定字段和两个列表区域的失败导入工作簿。
    /// </summary>
    /// <param name="xls">是否生成 HSSF 工作簿。</param>
    private static byte[] CreateMultiRegionFailureWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(0);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        header.CreateCell(3).SetCellValue("预留");
        header.CreateCell(4).SetCellValue("预留金额");
        header.CreateCell(6).SetCellValue(42);
        var first = sheet.CreateRow(1);
        first.CreateCell(0).SetCellValue("A");
        first.CreateCell(1).SetCellValue(2);
        var invalid = sheet.CreateRow(2);
        invalid.CreateCell(0).SetCellValue("B");
        invalid.CreateCell(1).SetCellValue("wrong");
        var third = sheet.CreateRow(3);
        third.CreateCell(0).SetCellValue("C");
        third.CreateCell(1).SetCellValue(4);
        var taxHeader = sheet.CreateRow(4);
        taxHeader.CreateCell(3).SetCellValue("Name");
        taxHeader.CreateCell(4).SetCellValue("Amount");
        var tax = sheet.CreateRow(5);
        tax.CreateCell(3).SetCellValue("增值税");
        tax.CreateCell(4).SetCellValue(10.5);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含分页小计、最终尾部和一条无效明细的工作簿。
    /// </summary>
    /// <param name="xls">是否生成 HSSF 工作簿。</param>
    private static byte[] CreatePaginationFailureWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        CreateHeader(sheet, "Code", "Quantity");
        CreateLine(sheet, 1, "A", 2);
        CreateLine(sheet, 2, "B", "wrong");
        var subtotal = sheet.CreateRow(3);
        subtotal.CreateCell(0).SetCellValue("PAGE_TOTAL");
        subtotal.CreateCell(1).SetCellValue("subtotal");
        CreateLine(sheet, 4, "C", 4);
        var footer = sheet.CreateRow(5);
        footer.CreateCell(0).SetCellValue("END");
        footer.CreateCell(1).SetCellValue("total");
        var templateHeader = sheet.GetRow(0).CreateCell(2);
        templateHeader.SetCellValue("模板表头");
        var headerStyle = workbook.CreateCellStyle();
        var headerFont = workbook.CreateFont();
        headerFont.IsBold = true;
        headerStyle.SetFont(headerFont);
        templateHeader.CellStyle = headerStyle;
        sheet.AddMergedRegion(new CellRangeAddress(0, 0, 2, 3));
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含显式动态列转换错误的工作簿。
    /// </summary>
    /// <param name="xls">是否生成 HSSF 工作簿。</param>
    private static byte[] CreateExplicitDynamicFailureWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        CreateHeader(sheet, "Code", "Zone", "Score");
        var row = sheet.CreateRow(1);
        row.CreateCell(0).SetCellValue("A");
        row.CreateCell(1).SetCellValue("North");
        row.CreateCell(2).SetCellValue("wrong");
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建前向模式使用的单 Sheet 工作簿。
    /// </summary>
    /// <param name="xls">是否生成 HSSF 工作簿。</param>
    private static byte[] CreateForwardWorkbook(bool xls, bool invalidQuantity = false,
        bool extraColumn = false, bool invalidLastQuantity = false)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        CreateHeader(sheet, extraColumn
            ? new[] { "Code", "Quantity", "Overflow" }
            : new[] { "Code", "Quantity" });
        sheet.SetColumnWidth(0, 18 * 256);
        sheet.SetColumnWidth(1, 12 * 256);
        sheet.SetColumnHidden(1, true);
        CreateLine(sheet, 1, "A", 2);
        if (invalidQuantity)
            CreateLine(sheet, 2, "B", "wrong");
        else
            CreateLine(sheet, 2, "B", 3);
        if (invalidLastQuantity)
            CreateLine(sheet, 3, "C", "wrong");
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含固定单元格和列表区域的 XLSX 前向工作簿。
    /// </summary>
    /// <param name="invalidNumber">是否将固定数字写为不可转换文本。</param>
    private static byte[] CreateFixedForwardWorkbook(bool invalidNumber = false)
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(3).SetCellValue("PO-001");
        sheet.CreateRow(1).CreateCell(3).SetCellValue(invalidNumber ? "wrong" : "42");
        var header = sheet.CreateRow(3);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        CreateLine(sheet, 4, "A", 2);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含工作簿级命名锚点的 XLSX 前向工作簿。
    /// </summary>
    /// <param name="includeNumber">是否创建数字字段命名锚点。</param>
    /// <param name="areaNumber">数字字段名称是否引用多单元格区域。</param>
    private static byte[] CreateNamedFixedForwardWorkbook(bool includeNumber = true,
        bool areaNumber = false)
    {
        using var input = new MemoryStream(CreateWorkbook(false));
        using var workbook = new XSSFWorkbook(input);
        var number = workbook.CreateName();
        number.NameName = "OrderNumber";
        if (includeNumber)
            number.RefersToFormula = areaNumber ? "'单据'!$B$2:$C$2" : "'单据'!$B$2";
        else
            workbook.RemoveName(number);
        var title = workbook.CreateName();
        title.NameName = "OrderTitle";
        title.RefersToFormula = "'单据'!$B$3";
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含偏移列表和列表命名锚点的 XLSX 前向工作簿。
    /// </summary>
    /// <param name="includeAnchor">是否创建列表起始命名锚点。</param>
    /// <param name="areaAnchor">列表起始名称是否引用多单元格区域。</param>
    private static byte[] CreateNamedListForwardWorkbook(bool includeAnchor = true,
        bool areaAnchor = false)
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(2);
        header.CreateCell(2).SetCellValue("Code");
        header.CreateCell(3).SetCellValue("Quantity");
        var first = sheet.CreateRow(3);
        first.CreateCell(2).SetCellValue("A-1");
        first.CreateCell(3).SetCellValue(2);
        var second = sheet.CreateRow(4);
        second.CreateCell(2).SetCellValue("B-2");
        second.CreateCell(3).SetCellValue(3);
        var name = workbook.CreateName();
        name.NameName = "LinesStart";
        if (includeAnchor)
            name.RefersToFormula = areaAnchor ? "'明细'!$C$3:$D$3" : "'明细'!$C$3";
        else
            workbook.RemoveName(name);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含同表固定单元格和列表区域的 HSSF 前向工作簿。
    /// </summary>
    /// <param name="invalidNumber">是否将固定数字写为不可转换文本。</param>
    private static byte[] CreateFixedForwardHssfWorkbook(bool invalidNumber = false)
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        sheet.CreateRow(0).CreateCell(3).SetCellValue("PO-001");
        sheet.CreateRow(1).CreateCell(3).SetCellValue(invalidNumber ? "wrong" : "42");
        var header = sheet.CreateRow(3);
        header.CreateCell(0).SetCellValue("Code");
        header.CreateCell(1).SetCellValue("Quantity");
        CreateLine(sheet, 4, "A", 2);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含工作簿级命名锚点的 HSSF 前向工作簿。
    /// </summary>
    private static byte[] CreateNamedFixedHssfForwardWorkbook(bool includeNumber = true,
        bool areaNumber = false)
    {
        using var input = new MemoryStream(CreateWorkbook(true));
        using var workbook = new HSSFWorkbook(input);
        var number = workbook.CreateName();
        number.NameName = "OrderNumber";
        if (includeNumber)
            number.RefersToFormula = areaNumber ? "'单据'!$B$2:$C$2" : "'单据'!$B$2";
        else
            workbook.RemoveName(number);
        var title = workbook.CreateName();
        title.NameName = "OrderTitle";
        title.RefersToFormula = "'单据'!$B$3";
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含偏移列表和列表命名锚点的 HSSF 前向工作簿。
    /// </summary>
    private static byte[] CreateNamedListHssfForwardWorkbook(bool includeAnchor = true,
        bool areaAnchor = false, bool localName = false)
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        var header = sheet.CreateRow(2);
        header.CreateCell(2).SetCellValue("Code");
        header.CreateCell(3).SetCellValue("Quantity");
        var first = sheet.CreateRow(3);
        first.CreateCell(2).SetCellValue("A-1");
        first.CreateCell(3).SetCellValue(2);
        var second = sheet.CreateRow(4);
        second.CreateCell(2).SetCellValue("B-2");
        second.CreateCell(3).SetCellValue(3);
        var name = workbook.CreateName();
        name.NameName = "LinesStart";
        if (localName)
            name.SheetIndex = workbook.GetSheetIndex(sheet);
        if (includeAnchor)
            name.RefersToFormula = areaAnchor ? "'明细'!$C$3:$D$3" : "'明细'!$C$3";
        else
            workbook.RemoveName(name);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建固定字段位于其他工作表的 HSSF 前向工作簿。
    /// </summary>
    private static byte[] CreateCrossSheetFixedForwardHssfWorkbook()
    {
        using var workbook = new HSSFWorkbook();
        var header = workbook.CreateSheet("单据");
        header.CreateRow(0).CreateCell(1).SetCellValue("PO-CROSS");
        header.CreateRow(1).CreateCell(1).SetCellValue(84);
        var lines = workbook.CreateSheet("明细");
        CreateHeader(lines, "Code", "Quantity");
        CreateLine(lines, 1, "A", 2);
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建超过单个复制块大小的前向批次工作簿。
    /// </summary>
    /// <param name="rowCount">明细行数量。</param>
    private static byte[] CreateLargeForwardWorkbook(int rowCount)
    {
        using var workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        CreateHeader(sheet, "Code", "Quantity");
        for (var index = 1; index <= rowCount; index++)
        {
            var row = sheet.CreateRow(index);
            row.CreateCell(0).SetCellValue(
                $"SKU-{index:D6}-{index * 2654435761L:X16}-{new string((char)('A' + index % 26), 64)}");
            row.CreateCell(1).SetCellValue(index % 1000);
        }

        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含单个隐式动态字典的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateDynamicForwardWorkbook(bool unknownColumn = false)
    {
        using IWorkbook workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("子项");
        CreateHeader(sheet, "ParentId", "Name", "Zone", "Score");
        var row = sheet.CreateRow(1);
        row.CreateCell(0).SetCellValue("P-1");
        row.CreateCell(1).SetCellValue("detail");
        row.CreateCell(2).SetCellValue("North");
        row.CreateCell(3).SetCellValue(12.5d);
        if (unknownColumn)
        {
            sheet.GetRow(0).CreateCell(4).SetCellValue("Unknown");
            row.CreateCell(4).SetCellValue("unexpected");
        }
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含两个隐式动态字典列的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateMultipleImplicitDynamicForwardWorkbook()
    {
        using IWorkbook workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        CreateHeader(sheet, "Code", "Zone", "Note");
        var row = sheet.CreateRow(1);
        row.CreateCell(0).SetCellValue("A-1");
        row.CreateCell(1).SetCellValue("North");
        row.CreateCell(2).SetCellValue("Extended");
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含单个显式动态列组的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateExplicitDynamicForwardWorkbook(bool unknownColumn = false)
    {
        using IWorkbook workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        CreateHeader(sheet, "Code", "Zone", "Score");
        var row = sheet.CreateRow(1);
        row.CreateCell(0).SetCellValue("A-1");
        row.CreateCell(1).SetCellValue("North");
        row.CreateCell(2).SetCellValue(12.5d);
        if (unknownColumn)
        {
            sheet.GetRow(0).CreateCell(3).SetCellValue("Unknown");
            row.CreateCell(3).SetCellValue("unexpected");
        }
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含多个显式动态列组的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateMultipleExplicitDynamicForwardWorkbook()
    {
        using IWorkbook workbook = new XSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        CreateHeader(sheet, "Code", "Zone", "Other");
        var row = sheet.CreateRow(1);
        row.CreateCell(0).SetCellValue("A-1");
        row.CreateCell(1).SetCellValue("North");
        row.CreateCell(2).SetCellValue("Extended");
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建只包含指定 Sheet 的 Forward 模板工作簿。
    /// </summary>
    private static byte[] CreateForwardTemplateWorkbook(string sheetName, bool xls = false)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        var sheet = workbook.CreateSheet(sheetName);
        CreateHeader(sheet, "Code", "Zone", "Score");
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含日期、数值格式和 Unicode 文本的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateStyledForwardWorkbook(bool date1904)
    {
        using var workbook = new XSSFWorkbook();
        var workbookProperties = workbook.GetCTWorkbook().workbookPr;
        workbookProperties.date1904 = date1904;
        workbookProperties.date1904Specified = true;
        var sheet = workbook.CreateSheet("明细");
        CreateHeader(sheet, "Code", "OccurredAt", "Amount", "Description");
        var row = sheet.CreateRow(1);
        row.CreateCell(0).SetCellValue("商品😀");
        var dateCell = row.CreateCell(1);
        dateCell.SetCellValue(new DateTime(2024, 2, 29));
        var dataFormat = workbook.CreateDataFormat();
        var dateStyle = workbook.CreateCellStyle();
        dateStyle.DataFormat = dataFormat.GetFormat("yyyy-mm-dd");
        dateCell.CellStyle = dateStyle;
        var amountCell = row.CreateCell(2);
        amountCell.SetCellValue(1234.5d);
        var amountStyle = workbook.CreateCellStyle();
        amountStyle.DataFormat = dataFormat.GetFormat("#,##0.00");
        amountCell.CellStyle = amountStyle;
        row.CreateCell(3).SetCellValue("中文\n换行");
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含日期、数值格式和 Unicode 文本的 HSSF 工作簿。
    /// </summary>
    private static byte[] CreateStyledForwardHssfWorkbook()
    {
        using var workbook = new HSSFWorkbook();
        var sheet = workbook.CreateSheet("明细");
        CreateHeader(sheet, "Code", "OccurredAt", "Amount", "Description");
        var row = sheet.CreateRow(1);
        row.CreateCell(0).SetCellValue("商品😀");
        var dataFormat = workbook.CreateDataFormat();
        var dateCell = row.CreateCell(1);
        dateCell.SetCellValue(new DateTime(2024, 2, 29));
        var dateStyle = workbook.CreateCellStyle();
        dateStyle.DataFormat = dataFormat.GetFormat("yyyy-mm-dd");
        dateCell.CellStyle = dateStyle;
        var amountCell = row.CreateCell(2);
        amountCell.SetCellValue(1234.5d);
        var amountStyle = workbook.CreateCellStyle();
        amountStyle.DataFormat = dataFormat.GetFormat("#,##0.00");
        amountCell.CellStyle = amountStyle;
        row.CreateCell(3).SetCellValue("中文\n换行");
        using var output = new MemoryStream();
        workbook.Write(output, true);
        return output.ToArray();
    }

    /// <summary>
    /// 创建用于指定格式导出的模板工作簿。
    /// </summary>
    private static MemoryStream CreateTemplateWorkbook(bool xls)
    {
        using IWorkbook workbook = xls ? new HSSFWorkbook() : new XSSFWorkbook();
        workbook.CreateSheet("父项");
        workbook.CreateSheet("子项");
        workbook.CreateSheet("明细");
        var output = new MemoryStream();
        workbook.Write(output, true);
        output.Position = 0;
        return output;
    }

    /// <summary>
    /// 写入列表表头。
    /// </summary>
    private static void CreateHeader(ISheet sheet, params string[] headers)
    {
        var row = sheet.CreateRow(0);
        for (var index = 0; index < headers.Length; index++)
            row.CreateCell(index).SetCellValue(headers[index]);
    }

    /// <summary>
    /// 写入采购明细行。
    /// </summary>
    private static void CreateLine(ISheet sheet, int rowIndex, string code, int quantity)
    {
        var row = sheet.CreateRow(rowIndex);
        row.CreateCell(0).SetCellValue(code);
        row.CreateCell(1).SetCellValue(quantity);
    }

    /// <summary>
    /// 写入采购明细行并允许直接构造转换错误。
    /// </summary>
    private static void CreateLine(ISheet sheet, int rowIndex, string code, string quantity)
    {
        var row = sheet.CreateRow(rowIndex);
        row.CreateCell(0).SetCellValue(code);
        row.CreateCell(1).SetCellValue(quantity);
    }

    /// <summary>
    /// 写入税额明细行。
    /// </summary>
    private static void CreateTax(ISheet sheet, int rowIndex, string name, double amount)
    {
        var row = sheet.CreateRow(rowIndex);
        row.CreateCell(0).SetCellValue(name);
        row.CreateCell(1).SetCellValue(amount);
    }

    /// <summary>
    /// 保存用于投影取消测试的根模型。
    /// </summary>
    private sealed class ProjectionCancellationDocument
    {
        /// <summary>
        /// 获取或设置用于投影的明细集合。
        /// </summary>
        public IEnumerable<BatchLine> Lines { get; set; }
    }

    /// <summary>
    /// 在指定枚举过程中取消令牌的实体集合。
    /// </summary>
    private sealed class CancelOnEnumerationEnumerable<T> : IEnumerable<T>
    {
        /// <summary>
        /// 保存用于投影的实体。
        /// </summary>
        private readonly IReadOnlyList<T> _items;
        /// <summary>
        /// 保存取消令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancellationTokenSource;
        /// <summary>
        /// 保存触发取消的枚举序号。
        /// </summary>
        private readonly int _cancelOnEnumeration;

        internal CancelOnEnumerationEnumerable(IReadOnlyList<T> items,
            CancellationTokenSource cancellationTokenSource, int cancelOnEnumeration)
        {
            _items = items;
            _cancellationTokenSource = cancellationTokenSource;
            _cancelOnEnumeration = cancelOnEnumeration;
        }

        internal int EnumerationCount { get; private set; }
        internal int CancelledEnumerationItemsRead { get; private set; }

        public IEnumerator<T> GetEnumerator()
        {
            return Enumerate(++EnumerationCount).GetEnumerator();
        }

        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();

        private IEnumerable<T> Enumerate(int enumeration)
        {
            foreach (var item in _items)
            {
                if (enumeration == _cancelOnEnumeration)
                {
                    CancelledEnumerationItemsRead++;
                    _cancellationTokenSource.Cancel();
                }

                yield return item;
            }
        }
    }

    /// <summary>
    /// 保存实体布局批次测试根模型。
    /// </summary>
    public sealed class BatchDocument
    {
        /// <summary>
        /// 获取或设置单据编号。
        /// </summary>
        public int Number { get; set; }

        /// <summary>
        /// 获取或设置单据标题。
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 获取或设置采购明细。
        /// </summary>
        public List<BatchLine> Lines { get; set; } = new();

        /// <summary>
        /// 获取或设置税额明细。
        /// </summary>
        public List<BatchTax> Taxes { get; set; } = new();
    }

    /// <summary>
    /// 保存采购明细行。
    /// </summary>
    public sealed class BatchLine
    {
        /// <summary>
        /// 获取或设置商品编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }
    }

    /// <summary>
    /// 保存税额明细行。
    /// </summary>
    public sealed class BatchTax
    {
        /// <summary>
        /// 获取或设置税种名称。
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 获取或设置税额。
        /// </summary>
        public decimal Amount { get; set; }
    }

    /// <summary>
    /// 保存关系批次测试根模型。
    /// </summary>
    public sealed class RelationBatchDocument
    {
        /// <summary>
        /// 获取或设置父项集合。
        /// </summary>
        public List<RelationBatchParent> Parents { get; set; } = new();
        /// <summary>
        /// 获取或设置子项集合。
        /// </summary>
        public List<RelationBatchChild> Children { get; set; } = new();
    }

    /// <summary>
    /// 保存关系批次测试父项。
    /// </summary>
    public sealed class RelationBatchParent
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string Id { get; set; }
        /// <summary>
        /// 获取父项的子项集合。
        /// </summary>
        public List<RelationBatchChild> Children { get; } = new();
    }

    /// <summary>
    /// 保存包含动态字典的关系批次测试子项。
    /// </summary>
    public sealed class RelationBatchChild
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string ParentId { get; set; }
        /// <summary>
        /// 获取或设置子项名称。
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置动态字段。
        /// </summary>
        [DynamicColumn]
        public Dictionary<string, object> CustomFields { get; set; } = new();
    }

    /// <summary>
    /// 保存关系错误暂存测试根模型。
    /// </summary>
    public sealed class RelationErrorDocument
    {
        /// <summary>
        /// 获取或设置父项集合。
        /// </summary>
        public List<RelationErrorParent> Parents { get; set; } = new();

        /// <summary>
        /// 获取或设置子项集合。
        /// </summary>
        public List<RelationErrorChild> Children { get; set; } = new();
    }

    /// <summary>
    /// 保存关系错误暂存测试父项。
    /// </summary>
    public sealed class RelationErrorParent
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 获取父项的子项集合。
        /// </summary>
        public List<RelationErrorChild> Children { get; } = new();
    }

    /// <summary>
    /// 保存关系错误暂存测试子项。
    /// </summary>
    public sealed class RelationErrorChild
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string ParentId { get; set; }

        /// <summary>
        /// 获取或设置数量。
        /// </summary>
        public int Quantity { get; set; }
    }

    /// <summary>
    /// 保存多个隐式动态字典的 Forward 测试根模型。
    /// </summary>
    public sealed class MultipleImplicitDynamicDocument
    {
        /// <summary>
        /// 获取 Forward 测试明细。
        /// </summary>
        public List<MultipleImplicitDynamicLine> Lines { get; } = new();
    }

    /// <summary>
    /// 保存多个隐式动态字典的 Forward 测试明细。
    /// </summary>
    public sealed class MultipleImplicitDynamicLine
    {
        /// <summary>
        /// 获取或设置商品编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置商品属性字典。
        /// </summary>
        [DynamicColumn]
        public Dictionary<string, object> Attributes { get; set; } = new();

        /// <summary>
        /// 获取或设置扩展属性字典。
        /// </summary>
        [DynamicColumn]
        public Dictionary<string, object> Extensions { get; set; } = new();
    }

    /// <summary>
    /// 保存显式动态列组测试根模型。
    /// </summary>
    public sealed class ExplicitDynamicDocument
    {
        /// <summary>
        /// 获取显式动态列组测试明细。
        /// </summary>
        public List<ExplicitDynamicLine> Lines { get; } = new();
    }

    /// <summary>
    /// 保存显式动态列组测试明细。
    /// </summary>
    public sealed class ExplicitDynamicLine
    {
        /// <summary>
        /// 获取或设置商品编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置动态字段。
        /// </summary>
        public Dictionary<string, object> Fields { get; set; } = new();

    }

    /// <summary>
    /// 保存动态字典为空的 Forward 测试根模型。
    /// </summary>
    public sealed class NullDynamicDocument
    {
        /// <summary>
        /// 获取 Forward 测试明细。
        /// </summary>
        public List<NullDynamicLine> Lines { get; } = new();
    }

    /// <summary>
    /// 保存动态字典为空的 Forward 测试明细。
    /// </summary>
    public sealed class NullDynamicLine
    {
        /// <summary>
        /// 获取或设置商品编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置可由 Forward 路径创建的动态字典。
        /// </summary>
        public Dictionary<string, object> Fields { get; set; }
    }

    /// <summary>
    /// 保存多个显式动态列组边界测试根模型。
    /// </summary>
    public sealed class MultipleExplicitDynamicDocument
    {
        /// <summary>
        /// 获取多个显式动态列组测试明细。
        /// </summary>
        public List<MultipleExplicitDynamicLine> Lines { get; } = new();
    }

    /// <summary>
    /// 保存多个显式动态列组边界测试明细。
    /// </summary>
    public sealed class MultipleExplicitDynamicLine
    {
        /// <summary>
        /// 获取或设置商品编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置第一个动态字段组。
        /// </summary>
        public Dictionary<string, object> Fields { get; set; } = new();

        /// <summary>
        /// 获取或设置第二个动态字段组。
        /// </summary>
        public Dictionary<string, object> OtherFields { get; set; } = new();
    }

    /// <summary>
    /// 保存具有 getter-only 明细集合的测试根模型。
    /// </summary>
    public sealed class GetterOnlyBatchDocument
    {
        /// <summary>
        /// 获取明细集合。
        /// </summary>
        public List<BatchLine> Lines { get; } = new();
    }

    /// <summary>
    /// 保存摘要模式无需访问的结果集合测试模型。
    /// </summary>
    public sealed class SummaryOnlyNoCollectionDocument
    {
        /// <summary>
        /// 获取结果集合；摘要模式不应访问该成员。
        /// </summary>
        public List<BatchLine> Lines =>
            throw new InvalidOperationException("摘要模式不应访问结果集合。");
    }

    /// <summary>
    /// 保存唯一性 Forward 测试根模型。
    /// </summary>
    public sealed class UniqueForwardDocument
    {
        /// <summary>
        /// 获取唯一性测试明细。
        /// </summary>
        public List<UniqueForwardLine> Lines { get; } = new();

        /// <summary>
        /// 获取第二个唯一性测试明细集合。
        /// </summary>
        public List<UniqueForwardLine> OtherLines { get; } = new();
    }

    /// <summary>
    /// 保存唯一性 Forward 测试明细。
    /// </summary>
    public sealed class UniqueForwardLine
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
    /// 保存固定单元格 Forward 测试单据。
    /// </summary>
    public sealed class FixedForwardDocument
    {
        /// <summary>
        /// 获取或设置单据标题。
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 获取或设置单据编号。
        /// </summary>
        public int Number { get; set; }

        /// <summary>
        /// 获取明细集合。
        /// </summary>
        public List<BatchLine> Lines { get; } = new();
    }

    /// <summary>
    /// 保存样式读取测试根模型。
    /// </summary>
    public sealed class StyledForwardDocument
    {
        /// <summary>
        /// 获取样式读取测试明细。
        /// </summary>
        public List<StyledForwardLine> Lines { get; } = new();
    }

    /// <summary>
    /// 保存样式读取测试明细。
    /// </summary>
    public sealed class StyledForwardLine
    {
        /// <summary>
        /// 获取或设置商品编码。
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 获取或设置发生日期。
        /// </summary>
        public DateTime OccurredAt { get; set; }

        /// <summary>
        /// 获取或设置金额。
        /// </summary>
        public decimal Amount { get; set; }

        /// <summary>
        /// 获取或设置描述文本。
        /// </summary>
        public string Description { get; set; }
    }

    /// <summary>
    /// 捕获前向路径传入的数值格式信息。
    /// </summary>
    private sealed class ForwardStyleCaptureConverter : INamedExcelValueConverter
    {
        public string Name => "forward-style-capture";

        public ExcelCellKind Kind { get; private set; }

        public int FormatIndex { get; private set; }

        public bool CanConvert(Type propertyType) => propertyType == typeof(decimal);

        public bool TryConvertFrom(ExcelConversionContext context, out object value)
        {
            Kind = context.Cell?.Kind ?? ExcelCellKind.Empty;
            FormatIndex = context.Cell?.FormatIndex ?? 0;
            value = Convert.ToDecimal(context.Cell?.Value, CultureInfo.InvariantCulture);
            return true;
        }

        public bool TryConvertTo(ExcelConversionContext context, out object value)
        {
            value = null;
            return false;
        }
    }

    /// <summary>
    /// 提供可检查所有权的内存流。
    /// </summary>
    private sealed class TrackingStream : MemoryStream
    {
        internal TrackingStream(byte[] buffer) : base(buffer, writable: false)
        {
        }

        /// <summary>
        /// 获取读取调用次数。
        /// </summary>
        internal int ReadCount { get; private set; }

        public override int Read(byte[] buffer, int offset, int count)
        {
            ReadCount++;
            return base.Read(buffer, offset, count);
        }

        public override int Read(Span<byte> buffer)
        {
            ReadCount++;
            return base.Read(buffer);
        }

        public override int ReadByte()
        {
            ReadCount++;
            return base.ReadByte();
        }

        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            ReadCount++;
            return base.ReadAsync(buffer, offset, count, cancellationToken);
        }

        public override ValueTask<int> ReadAsync(Memory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            ReadCount++;
            return base.ReadAsync(buffer, cancellationToken);
        }
    }

    /// <summary>
    /// 获取当前 NPOI 前向输入暂存文件清单。
    /// </summary>
    /// <returns>按路径序排序的暂存文件路径。</returns>
    private static string[] ForwardTemporaryFiles() => new[]
    {
        "bing-offices-forward-input-*.bin",
        "bing-offices-forward-*.xlsx",
        "bing-offices-forward-hssf-*.xlsx",
        "bing-offices-forward-hssf-sheet-*.xml"
    }.SelectMany(pattern => Directory.GetFiles(Path.GetTempPath(), pattern))
        .Distinct(StringComparer.OrdinalIgnoreCase)
        .OrderBy(path => path, StringComparer.Ordinal)
        .ToArray();

    /// <summary>
    /// 提供不可定位但可读取的测试流。
    /// </summary>
    private sealed class NonSeekableTrackingStream : Stream
    {
        private readonly MemoryStream _inner;

        internal NonSeekableTrackingStream(byte[] buffer)
        {
            _inner = new MemoryStream(buffer, writable: false);
        }

        public override bool CanRead => !_inner.CanRead ? false : true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() => throw new NotSupportedException();
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override int Read(Span<byte> buffer) => _inner.Read(buffer);
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) => _inner.ReadAsync(buffer, offset, count, cancellationToken);
        public override ValueTask<int> ReadAsync(Memory<byte> buffer,
            CancellationToken cancellationToken = default) => _inner.ReadAsync(buffer, cancellationToken);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 在指定异步读取次数后触发取消的不可定位测试流。
    /// </summary>
    private sealed class CancellationOnReadStream : Stream
    {
        /// <summary>
        /// 保存实际 XLSX 数据的只读流。
        /// </summary>
        private readonly MemoryStream _inner;

        /// <summary>
        /// 读取达到阈值后触发的取消令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation;

        /// <summary>
        /// 触发取消所需的读取次数。
        /// </summary>
        private readonly int _cancelAfterReads;

        /// <summary>
        /// 已完成的异步读取次数。
        /// </summary>
        private int _readCount;

        /// <summary>
        /// 初始化一个 <see cref="CancellationOnReadStream" /> 类型的实例。
        /// </summary>
        /// <param name="buffer">只读的 XLSX 数据。</param>
        /// <param name="cancellation">读取达到阈值后取消的令牌源。</param>
        /// <param name="cancelAfterReads">触发取消所需的读取次数。</param>
        internal CancellationOnReadStream(byte[] buffer, CancellationTokenSource cancellation,
            int cancelAfterReads)
        {
            _inner = new MemoryStream(buffer, writable: false);
            _cancellation = cancellation ?? throw new ArgumentNullException(nameof(cancellation));
            _cancelAfterReads = cancelAfterReads;
        }

        /// <inheritdoc />
        public override bool CanRead => _inner.CanRead;
        /// <inheritdoc />
        public override bool CanSeek => false;
        /// <inheritdoc />
        public override bool CanWrite => false;
        /// <inheritdoc />
        public override long Length => throw new NotSupportedException();
        /// <inheritdoc />
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        /// <inheritdoc />
        public override void Flush() => throw new NotSupportedException();
        /// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count)
        {
            var read = _inner.Read(buffer, offset, count);
            AfterRead(read);
            return read;
        }

        /// <inheritdoc />
        public override int Read(Span<byte> buffer)
        {
            var read = _inner.Read(buffer);
            AfterRead(read);
            return read;
        }

        /// <inheritdoc />
        public override async ValueTask<int> ReadAsync(Memory<byte> buffer,
            CancellationToken cancellationToken = default)
        {
            var count = await _inner.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
            AfterRead(count);
            return count;
        }

        /// <inheritdoc />
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) => ReadAsync(buffer.AsMemory(offset, count),
            cancellationToken).AsTask();

        /// <summary>
        /// 记录读取次数并在达到阈值后取消导入。
        /// </summary>
        /// <param name="count">本次读取的字节数。</param>
        private void AfterRead(int count)
        {
            if (count > 0 && ++_readCount >= _cancelAfterReads)
                _cancellation.Cancel();
        }

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();
        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
