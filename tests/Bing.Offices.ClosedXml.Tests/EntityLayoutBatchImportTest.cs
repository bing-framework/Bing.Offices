using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Bing.Offices.Attributes;
using Bing.Offices.Configurations;
using Bing.Offices.ClosedXml.Extensions;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// ClosedXML 实体布局分批导入测试。
/// </summary>
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
        using var source = new TrackingStream(CreateWorkbook(invalidQuantity: true));
        using var templateStream = new TrackingStream(CreateWorkbook());
        using var failureTarget = new MemoryStream();
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
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
    /// 验证前向实体批次对未知失败工作簿模式在计划阶段拒绝请求。
    /// </summary>
    [Fact]
    public void ForwardFailureWorkbook_UnknownMode_ShouldBeRejectedBeforeReading()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
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
            importer.ImportEntityBatches(new MemoryStream(CreateWorkbook()), CreateLayout(),
                options, _ => { }));

        Assert.Equal(BingOfficesStage.Plan, exception.Stage);
        Assert.Empty(failure.ToArray());
    }

    /// <summary>
    /// 验证实体布局按列表区域和批次大小交付实体。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_ShouldProjectRegions()
    {
        var layout = CreateLayout();
        var source = new TrackingStream(CreateWorkbook());
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();

        var summary = importer.ImportEntityBatches(source, layout,
            new ExcelEntityBatchImportOptions(1), batches.Add);

        Assert.True(source.CanRead);
        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(2, summary.Regions.Count);
        Assert.Equal(5, summary.RowsRead);
        Assert.Equal(5, summary.SucceededRows);
        Assert.Equal(new[] { "Lines", "Lines", "Lines", "Taxes", "Taxes" },
            batches.Select(batch => batch.PropertyName));
        Assert.Equal(new[] { 2, 3, 4 }, batches.Take(3).Select(batch => batch.FirstRowIndex));
        Assert.Equal("采购单", summary.Result.Entity.Title);
        Assert.Equal(42, summary.Result.Entity.Number);
        Assert.Equal(3, summary.Result.Entity.Lines.Count);
        Assert.Equal(2, summary.Result.Entity.Taxes.Count);
    }

    /// <summary>
    /// 验证异步模板批次等待回调并保留模板流所有权。
    /// </summary>
    [Fact]
    public async Task ImportForTemplateBatchesAsync_ShouldAwaitCallbacksAndLeaveTemplateOpen()
    {
        var layout = CreateLayout();
        var bytes = CreateWorkbook();
        using var source = new TrackingStream(bytes);
        using var templateStream = new TrackingStream(bytes);
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
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
    /// 验证固定单元格转换错误以实体级诊断批次交付。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_ShouldDeliverDiagnosticBatch()
    {
        var source = new TrackingStream(CreateWorkbook(invalidNumber: true));
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateLayout(),
            new ExcelEntityBatchImportOptions(2), batches.Add);

        Assert.False(summary.IsSuccess);
        var diagnostic = Assert.Single(batches.Where(batch => batch.IsDiagnostic));
        Assert.Equal(-1, diagnostic.RegionIndex);
        Assert.Empty(diagnostic.Items);
        Assert.NotEmpty(diagnostic.Errors);
        Assert.NotEmpty(summary.GlobalErrors);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证区域转换失败按原始行号交付并计入失败摘要。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_ShouldDeliverRegionFailureAtSourceRow()
    {
        var source = new TrackingStream(CreateWorkbook(invalidQuantity: true));
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
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
    [Fact]
    public void EntityFailureWorkbook_ErrorRowsOnly_ShouldRetainEntityStructureRows()
    {
        var layout = ExcelEntity.Layout<BatchDocument>(builder => builder
            .Cell("明细", "G1", document => document.Number)
            .ListRegion("明细", "A1", document => document.Lines,
                region => region.End("B4"))
            .ListRegion("明细", "D5", document => document.Taxes,
                region => region.End("E7")));
        using var input = new MemoryStream(CreateMultiRegionFailureWorkbook(), writable: false);
        using var failure = new MemoryStream();
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            });

        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout, options);

        Assert.False(result.IsSuccess);
        using var output = new XLWorkbook(new MemoryStream(failure.ToArray(), writable: false));
        var sheet = output.Worksheet("明细");
        Assert.Equal("Code", sheet.Cell(1, 1).GetString());
        Assert.Equal(42, sheet.Cell(1, 7).GetDouble());
        Assert.Equal("B", sheet.Cell(2, 1).GetString());
        Assert.Equal("Name", sheet.Cell(3, 4).GetString());
        Assert.True(output.TryGetWorksheet("_ImportErrors", out _));
    }

    /// <summary>
    /// 验证实体失败工作簿保留没有错误的单据头和其他列表工作表。
    /// </summary>
    [Fact]
    public void EntityFailureWorkbook_ErrorRowsOnly_ShouldRetainContextSheets()
    {
        using var input = new MemoryStream(CreateWorkbook(invalidQuantity: true), writable: false);
        using var failure = new MemoryStream();
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            });

        var result = new ClosedXmlExcelImporter().ImportEntity(input, CreateLayout(), options);

        Assert.False(result.IsSuccess);
        using var output = new XLWorkbook(new MemoryStream(failure.ToArray(), writable: false));
        Assert.Equal("42", output.Worksheet("单据").Cell("B1").GetString());
        Assert.Equal("采购单", output.Worksheet("单据").Cell("B2").GetString());
        Assert.Equal("Code", output.Worksheet("明细").Cell("A1").GetString());
        Assert.Equal("Name", output.Worksheet("税额").Cell("A1").GetString());
    }

    /// <summary>
    /// 验证实体失败工作簿保留分页小计和最终尾部行。
    /// </summary>
    [Fact]
    public void EntityFailureWorkbook_ErrorRowsOnly_ShouldRetainPaginationAndFooterRows()
    {
        var layout = ExcelEntity.Layout<BatchDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines, region => region
                .PageBreak(2)
                .PageSubtotal("PAGE_TOTAL", footer => footer.Cell("B1", "subtotal"))
                .Footer("END", footer => footer.Cell("B1", "total"))));
        using var input = new MemoryStream(CreatePaginationFailureWorkbook(), writable: false);
        using var failure = new MemoryStream();
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            });

        var result = new ClosedXmlExcelImporter().ImportEntity(input, layout, options);

        Assert.False(result.IsSuccess);
        using var output = new XLWorkbook(new MemoryStream(failure.ToArray(), writable: false));
        var sheet = output.Worksheet("明细");
        Assert.Equal("Code", sheet.Cell(1, 1).GetString());
        Assert.Equal("B", sheet.Cell(2, 1).GetString());
        Assert.Equal("PAGE_TOTAL", sheet.Cell(3, 1).GetString());
        Assert.Equal("subtotal", sheet.Cell(3, 2).GetString());
        Assert.Equal("END", sheet.Cell(4, 1).GetString());
        Assert.Equal("total", sheet.Cell(4, 2).GetString());
        Assert.Contains(sheet.MergedRanges,
            merge => merge.RangeAddress.ToString() == "C1:D1");
        Assert.True(sheet.Cell(1, 3).Style.Font.Bold);
    }

    /// <summary>
    /// 验证实体失败工作簿保留显式动态列的表头和错误值。
    /// </summary>
    [Fact]
    public void EntityFailureWorkbook_ErrorRowsOnly_ShouldRetainDynamicColumnBoundary()
    {
        using var input = new MemoryStream(CreateExplicitDynamicFailureWorkbook(), writable: false);
        using var failure = new MemoryStream();
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            });

        var result = new ClosedXmlExcelImporter().ImportEntity(input,
            CreateExplicitDynamicForwardLayout(), options);

        Assert.False(result.IsSuccess);
        using var output = new XLWorkbook(new MemoryStream(failure.ToArray(), writable: false));
        var sheet = output.Worksheet("明细");
        Assert.Equal("Code", sheet.Cell(1, 1).GetString());
        Assert.Equal("Zone", sheet.Cell(1, 2).GetString());
        Assert.Equal("Score", sheet.Cell(1, 3).GetString());
        Assert.Equal("A", sheet.Cell(2, 1).GetString());
        Assert.Equal("North", sheet.Cell(2, 2).GetString());
        Assert.Equal("wrong", sheet.Cell(2, 3).GetString());
    }

    /// <summary>
    /// 验证空列表区域和没有列表区域时均不交付空批次。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_ShouldSkipEmptyAndMissingRegions()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var emptySource = new TrackingStream(CreateWorkbook(emptyLists: true));
        var emptyBatches = new List<ExcelEntityImportBatch<BatchDocument>>();

        var emptySummary = importer.ImportEntityBatches(emptySource, CreateLayout(),
            new ExcelEntityBatchImportOptions(1), emptyBatches.Add);

        Assert.True(emptySummary.IsSuccess);
        Assert.Equal(2, emptySummary.Regions.Count);
        Assert.Equal(0, emptySummary.RowsRead);
        Assert.Empty(emptyBatches);
        Assert.Empty(emptySummary.Result.Entity.Lines);
        Assert.Empty(emptySummary.Result.Entity.Taxes);

        using var source = new TrackingStream(CreateWorkbook());
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
    [Fact]
    public void ImportEntityBatches_ShouldPreserveDynamicColumnsAndRelations()
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
        var entity = new RelationBatchDocument
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
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(entity, layout, exported);
        using var source = new TrackingStream(exported.ToArray());
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
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
    /// 验证 ClosedXML 前向关系绑定在同步和异步路径遵守父键索引资源上限。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_RelationKeyLimit_ShouldReturnStructuredResourceError(
        bool asynchronous)
    {
        var layout = ExcelEntity.Layout<RelationBatchDocument>(builder => builder
            .ListRegion("父项", "A1", root => root.Parents)
            .ListRegion("子项", "A1", root => root.Children)
            .HasMany(root => root.Parents, root => root.Children,
                parent => parent.Id, child => child.ParentId, parent => parent.Children));
        using var source = new TrackingStream(CreateRelationKeyLimitForwardWorkbook());
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(new ExcelResourceLimits { MaxTrackedRelationKeys = 1 }),
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, layout, options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, layout, options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.True(summary.ErrorsTruncated);
        Assert.Contains(summary.Result.Errors, error =>
            error.Code == ExcelImportErrorCode.ResourceLimit);
        Assert.Equal(2, summary.Result.Entity.Parents.Count);
        Assert.Empty(summary.Result.Entity.Parents[0].Children);
        Assert.Empty(summary.Result.Entity.Parents[1].Children);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证页小计和最终尾部行被跳过后仍报告明细原始物理行号。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_ShouldKeepSourceRowsAfterPageSubtotalAndFooter()
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
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(entity, layout, exported);
        using var source = new TrackingStream(exported.ToArray());
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
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
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateWorkbook());
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
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateWorkbook());
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
    [Fact]
    public void ImportEntityBatches_ShouldFillGetterOnlyCollection()
    {
        var layout = ExcelEntity.Layout<GetterOnlyBatchDocument>(builder =>
            builder.ListRegion("明细", "A1", document => document.Lines));
        var entity = new GetterOnlyBatchDocument();
        entity.Lines.Add(new BatchLine { Code = "A", Quantity = 2 });
        entity.Lines.Add(new BatchLine { Code = "B", Quantity = 3 });
        using var exported = new MemoryStream();
        new ClosedXmlExcelExporter().ExportEntity(entity, layout, exported);
        using var source = new TrackingStream(exported.ToArray());
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
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
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateWorkbook());
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        Assert.Throws<OperationCanceledException>(() => importer.ImportEntityBatches(source,
            CreateLayout(), new ExcelEntityBatchImportOptions(), _ => { }, cancellation.Token));
    }

    /// <summary>
    /// 验证 ClosedXML Forward 路径按固定列列表交付同步批次。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldDeliverSingleList()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook());
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(3, summary.RowsRead);
        Assert.Equal(3, summary.SucceededRows);
        Assert.Equal(new[] { 2, 3, 4 }, batches.Select(batch => batch.FirstRowIndex));
        Assert.Equal(new[] { "A", "B", "C" }, batches.Select(batch =>
            Assert.IsType<BatchLine>(Assert.Single(batch.Items)).Code));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证已交付的批次集合在后续同步或异步读取后保持稳定。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldKeepDeliveredBatchItemsStable(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook());
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
        Assert.Equal(new[] { "A", "B", "C" }, batches.Select(batch =>
            Assert.IsType<BatchLine>(Assert.Single(batch.Items)).Code));
        Assert.Same(batches[0].Items, firstItems);
        Assert.Equal("A", Assert.IsType<BatchLine>(Assert.Single(firstItems)).Code);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证后续错误行不会修改已交付批次的错误集合。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldKeepDeliveredBatchErrorsStable(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook(invalidQuantity: true,
            invalidLastQuantity: true));
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
    /// 验证 ClosedXML Forward 路径读取固定单元格并继续读取多个列表区域。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldReadFixedCells(bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateWorkbook());
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();
        var options = new ExcelEntityBatchImportOptions(1, null,
            ExcelEntityBatchParsingMode.Forward);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateLayout(), options, batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal("采购单", summary.Result.Entity.Title);
        Assert.Equal(42, summary.Result.Entity.Number);
        Assert.Equal(3, summary.Result.Entity.Lines.Count);
        Assert.Equal(2, summary.Result.Entity.Taxes.Count);
        Assert.DoesNotContain(batches, batch => batch.IsDiagnostic);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 固定单元格转换失败时交付实体级诊断批次。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReportFixedCellFailureAsDiagnostic()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateWorkbook(invalidNumber: true));
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.False(summary.IsSuccess);
        var diagnostic = Assert.Single(batches.Where(batch => batch.IsDiagnostic));
        Assert.Equal(-1, diagnostic.RegionIndex);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, Assert.Single(diagnostic.Errors).Code);
        Assert.Equal(nameof(BatchDocument.Number), Assert.Single(diagnostic.Errors).PropertyName);
        Assert.Equal(3, summary.Result.Entity.Lines.Count);
        Assert.Equal(2, summary.Result.Entity.Taxes.Count);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 路径解析工作簿级命名锚点固定单元格。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldReadNamedFixedCells(bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateNamedFixedForwardWorkbook());
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
    /// 验证 ClosedXML Forward 路径在缺失或多单格命名锚点时于读取前拒绝请求。
    /// </summary>
    [Theory]
    [InlineData(true, "缺少命名锚点")]
    [InlineData(false, "必须引用单个单元格")]
    public void ImportEntityBatches_Forward_ShouldRejectInvalidNamedFixedCells(bool missing,
        string expectedMessage)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateNamedFixedForwardWorkbook(
            includeNumber: !missing, areaNumber: !missing));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateNamedFixedForwardLayout(),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Contains(expectedMessage, exception.Message, StringComparison.Ordinal);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 路径按命名锚点读取偏移后的列表区域。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ShouldReadNamedListRegion(bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateNamedListForwardWorkbook());
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
    /// 验证 ClosedXML Forward 路径在列表命名锚点缺失或引用区域时于读取前拒绝请求。
    /// </summary>
    [Theory]
    [InlineData(true, "缺少命名锚点")]
    [InlineData(false, "必须引用单个单元格")]
    public void ImportEntityBatches_Forward_ShouldRejectInvalidNamedListRegion(bool missing,
        string expectedMessage)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateNamedListForwardWorkbook(
            includeAnchor: !missing, areaAnchor: !missing));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportEntityBatches(source, CreateNamedListForwardLayout(),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Contains(expectedMessage, exception.Message, StringComparison.Ordinal);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 路径读取单个隐式动态字典并按定义转换值。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReadImplicitDynamicDictionary()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateDynamicForwardWorkbook());
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
    /// 验证 ClosedXML Forward 路径异步读取单个隐式动态字典。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldReadImplicitDynamicDictionary()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateDynamicForwardWorkbook());
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
    /// 验证 ClosedXML Forward 路径按目标属性将多个隐式动态字典分发。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReadMultipleImplicitDynamicDictionaries()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateMultipleImplicitDynamicForwardWorkbook());
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
    /// 验证 ClosedXML Forward 异步路径按目标属性分发多个隐式动态字典。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldReadMultipleImplicitDynamicDictionaries()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateMultipleImplicitDynamicForwardWorkbook());
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
    /// 验证 ClosedXML 完整实体布局同步路径可导出并读回多个隐式动态字典。
    /// </summary>
    [Fact]
    public void ClosedXml_EntityLayout_Buffered_ShouldRoundTripMultipleImplicitDynamicDictionaries()
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
        new ClosedXmlExcelExporter().ExportEntity(source, layout, output);
        var bytes = output.ToArray();
        using (var workbook = new XLWorkbook(new MemoryStream(bytes, writable: false)))
        {
            var sheet = workbook.Worksheet("明细");
            Assert.Equal("Code", sheet.Cell(1, 1).GetString());
            Assert.Equal("Zone", sheet.Cell(1, 2).GetString());
            Assert.Equal("Note", sheet.Cell(1, 3).GetString());
            Assert.Equal("North", sheet.Cell(2, 2).GetString());
            Assert.Equal("Extended", sheet.Cell(2, 3).GetString());
        }

        var result = new ClosedXmlExcelImporter().ImportEntity(
            new MemoryStream(bytes, writable: false), layout);
        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var line = Assert.Single(result.Entity.Lines);
        Assert.Equal("North", line.Attributes["zone"]);
        Assert.Equal("Extended", line.Extensions["note"]);
    }

    /// <summary>
    /// 验证 ClosedXML 完整实体布局异步路径可导出并读回多个隐式动态字典。
    /// </summary>
    [Fact]
    public async Task ClosedXml_EntityLayout_BufferedAsync_ShouldRoundTripMultipleImplicitDynamicDictionaries()
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
        await new ClosedXmlExcelExporter().ExportEntityAsync(source, layout, output);
        var result = await new ClosedXmlExcelImporter().ImportEntityAsync(
            new MemoryStream(output.ToArray(), writable: false), layout);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var line = Assert.Single(result.Entity.Lines);
        Assert.Equal("North", line.Attributes["zone"]);
        Assert.Equal("Extended", line.Extensions["note"]);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 路径读取单个显式动态列组并按定义转换值。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReadExplicitDynamicGroup()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateExplicitDynamicForwardWorkbook());
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
    /// 验证 ClosedXML Forward 路径异步读取单个显式动态列组。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldReadExplicitDynamicGroup()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateExplicitDynamicForwardWorkbook());
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
    /// 验证 ClosedXML Forward 路径会为显式动态组创建空字典。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldCreateNullExplicitDynamicDictionary()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateExplicitDynamicForwardWorkbook());
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
    /// 验证 ClosedXML Forward 异步路径创建空字典并忽略未声明动态表头。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldCreateNullExplicitDynamicDictionaryAndIgnoreUnknown()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateExplicitDynamicForwardWorkbook(unknownColumn: true));
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
    /// 验证 ClosedXML Forward 模式同步预检模板后读取源工作簿。
    /// </summary>
    [Fact]
    public void ImportForTemplateBatches_Forward_ShouldValidateTemplateAndReadSource()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateExplicitDynamicForwardWorkbook());
        using var template = new TrackingStream(CreateExplicitDynamicForwardWorkbook());
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
    /// 验证 ClosedXML Forward 模式异步预检模板后读取源工作簿。
    /// </summary>
    [Fact]
    public async Task ImportForTemplateBatchesAsync_Forward_ShouldValidateTemplateAndReadSource()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateExplicitDynamicForwardWorkbook());
        using var template = new TrackingStream(CreateExplicitDynamicForwardWorkbook());
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
    /// 验证 ClosedXML Forward 模式在模板缺少目标工作表时返回配置错误。
    /// </summary>
    [Fact]
    public void ImportForTemplateBatches_Forward_ShouldRejectTemplateWithoutTargetSheet()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateExplicitDynamicForwardWorkbook());
        using var template = new TrackingStream(CreateForwardTemplateWorkbook("其他"));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() =>
            importer.ImportForTemplateBatches(source, CreateExplicitDynamicForwardLayout(),
                new ExcelEntityTemplateOptions(template, leaveOpen: true),
                new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward), _ => { }));

        Assert.Contains("Sheet", exception.Message);
        Assert.True(source.CanRead);
        Assert.True(template.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 路径按未知动态列策略拒绝显式组中的未声明表头。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldRejectUnknownExplicitDynamicColumn()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateExplicitDynamicForwardWorkbook(unknownColumn: true));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() => importer.ImportEntityBatches(
            source, CreateExplicitDynamicForwardLayout(ExcelUnknownDynamicValuePolicy.Fail),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward), _ => { }));

        Assert.Contains("未知动态列", exception.Message);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML XLSX 前向模式将多个显式动态列组写入各自字典。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReadMultipleExplicitDynamicGroups()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateMultipleExplicitDynamicForwardWorkbook());
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
        Assert.DoesNotContain("otherValue", line.Fields.Keys);
        Assert.DoesNotContain("zoneValue", line.OtherFields.Keys);
        Assert.Single(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML XLSX 前向模式异步将多个显式动态列组写入各自字典。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldReadMultipleExplicitDynamicGroups()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateMultipleExplicitDynamicForwardWorkbook());
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
        Assert.Equal("A-1", line.Code);
        Assert.Equal("North", line.Fields["zoneValue"]);
        Assert.Equal("Extended", line.OtherFields["otherValue"]);
        Assert.DoesNotContain("otherValue", line.Fields.Keys);
        Assert.DoesNotContain("zoneValue", line.OtherFields.Keys);
        Assert.Single(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证多个显式动态列组仍按 Fail 策略拒绝未声明表头。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldRejectUnknownColumnInMultipleDynamicGroups()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateMultipleExplicitDynamicForwardWorkbook(unknownColumn: true));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() => importer.ImportEntityBatches(
            source, CreateMultipleExplicitDynamicForwardLayout(ExcelUnknownDynamicValuePolicy.Fail),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward), _ => { }));

        Assert.Contains("未知动态列", exception.Message);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 按未知动态列策略拒绝未声明表头。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldRejectUnknownDynamicColumn()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateDynamicForwardWorkbook(unknownColumn: true));

        var exception = Assert.Throws<BingOfficesConfigurationException>(() => importer.ImportEntityBatches(
            source, CreateDynamicForwardLayout(ExcelUnknownDynamicValuePolicy.Fail),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward), _ => { }));

        Assert.Contains("未知动态列", exception.Message);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 摘要模式仅保留批次结果和统计信息。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_SummaryOnly_ShouldAvoidRetainingEntityRows()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook());
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(2, null, ExcelEntityBatchParsingMode.Forward,
                ExcelEntityBatchResultMode.SummaryOnly), batches.Add);

        Assert.True(summary.IsSuccess);
        Assert.False(summary.HasCompleteResult);
        Assert.Equal(3, summary.RowsRead);
        Assert.Equal(3, summary.SucceededRows);
        Assert.Empty(summary.Result.Entity.Lines);
        Assert.Empty(summary.Result.Sheets[0].SourceRows);
        Assert.Equal(new[] { "A", "B", "C" }, batches.SelectMany(batch => batch.Items)
            .Cast<BatchLine>().Select(line => line.Code));
    }

    /// <summary>
    /// 验证 ClosedXML Forward 摘要模式不会访问根实体的结果集合。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_SummaryOnly_ShouldSkipResultCollection(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook());
        var batches = new List<ExcelEntityImportBatch<SummaryOnlyNoCollectionDocument>>();
        var options = new ExcelEntityBatchImportOptions(2, null,
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
        Assert.Equal(3, summary.SucceededRows);
        Assert.Equal(new[] { "A", "B", "C" }, batches.SelectMany(batch => batch.Items)
            .Cast<BatchLine>().Select(line => line.Code));
    }

    /// <summary>
    /// 验证 ClosedXML Buffered 摘要模式复用有界 Forward 路径。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_SummaryOnly_Buffered_ShouldUseForwardPath(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook());
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
        Assert.Equal(3, summary.SucceededRows);
        Assert.Equal(new[] { "A", "B", "C" }, batches.SelectMany(batch => batch.Items)
            .Cast<BatchLine>().Select(line => line.Code));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证摘要前向解析允许仅影响导出的计算列和逻辑分页。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_SummaryOnly_ShouldAllowCalculatedColumnsAndLogicalPageBreak()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook());
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
        Assert.Equal(3, summary.SucceededRows);
        Assert.Equal(new[] { "A", "B", "C" }, batches.SelectMany(batch => batch.Items)
            .Cast<BatchLine>().Select(line => line.Code));
    }

    /// <summary>
    /// 验证 ClosedXML Forward 唯一值索引达到上限时返回结构化资源错误。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_UniqueResourceLimit_ShouldBeReported(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook());
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
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntity_MultipleRegions_ShouldShareUniqueResourceLimit(bool asynchronous)
    {
        using var source = new MemoryStream(CreateMultipleSheetUniqueForwardWorkbook(), writable: false);
        var options = new ExcelEntityImportOptions(
            new ExcelResourceLimits { MaxTrackedUniqueValues = 1 },
            ExcelValidationFailureMode.Continue);
        var importer = new ClosedXmlExcelImporter();
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
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntity_MultipleRegions_ShouldShareMaxRows(bool asynchronous)
    {
        using var source = new MemoryStream(CreateMultipleSheetUniqueForwardWorkbook(), writable: false);
        var options = new ExcelEntityImportOptions(
            new ExcelResourceLimits { MaxRows = 1 },
            ExcelValidationFailureMode.Continue);
        var importer = new ClosedXmlExcelImporter();
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
    /// 验证 ClosedXML Forward 异步摘要模式不累积完整实体列表。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_SummaryOnly_ShouldAvoidRetainingEntityRows()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook());
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
        Assert.Equal(3, summary.SucceededRows);
        Assert.Empty(summary.Result.Entity.Lines);
        Assert.Empty(summary.Result.Sheets[0].SourceRows);
        Assert.Equal(3, batches.SelectMany(batch => batch.Items).Count());
    }

    /// <summary>
    /// 验证 ClosedXML Forward 摘要模式不保留最终结果中的固定字段实体。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_SummaryOnly_ShouldReleaseFixedFields(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateWorkbook());
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
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook(invalidQuantity: true,
            removeDefaultColumnStyle: true));
        using var failure = new MemoryStream();
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
        Assert.Equal(3, summary.RowsRead);
        Assert.Equal(2, summary.SucceededRows);
        Assert.Equal(1, summary.FailedRows);
        Assert.Empty(summary.Result.Entity.Lines);
        Assert.Empty(summary.Result.Sheets[0].SourceRows);
        Assert.NotEmpty(failure.ToArray());
        failure.Position = 0;
        using var workbook = new XLWorkbook(failure);
        var sheet = workbook.Worksheet("明细");
        Assert.Equal(2, sheet.LastRowUsed().RowNumber());
        Assert.Equal("Code", sheet.Cell(1, 1).GetString());
        Assert.Equal("B", sheet.Cell(2, 1).GetString());
        Assert.Equal("wrong", sheet.Cell(2, 2).GetString());
        Assert.Equal(18, sheet.Column(1).Width, 3);
        Assert.Equal(12, sheet.Column(2).Width, 3);
        Assert.True(sheet.Column(2).IsHidden);
        var errorSheet = workbook.Worksheet("_ImportErrors");
        Assert.True(errorSheet.LastRowUsed().RowNumber() >= 2);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML 低保留失败工作簿写出阶段的资源异常保留 Provider 上下文。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_SummaryOnly_FailureWriterBudget_ShouldIdentifyProvider(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook(invalidQuantity: true));
        using var failure = new MemoryStream();
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
            Assert.Equal("ClosedXML", exception.Provider);
        }
        else
        {
            var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
                importer.ImportEntityBatches(source, CreateForwardLayout(), options, _ => { }));
            Assert.Equal("ClosedXML", exception.Provider);
        }

        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML 低保留失败输出的资源异常保留 Provider 上下文。
    /// </summary>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_SummaryOnly_FailureBudget_ShouldIdentifyProvider(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook(invalidQuantity: true));
        using var failure = new MemoryStream();
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
            Assert.Equal("ClosedXML", exception.Provider);
        }
        else
        {
            var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
                importer.ImportEntityBatches(source, CreateForwardLayout(), options, _ => { }));
            Assert.Equal("ClosedXML", exception.Provider);
        }

        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 路径异步等待批次回调并保留输入流所有权。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldAwaitCallbacks()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook());
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var summary = await importer.ImportEntityBatchesAsync(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(2, null, ExcelEntityBatchParsingMode.Forward),
            async (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                await Task.Yield();
                batches.Add(batch);
            });

        Assert.True(summary.IsSuccess);
        Assert.Equal(2, batches.Count);
        Assert.Equal(new[] { 2, 4 }, batches.Select(batch => batch.FirstRowIndex));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 路径在开始读取前响应预取消。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldHonorPreCancellation()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook());
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
    /// 验证 ClosedXML Forward 路径在异步回调取消后停止读取。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldStopAfterCallbackCancellation()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook());
        using var cancellation = new CancellationTokenSource();
        var callbackCount = 0;

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => importer.ImportEntityBatchesAsync(source,
            CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            (batch, token) =>
            {
                callbackCount++;
                cancellation.Cancel();
                return Task.CompletedTask;
            }, cancellation.Token));

        Assert.Equal(1, callbackCount);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 路径保留转换错误和原始坐标。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReportConversionFailure()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook(invalidQuantity: true));
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(3, summary.RowsRead);
        Assert.Equal(2, summary.SucceededRows);
        Assert.Equal(1, summary.FailedRows);
        var failed = Assert.Single(batches.Where(batch => batch.Errors.Count > 0));
        var error = Assert.Single(failed.Errors);
        Assert.Equal(ExcelImportErrorCode.ValueConversion, error.Code);
        Assert.Equal(3, error.RowIndex);
        Assert.Equal(nameof(BatchLine.Quantity), error.PropertyName);
    }

    /// <summary>
    /// 验证 ClosedXML 前向同步批次可生成错误行失败工作簿。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldWriteErrorRowsWorkbook()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook(invalidQuantity: true));
        using var failure = new MemoryStream();
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
        using var workbook = new XLWorkbook(failure);
        Assert.NotNull(workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "明细"));
        var errorSheet = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "_ImportErrors");
        Assert.NotNull(errorSheet);
        Assert.True(errorSheet.LastRow().RowNumber() >= 2);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML 前向错误行工作簿在源坐标稳定时保留条件格式。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ErrorRowsOnly_ShouldPreserveStableConditionalFormatting(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardConditionalFormattingWorkbook());
        using var failure = new MemoryStream();
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
        Assert.Contains(summary.Result.Errors, error =>
            error.Code == ExcelImportErrorCode.ValueConversion);
        failure.Position = 0;
        using var workbook = new XLWorkbook(failure);
        var sheet = workbook.Worksheet("明细");
        var formatting = Assert.Single(sheet.ConditionalFormats);
        Assert.Equal("B2:B2", Assert.Single(formatting.Ranges).RangeAddress.ToString());
        Assert.Equal("wrong", sheet.Cell("B2").GetString());
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证错误行压缩改变条件格式坐标时 ClosedXML 前向路径会拒绝写出。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_ErrorRowsOnly_ShouldRejectCompressedConditionalFormatting(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(
            CreateForwardConditionalFormattingWorkbook(includeSuccessfulRow: true));
        using var failure = new MemoryStream();
        var importOptions = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = failure
            });
        var options = new ExcelEntityBatchImportOptions(1, importOptions,
            ExcelEntityBatchParsingMode.Forward, ExcelEntityBatchResultMode.SummaryOnly);

        var exception = asynchronous
            ? await Assert.ThrowsAsync<BingOfficesUnsupportedFeatureException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateForwardLayout(), options,
                    (_, _) => Task.CompletedTask))
            : Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
                importer.ImportEntityBatches(source, CreateForwardLayout(), options, _ => { }));

        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Contains("源坐标不变", exception.Message);
        Assert.Empty(failure.ToArray());
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML 前向同步批次可在原始单元格写入错误批注。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldAnnotateOriginalWorkbook()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook(invalidQuantity: true));
        using var failure = new MemoryStream();
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
        using var workbook = new XLWorkbook(failure);
        var sheet = workbook.Worksheet("明细");
        Assert.True(sheet.Cell("B3").HasComment);
        Assert.NotNull(workbook.Worksheets.FirstOrDefault(item => item.Name == "_ImportErrors"));
        Assert.Equal("wrong", sheet.Cell("B3").GetString());
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML 前向异步批次可生成错误行失败工作簿。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldWriteErrorRowsWorkbook()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook(invalidQuantity: true));
        using var failure = new MemoryStream();
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
        using var workbook = new XLWorkbook(failure);
        Assert.NotNull(workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "明细"));
        var errorSheet = workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "_ImportErrors");
        Assert.NotNull(errorSheet);
        Assert.True(errorSheet.LastRow().RowNumber() >= 2);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML 前向异步批次可在原始单元格写入错误批注。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldAnnotateOriginalWorkbook()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook(invalidQuantity: true));
        using var failure = new MemoryStream();
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
        using var workbook = new XLWorkbook(failure);
        Assert.True(workbook.Worksheet("明细").Cell("B3").HasComment);
        Assert.NotNull(workbook.Worksheets.FirstOrDefault(item => item.Name == "_ImportErrors"));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 模板批次在同步和异步入口均生成错误行工作簿。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportForTemplateBatches_Forward_ShouldWriteErrorRowsWorkbook(bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook(invalidQuantity: true));
        using var templateStream = new TrackingStream(CreateForwardWorkbook());
        using var failure = new MemoryStream();
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
    /// 验证 ClosedXML Forward 模板批次可在原始工作簿写入错误批注。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportForTemplateBatches_Forward_ShouldAnnotateOriginalWorkbook(bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook(invalidQuantity: true));
        using var templateStream = new TrackingStream(CreateForwardWorkbook());
        using var failure = new MemoryStream();
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
        using var workbook = new XLWorkbook(failure);
        Assert.True(workbook.Worksheet("明细").Cell("B3").HasComment);
        Assert.NotNull(workbook.Worksheets.FirstOrDefault(item => item.Name == "_ImportErrors"));
        Assert.True(source.CanRead);
        Assert.True(templateStream.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 路径遵守物理列资源上限。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldHonorMaxColumns()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook(extraColumn: true));
        var options = new ExcelEntityImportOptions(new ExcelResourceLimits
        {
            MaxColumnsPerSheet = 2
        });

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            importer.ImportEntityBatches(source, CreateForwardLayout(),
                new ExcelEntityBatchImportOptions(1, options, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Contains("最大使用列", exception.Message);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 路径在读取前执行输入字节资源限制。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldHonorMaxInputBytesBeforeRead()
    {
        var bytes = CreateForwardWorkbook();
        using var source = new TrackingStream(bytes);
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var options = new ExcelEntityImportOptions(new ExcelResourceLimits
        {
            MaxInputBytes = bytes.Length - 1
        });

        var exception = Assert.Throws<BingOfficesResourceLimitException>(() =>
            importer.ImportEntityBatches(source, CreateForwardLayout(),
                new ExcelEntityBatchImportOptions(1, options, ExcelEntityBatchParsingMode.Forward),
                _ => { }));

        Assert.Contains("最大字节数", exception.Message);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 路径可从不可定位流读取并保持调用方流打开。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReadNonSeekableInput()
    {
        using var source = new NonSeekableTrackingStream(CreateForwardWorkbook());
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(2, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(3, summary.RowsRead);
        Assert.Equal(2, batches.Count);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 路径报告数据行资源上限并停止读取。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldHonorMaxRows()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateForwardWorkbook());
        var options = new ExcelEntityImportOptions(new ExcelResourceLimits
        {
            MaxRows = 2
        });
        var batches = new List<ExcelEntityImportBatch<GetterOnlyBatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateForwardLayout(),
            new ExcelEntityBatchImportOptions(1, options, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.Equal(2, summary.RowsRead);
        Assert.Equal(2, summary.SucceededRows);
        Assert.Equal(0, summary.FailedRows);
        var error = Assert.Single(summary.Result.Errors);
        Assert.Equal(ExcelImportErrorCode.ResourceLimit, error.Code);
        Assert.Equal(4, error.RowIndex);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 路径读取日期样式、数值和 Unicode 文本。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReadStylesDatesAndUnicode()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateStyledForwardWorkbook());
        var batches = new List<ExcelEntityImportBatch<StyledForwardDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateStyledForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        var line = Assert.Single(summary.Result.Entity.Lines);
        Assert.Equal("商品😀", line.Code);
        Assert.Equal(new DateTime(2024, 2, 29), line.OccurredAt);
        Assert.Equal(1234.5m, line.Amount);
        Assert.Equal("中文\n换行", line.Description);
        Assert.Single(batches);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 路径按不同工作表读取多个列表区域。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_ShouldReadMultipleSheets()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateWorkbook());
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();

        var summary = importer.ImportEntityBatches(source, CreateMultipleSheetForwardLayout(),
            new ExcelEntityBatchImportOptions(1, null, ExcelEntityBatchParsingMode.Forward),
            batches.Add);

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(2, summary.Regions.Count);
        Assert.Equal(new[] { "明细", "明细", "明细", "税额", "税额" },
            batches.Select(batch => batch.SheetName));
        Assert.Equal(new[] { 1, 2, 3, 4, 5 }, batches.Select(batch => batch.BatchNumber));
        Assert.All(batches, batch => Assert.Same(summary.Result.Entity, batch.Entity));
        Assert.Equal(3, summary.Result.Entity.Lines.Count);
        Assert.Equal(2, summary.Result.Entity.Taxes.Count);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 ClosedXML Forward 模式异步读取多个工作表并保持全局批次编号。
    /// </summary>
    [Fact]
    public async Task ImportEntityBatchesAsync_Forward_ShouldReadMultipleSheets()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateWorkbook());
        var batches = new List<ExcelEntityImportBatch<BatchDocument>>();

        var summary = await importer.ImportEntityBatchesAsync(source,
            CreateMultipleSheetForwardLayout(),
            new ExcelEntityBatchImportOptions(2, null, ExcelEntityBatchParsingMode.Forward),
            (batch, _) =>
            {
                batches.Add(batch);
                return Task.CompletedTask;
            });

        Assert.True(summary.IsSuccess, string.Join(";", summary.Result.Errors.Select(error => error.Message)));
        Assert.Equal(2, summary.Regions.Count);
        Assert.Equal(new[] { 1, 2, 3 }, batches.Select(batch => batch.BatchNumber));
        Assert.All(batches, batch => Assert.Same(summary.Result.Entity, batch.Entity));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 Forward 模式在多个列表区域之间共享最大数据行预算。
    /// </summary>
    [Fact]
    public void ImportEntityBatches_Forward_MultipleRegions_ShouldShareMaxRows()
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateWorkbook());
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
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateWorkbook());
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
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_MultipleRegions_ShouldShareMaxErrors(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateMultipleSheetForwardErrorWorkbook());
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
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateWorkbook(invalidNumber: true));
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
    /// 验证 ClosedXML Forward 模式在多个列表区域之间共享唯一值预算。
    /// </summary>
    /// <param name="asynchronous">是否执行异步导入。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_MultipleRegions_ShouldShareMaxUniqueValues(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateMultipleSheetUniqueForwardWorkbook());
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
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Forward_MultipleRegions_ShouldShareMaxErrorsWithUniqueValidation(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new TrackingStream(CreateMultipleSheetUniqueErrorForwardWorkbook());
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
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
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
        using var provider = new ServiceCollection().AddBingOfficesClosedXml().BuildServiceProvider();
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
    /// 创建 Forward 路径使用的单 Sheet 工作簿。
    /// </summary>
    private static byte[] CreateForwardWorkbook(bool invalidQuantity = false,
        bool extraColumn = false, bool invalidLastQuantity = false,
        bool removeDefaultColumnStyle = false)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        SetHeader(sheet, "Code", "Quantity");
        sheet.Column(1).Width = 18;
        sheet.Column(2).Width = 12;
        sheet.Column(2).Hide();
        if (extraColumn)
            sheet.Cell(1, 3).Value = "Overflow";
        SetLine(sheet, 2, "A", 2);
        if (invalidQuantity)
            SetLine(sheet, 3, "B", "wrong");
        else
            SetLine(sheet, 3, "B", 3);
        if (invalidLastQuantity)
            SetLine(sheet, 4, "C", "wrong");
        else
            SetLine(sheet, 4, "C", 4);
        if (extraColumn)
        {
            sheet.Cell(2, 3).Value = "x";
            sheet.Cell(3, 3).Value = "x";
            sheet.Cell(4, 3).Value = "x";
        }
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        var bytes = output.ToArray();
        return removeDefaultColumnStyle ? RemoveDefaultColumnStyle(bytes) : bytes;
    }

    /// <summary>
    /// 从测试工作簿的列定义中移除 ClosedXML 写入的默认样式索引。
    /// </summary>
    /// <param name="source">原始 XLSX 字节。</param>
    /// <returns>保留列宽和隐藏状态的 XLSX 字节。</returns>
    private static byte[] RemoveDefaultColumnStyle(byte[] source)
    {
        using var package = new MemoryStream();
        package.Write(source, 0, source.Length);
        package.Position = 0;
        using (var archive = new ZipArchive(package, ZipArchiveMode.Update, leaveOpen: true))
        {
            const string worksheetPath = "xl/worksheets/sheet1.xml";
            var entry = archive.GetEntry(worksheetPath);
            if (entry == null)
                throw new InvalidOperationException("Forward 夹具缺少明细工作表部件。");

            XDocument document;
            using (var stream = entry.Open())
                document = XDocument.Load(stream, System.Xml.Linq.LoadOptions.PreserveWhitespace);

            XNamespace spreadsheet = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            foreach (var column in document.Descendants(spreadsheet + "col"))
            {
                var style = column.Attribute("style");
                if (string.Equals(style?.Value, "0", StringComparison.Ordinal))
                    style.Remove();
            }

            entry.Delete();
            var replacement = archive.CreateEntry(worksheetPath, CompressionLevel.Optimal);
            using var output = replacement.Open();
            document.Save(output, System.Xml.Linq.SaveOptions.DisableFormatting);
        }
        return package.ToArray();
    }

    /// <summary>
    /// 创建包含前向实体导入条件格式的测试工作簿。
    /// </summary>
    private static byte[] CreateForwardConditionalFormattingWorkbook(
        bool includeSuccessfulRow = false)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        SetHeader(sheet, "Code", "Quantity");
        var failureRow = includeSuccessfulRow ? 3 : 2;
        if (includeSuccessfulRow)
            SetLine(sheet, 2, "A", 2);
        SetLine(sheet, failureRow, "B", "wrong");
        sheet.Range($"B{failureRow}").AddConditionalFormat().WhenGreaterThan(0);

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含单个隐式动态字典的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateDynamicForwardWorkbook(bool unknownColumn = false)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("子项");
        sheet.Cell(1, 1).Value = "ParentId";
        sheet.Cell(1, 2).Value = "Name";
        sheet.Cell(1, 3).Value = "Zone";
        sheet.Cell(1, 4).Value = "Score";
        sheet.Cell(2, 1).Value = "P-1";
        sheet.Cell(2, 2).Value = "detail";
        sheet.Cell(2, 3).Value = "North";
        sheet.Cell(2, 4).Value = 12.5m;
        if (unknownColumn)
        {
            sheet.Cell(1, 5).Value = "Unknown";
            sheet.Cell(2, 5).Value = "unexpected";
        }
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建触发关系父键索引上限的 Forward 工作簿。
    /// </summary>
    /// <returns>工作簿字节内容。</returns>
    private static byte[] CreateRelationKeyLimitForwardWorkbook()
    {
        using var workbook = new XLWorkbook();
        var parents = workbook.AddWorksheet("父项");
        parents.Cell(1, 1).Value = "Id";
        parents.Cell(2, 1).Value = "P-1";
        parents.Cell(3, 1).Value = "P-2";
        var children = workbook.AddWorksheet("子项");
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
    /// 创建包含两个隐式动态字典列的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateMultipleImplicitDynamicForwardWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(1, 1).Value = "Code";
        sheet.Cell(1, 2).Value = "Zone";
        sheet.Cell(1, 3).Value = "Note";
        sheet.Cell(2, 1).Value = "A-1";
        sheet.Cell(2, 2).Value = "North";
        sheet.Cell(2, 3).Value = "Extended";
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含单个显式动态列组的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateExplicitDynamicForwardWorkbook(bool unknownColumn = false)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(1, 1).Value = "Code";
        sheet.Cell(1, 2).Value = "Zone";
        sheet.Cell(1, 3).Value = "Score";
        sheet.Cell(2, 1).Value = "A-1";
        sheet.Cell(2, 2).Value = "North";
        sheet.Cell(2, 3).Value = 12.5m;
        if (unknownColumn)
        {
            sheet.Cell(1, 4).Value = "Unknown";
            sheet.Cell(2, 4).Value = "unexpected";
        }
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含多个显式动态列组的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateMultipleExplicitDynamicForwardWorkbook(bool unknownColumn = false)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(1, 1).Value = "Code";
        sheet.Cell(1, 2).Value = "Zone";
        sheet.Cell(1, 3).Value = "Other";
        sheet.Cell(2, 1).Value = "A-1";
        sheet.Cell(2, 2).Value = "North";
        sheet.Cell(2, 3).Value = "Extended";
        if (unknownColumn)
        {
            sheet.Cell(1, 4).Value = "Unknown";
            sheet.Cell(2, 4).Value = "unexpected";
        }
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建只包含指定 Sheet 的 Forward 模板工作簿。
    /// </summary>
    private static byte[] CreateForwardTemplateWorkbook(string sheetName)
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet(sheetName);
        sheet.Cell(1, 1).Value = "Code";
        sheet.Cell(1, 2).Value = "Zone";
        sheet.Cell(1, 3).Value = "Score";
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含日期样式、数值和 Unicode 文本的 Forward 工作簿。
    /// </summary>
    private static byte[] CreateStyledForwardWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(1, 1).Value = "Code";
        sheet.Cell(1, 2).Value = "OccurredAt";
        sheet.Cell(1, 3).Value = "Amount";
        sheet.Cell(1, 4).Value = "Description";
        sheet.Cell(2, 1).Value = "商品😀";
        sheet.Cell(2, 2).Value = new DateTime(2024, 2, 29);
        sheet.Cell(2, 2).Style.DateFormat.Format = "yyyy-mm-dd";
        sheet.Cell(2, 3).Value = 1234.5m;
        sheet.Cell(2, 3).Style.NumberFormat.Format = "#,##0.00";
        sheet.Cell(2, 4).Value = "中文\n换行";
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建 Forward 路径使用的单列表实体布局。
    /// </summary>
    private static ExcelEntityLayout<GetterOnlyBatchDocument> CreateForwardLayout() =>
        ExcelEntity.Layout<GetterOnlyBatchDocument>(builder =>
            builder.ListRegion("明细", "A1", document => document.Lines));

    /// <summary>
    /// 创建摘要模式不需要访问结果集合的前向布局。
    /// </summary>
    private static ExcelEntityLayout<SummaryOnlyNoCollectionDocument>
        CreateSummaryOnlyNoCollectionLayout() =>
        ExcelEntity.Layout<SummaryOnlyNoCollectionDocument>(builder =>
            builder.ListRegion("明细", "A1", document => document.Lines));

    /// <summary>
    /// 创建包含唯一编码列的 Forward 实体布局。
    /// </summary>
    private static ExcelEntityLayout<UniqueForwardDocument> CreateUniqueForwardLayout() =>
        ExcelEntity.Layout<UniqueForwardDocument>(builder =>
            builder.ListRegion("明细", "A1", document => document.Lines));

    /// <summary>
    /// 创建包含两个唯一列列表区域的 Forward 实体布局。
    /// </summary>
    private static ExcelEntityLayout<UniqueForwardDocument> CreateMultipleSheetUniqueForwardLayout() =>
        ExcelEntity.Layout<UniqueForwardDocument>(builder => builder
            .ListRegion("明细", "A1", document => document.Lines)
            .ListRegion("税额", "A1", document => document.OtherLines));

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
    private static ExcelEntityLayout<MultipleExplicitDynamicDocument> CreateMultipleExplicitDynamicForwardLayout(
        ExcelUnknownDynamicValuePolicy unknownPolicy = ExcelUnknownDynamicValuePolicy.Ignore)
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
                    .DynamicColumnGroup("extensions", item => item.OtherFields, second)
                    .UnknownDynamicValues(unknownPolicy)));
    }

    /// <summary>
    /// 创建样式读取测试使用的实体布局。
    /// </summary>
    private static ExcelEntityLayout<StyledForwardDocument> CreateStyledForwardLayout() =>
        ExcelEntity.Layout<StyledForwardDocument>(builder =>
            builder.ListRegion("明细", "A1", document => document.Lines));

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
    /// 创建测试工作簿字节。
    /// </summary>
    /// <param name="invalidNumber">是否把固定数字单元格写成无效文本。</param>
    private static byte[] CreateWorkbook(bool invalidNumber = false, bool emptyLists = false,
        bool invalidQuantity = false)
    {
        using var workbook = new XLWorkbook();
        var header = workbook.AddWorksheet("单据");
        header.Cell("B2").Value = invalidNumber ? "错误" : 42;
        header.Cell("B3").Value = "采购单";

        var lines = workbook.AddWorksheet("明细");
        SetHeader(lines, "Code", "Quantity");
        if (!emptyLists)
        {
            SetLine(lines, 2, "A", 2);
            SetLine(lines, 3, "B", invalidQuantity ? "wrong" : "3");
            SetLine(lines, 4, "C", 4);
        }

        var taxes = workbook.AddWorksheet("税额");
        SetHeader(taxes, "Name", "Amount");
        if (!emptyLists)
        {
            SetTax(taxes, 2, "增值税", 10.5);
            SetTax(taxes, 3, "附加税", 2.5);
        }

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建两个列表区域均包含转换错误的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateMultipleSheetForwardErrorWorkbook()
    {
        using var workbook = new XLWorkbook();
        var lines = workbook.AddWorksheet("明细");
        SetHeader(lines, "Code", "Quantity");
        SetLine(lines, 2, "A-1", "wrong");
        var taxes = workbook.AddWorksheet("税额");
        SetHeader(taxes, "Name", "Amount");
        SetTax(taxes, 2, "增值税", 1.25);
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建两个列表区域均包含唯一编码的 Forward 工作簿。
    /// </summary>
    private static byte[] CreateMultipleSheetUniqueForwardWorkbook()
    {
        using var workbook = new XLWorkbook();
        var lines = workbook.AddWorksheet("明细");
        SetHeader(lines, "Code", "Quantity");
        SetLine(lines, 2, "A-1", 2);
        var taxes = workbook.AddWorksheet("税额");
        SetHeader(taxes, "Code", "Quantity");
        SetLine(taxes, 2, "T-1", 1);
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建第一列表包含重复编码、第二列表有效的 Forward 工作簿。
    /// </summary>
    private static byte[] CreateMultipleSheetUniqueErrorForwardWorkbook()
    {
        using var workbook = new XLWorkbook();
        var lines = workbook.AddWorksheet("明细");
        SetHeader(lines, "Code", "Quantity");
        SetLine(lines, 2, "A-1", 2);
        SetLine(lines, 3, "A-1", 3);
        var taxes = workbook.AddWorksheet("税额");
        SetHeader(taxes, "Code", "Quantity");
        SetLine(taxes, 2, "T-1", 1);
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建同一工作表承载两个非重叠列表区域的 XLSX 工作簿。
    /// </summary>
    private static byte[] CreateSameSheetForwardWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(1, 1).Value = "Code";
        sheet.Cell(1, 2).Value = "Quantity";
        sheet.Cell(1, 4).Value = "Code";
        sheet.Cell(1, 5).Value = "Quantity";
        sheet.Cell(2, 1).Value = "A-1";
        sheet.Cell(2, 2).Value = 2;
        sheet.Cell(2, 4).Value = "T-1";
        sheet.Cell(2, 5).Value = 1.25;
        sheet.Cell(3, 1).Value = "B-2";
        sheet.Cell(3, 2).Value = 3;
        sheet.Cell(3, 4).Value = "T-2";
        sheet.Cell(3, 5).Value = 2.5;
        using var output = new MemoryStream();
        workbook.SaveAs(output);
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
        using var input = new MemoryStream(CreateWorkbook());
        using var workbook = new XLWorkbook(input);
        var header = workbook.Worksheet("单据");
        if (includeNumber)
            workbook.DefinedNames.Add("OrderNumber", areaNumber
                ? "'单据'!$B$2:$C$2" : "'单据'!$B$2");
        workbook.DefinedNames.Add("OrderTitle", header.Range("B3:B3"));
        using var output = new MemoryStream();
        workbook.SaveAs(output);
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
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(3, 3).Value = "Code";
        sheet.Cell(3, 4).Value = "Quantity";
        sheet.Cell(4, 3).Value = "A-1";
        sheet.Cell(4, 4).Value = 2;
        sheet.Cell(5, 3).Value = "B-2";
        sheet.Cell(5, 4).Value = 3;
        if (includeAnchor)
            workbook.DefinedNames.Add("LinesStart", areaAnchor
                ? "'明细'!$C$3:$D$3" : "'明细'!$C$3");
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建同一工作表包含固定字段和两个列表区域的失败导入工作簿。
    /// </summary>
    private static byte[] CreateMultiRegionFailureWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(1, 1).Value = "Code";
        sheet.Cell(1, 2).Value = "Quantity";
        sheet.Cell(1, 4).Value = "预留";
        sheet.Cell(1, 5).Value = "预留金额";
        sheet.Cell(1, 7).Value = 42;
        sheet.Cell(2, 1).Value = "A";
        sheet.Cell(2, 2).Value = 2;
        sheet.Cell(3, 1).Value = "B";
        sheet.Cell(3, 2).Value = "wrong";
        sheet.Cell(4, 1).Value = "C";
        sheet.Cell(4, 2).Value = 4;
        sheet.Cell(5, 4).Value = "Name";
        sheet.Cell(5, 5).Value = "Amount";
        sheet.Cell(6, 4).Value = "增值税";
        sheet.Cell(6, 5).Value = 10.5;
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含分页小计、最终尾部和一条无效明细的工作簿。
    /// </summary>
    private static byte[] CreatePaginationFailureWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        SetHeader(sheet, "Code", "Quantity");
        SetLine(sheet, 2, "A", 2);
        SetLine(sheet, 3, "B", "wrong");
        sheet.Cell(4, 1).Value = "PAGE_TOTAL";
        sheet.Cell(4, 2).Value = "subtotal";
        SetLine(sheet, 5, "C", 4);
        sheet.Cell(6, 1).Value = "END";
        sheet.Cell(6, 2).Value = "total";
        sheet.Cell(1, 3).Value = "模板表头";
        sheet.Cell(1, 3).Style.Font.Bold = true;
        sheet.Range("C1:D1").Merge();
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 创建包含显式动态列转换错误的工作簿。
    /// </summary>
    private static byte[] CreateExplicitDynamicFailureWorkbook()
    {
        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("明细");
        sheet.Cell(1, 1).Value = "Code";
        sheet.Cell(1, 2).Value = "Zone";
        sheet.Cell(1, 3).Value = "Score";
        sheet.Cell(2, 1).Value = "A";
        sheet.Cell(2, 2).Value = "North";
        sheet.Cell(2, 3).Value = "wrong";
        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 写入列表表头。
    /// </summary>
    private static void SetHeader(IXLWorksheet sheet, string first, string second)
    {
        sheet.Cell(1, 1).Value = first;
        sheet.Cell(1, 2).Value = second;
    }

    /// <summary>
    /// 写入采购明细行。
    /// </summary>
    private static void SetLine(IXLWorksheet sheet, int row, string code, int quantity)
    {
        sheet.Cell(row, 1).Value = code;
        sheet.Cell(row, 2).Value = quantity;
    }

    /// <summary>
    /// 写入采购明细行并允许直接构造转换错误。
    /// </summary>
    private static void SetLine(IXLWorksheet sheet, int row, string code, string quantity)
    {
        sheet.Cell(row, 1).Value = code;
        sheet.Cell(row, 2).Value = quantity;
    }

    /// <summary>
    /// 写入税额明细行。
    /// </summary>
    private static void SetTax(IXLWorksheet sheet, int row, string name, double amount)
    {
        sheet.Cell(row, 1).Value = name;
        sheet.Cell(row, 2).Value = amount;
    }

    /// <summary>
    /// 保存实体布局批次测试根模型。
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
        public IDictionary<string, object> CustomFields { get; set; } = new Dictionary<string, object>();
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
        public IDictionary<string, object> Attributes { get; set; } = new Dictionary<string, object>();

        /// <summary>
        /// 获取或设置扩展属性字典。
        /// </summary>
        [DynamicColumn]
        public IDictionary<string, object> Extensions { get; set; } = new Dictionary<string, object>();
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
    /// 提供不可定位但可读取的测试流。
    /// </summary>
    private sealed class NonSeekableTrackingStream : Stream
    {
        private readonly MemoryStream _inner;

        internal NonSeekableTrackingStream(byte[] buffer)
        {
            _inner = new MemoryStream(buffer, writable: false);
        }

        public override bool CanRead => _inner.CanRead;
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
}
