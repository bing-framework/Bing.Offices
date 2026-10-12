using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.ClosedXml.Entities;
using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.IO;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.ClosedXml.Extensions;
using ClosedXML.Excel;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.Tests;

/// <summary>
/// 验证 ClosedXML Buffered 摘要模式组合布局的低保留行为。
/// </summary>
public sealed class EntityBufferedSummaryComplexLayoutTest
{
    /// <summary>
    /// 验证固定单元格、关系和动态列可以共同写出低保留失败工作簿。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Buffered_SummaryOnly_ShouldCombineComplexLayout(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(), writable: false);
        using var failure = new MemoryStream();
        var batches = new List<ExcelEntityImportBatch<ComplexDocument>>();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (batch, _) =>
                {
                    batches.Add(batch);
                    return Task.CompletedTask;
                })
            : importer.ImportEntityBatches(source, CreateLayout(), options, batches.Add);

        Assert.False(summary.IsSuccess);
        Assert.False(summary.HasCompleteResult);
        Assert.Contains(summary.Result.Errors, error =>
            error.Code == ExcelImportErrorCode.ValueConversion
            && error.SheetName == "单据"
            && error.RowIndex == 1
            && error.ColumnIndex == 1
            && error.PropertyName == nameof(ComplexDocument.Number));
        Assert.Contains(summary.Result.Errors, error =>
            error.Code == ExcelImportErrorCode.ValueConversion
            && error.SheetName == "子项"
            && error.RowIndex == 2
            && error.ColumnIndex == 4);
        Assert.Empty(summary.Result.Entity.Parents);
        Assert.Empty(summary.Result.Entity.Children);
        Assert.Contains(batches, batch => batch.IsDiagnostic);
        Assert.True(source.CanRead);

        failure.Position = 0;
        using var workbook = new XLWorkbook(failure);
        var document = workbook.Worksheet("单据");
        Assert.Equal("wrong", document.Cell("A1").GetString());

        var children = workbook.Worksheet("子项");
        Assert.Equal("ParentId", children.Cell("A1").GetString());
        Assert.Equal("Name", children.Cell("B1").GetString());
        Assert.Equal("Zone", children.Cell("C1").GetString());
        Assert.Equal("Score", children.Cell("D1").GetString());
        Assert.Equal("P-1", children.Cell("A2").GetString());
        Assert.Equal("North", children.Cell("C2").GetString());
        Assert.Equal("wrong", children.Cell("D2").GetString());
        Assert.NotNull(workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "_ImportErrors"));
    }

    /// <summary>
    /// 验证复杂布局达到错误上限后仍保持低保留结果。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Buffered_SummaryOnly_ComplexLayout_ShouldHonorMaxErrors(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(), writable: false);
        using var failure = new MemoryStream();
        var options = new ExcelEntityBatchImportOptions(1,
            CreateImportOptions(failure, new ExcelResourceLimits { MaxErrors = 1 }),
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.True(summary.Result.ErrorsTruncated);
        Assert.Single(summary.Result.Errors);
        Assert.Equal(1, summary.Result.MaxErrors);
        Assert.Empty(summary.Result.Entity.Parents);
        Assert.Empty(summary.Result.Entity.Children);
        Assert.True(source.CanRead);

        failure.Position = 0;
        using var workbook = new XLWorkbook(failure);
        Assert.NotNull(workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "_ImportErrors"));
    }

    /// <summary>
    /// 验证复杂布局从不可定位输入生成低保留失败工作簿。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Buffered_SummaryOnly_ComplexLayout_NonSeekableInput(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new NonSeekableReadStream(CreateWorkbook());
        using var failure = new MemoryStream();
        var options = new ExcelEntityBatchImportOptions(1, CreateImportOptions(failure),
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.False(summary.HasCompleteResult);
        Assert.NotEmpty(summary.Result.Errors);
        Assert.Empty(summary.Result.Entity.Parents);
        Assert.Empty(summary.Result.Entity.Children);
        Assert.True(source.CanRead);

        failure.Position = 0;
        using var workbook = new XLWorkbook(failure);
        Assert.NotNull(workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "单据"));
        Assert.NotNull(workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "子项"));
        Assert.NotNull(workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "_ImportErrors"));
    }

    /// <summary>
    /// 验证复杂布局的失败工作簿提交异常不替换原目标。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Buffered_SummaryOnly_ComplexLayout_CommitFailure(
        bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = new MemoryStream(CreateWorkbook(), writable: false);
        var path = Path.Combine(directory.Path, "complex-failure.xlsx");
        var sentinel = new byte[] { 7, 6, 5 };
        File.WriteAllBytes(path, sentinel);
        var expected = new BingOfficesFileCommitException("test", new IOException("locked"),
            "FailureWorkbook", BingOfficesStage.Commit);
        var committer = new CommitProbe { Failure = expected };
        var importer = new ClosedXmlEntityBatchImporter(new ClosedXmlExcelImporter(null, null,
            null, null, null, committer));
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                DestinationPath = path,
                TemporaryDirectory = directory.Path
            });
        var batchOptions = new ExcelEntityBatchImportOptions(1, options,
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

        BingOfficesFileCommitException actual;
        if (asynchronous)
        {
            actual = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateLayout(), batchOptions,
                    (_, _) => Task.CompletedTask));
        }
        else
        {
            actual = Assert.Throws<BingOfficesFileCommitException>(() =>
                importer.ImportEntityBatches(source, CreateLayout(), batchOptions, _ => { }));
        }

        Assert.Same(expected, actual);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证复杂布局的失败工作簿写入中取消时保留原目标并清理暂存文件。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Buffered_SummaryOnly_ComplexLayout_CancellationDuringCommit(
        bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = new MemoryStream(CreateWorkbook(), writable: false);
        using var cancellation = new CancellationTokenSource();
        var path = Path.Combine(directory.Path, "complex-cancel.xlsx");
        var sentinel = new byte[] { 8, 7, 6 };
        File.WriteAllBytes(path, sentinel);
        var committer = new CommitProbe { AfterWrite = cancellation.Cancel };
        var importer = new ClosedXmlEntityBatchImporter(new ClosedXmlExcelImporter(null, null,
            null, null, null, committer));
        var options = new ExcelEntityImportOptions(null, ExcelValidationFailureMode.Continue,
            ExcelImportValidationMode.ConfiguredRules, ExcelUnsupportedFeaturePolicy.Fail,
            new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                DestinationPath = path,
                TemporaryDirectory = directory.Path
            });
        var batchOptions = new ExcelEntityBatchImportOptions(1, options,
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

        OperationCanceledException actual;
        if (asynchronous)
        {
            actual = await Assert.ThrowsAsync<OperationCanceledException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateLayout(), batchOptions,
                    (_, _) => Task.CompletedTask, cancellation.Token));
        }
        else
        {
            actual = Assert.Throws<OperationCanceledException>(() =>
                importer.ImportEntityBatches(source, CreateLayout(), batchOptions, _ => { },
                    cancellation.Token));
        }

        Assert.Equal(cancellation.Token, actual.CancellationToken);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证复杂布局关系键预算耗尽时仍写出结构化失败工作簿。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Buffered_SummaryOnly_ComplexLayout_RelationKeyLimit(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(includeSecondParent: true),
            writable: false);
        using var failure = new MemoryStream();
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(new ExcelResourceLimits { MaxTrackedRelationKeys = 1 },
                ExcelValidationFailureMode.Continue, ExcelImportValidationMode.ConfiguredRules,
                ExcelUnsupportedFeaturePolicy.Fail, new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                    Destination = failure
                }), ExcelEntityBatchParsingMode.Buffered,
            ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.Contains(summary.Result.Errors, error =>
            error.Code == ExcelImportErrorCode.ResourceLimit);
        Assert.False(summary.HasCompleteResult);
        Assert.Empty(summary.Result.Entity.Parents);
        Assert.Empty(summary.Result.Entity.Children);
        Assert.True(source.CanRead);

        failure.Position = 0;
        using var workbook = new XLWorkbook(failure);
        Assert.NotNull(workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "父项"));
        Assert.NotNull(workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "_ImportErrors"));
    }

    /// <summary>
    /// 验证复杂布局关系键预算失败时提交异常不替换原目标。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Buffered_SummaryOnly_ComplexLayout_RelationKeyLimit_CommitFailure(
        bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = new MemoryStream(CreateWorkbook(includeSecondParent: true),
            writable: false);
        var path = Path.Combine(directory.Path, "complex-relation-limit-failure.xlsx");
        var sentinel = new byte[] { 9, 8, 7 };
        File.WriteAllBytes(path, sentinel);
        var expected = new BingOfficesFileCommitException("test", new IOException("locked"),
            "FailureWorkbook", BingOfficesStage.Commit);
        var committer = new CommitProbe { Failure = expected };
        var importer = new ClosedXmlEntityBatchImporter(new ClosedXmlExcelImporter(null, null,
            null, null, null, committer));
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(new ExcelResourceLimits { MaxTrackedRelationKeys = 1 },
                ExcelValidationFailureMode.Continue, ExcelImportValidationMode.ConfiguredRules,
                ExcelUnsupportedFeaturePolicy.Fail, new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                    DestinationPath = path,
                    TemporaryDirectory = directory.Path
                }), ExcelEntityBatchParsingMode.Buffered,
            ExcelEntityBatchResultMode.SummaryOnly);

        BingOfficesFileCommitException actual;
        if (asynchronous)
        {
            actual = await Assert.ThrowsAsync<BingOfficesFileCommitException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                    (_, _) => Task.CompletedTask));
        }
        else
        {
            actual = Assert.Throws<BingOfficesFileCommitException>(() =>
                importer.ImportEntityBatches(source, CreateLayout(), options, _ => { }));
        }

        Assert.Same(expected, actual);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证复杂布局关系键预算失败时写入中取消保留原目标并清理暂存文件。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Buffered_SummaryOnly_ComplexLayout_RelationKeyLimit_CancellationDuringCommit(
        bool asynchronous)
    {
        using var directory = new TestDirectory();
        using var source = new MemoryStream(CreateWorkbook(includeSecondParent: true),
            writable: false);
        using var cancellation = new CancellationTokenSource();
        var path = Path.Combine(directory.Path, "complex-relation-limit-cancel.xlsx");
        var sentinel = new byte[] { 6, 5, 4 };
        File.WriteAllBytes(path, sentinel);
        var committer = new CommitProbe { AfterWrite = cancellation.Cancel };
        var importer = new ClosedXmlEntityBatchImporter(new ClosedXmlExcelImporter(null, null,
            null, null, null, committer));
        var options = new ExcelEntityBatchImportOptions(1,
            new ExcelEntityImportOptions(new ExcelResourceLimits { MaxTrackedRelationKeys = 1 },
                ExcelValidationFailureMode.Continue, ExcelImportValidationMode.ConfiguredRules,
                ExcelUnsupportedFeaturePolicy.Fail, new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                    DestinationPath = path,
                    TemporaryDirectory = directory.Path
                }), ExcelEntityBatchParsingMode.Buffered,
            ExcelEntityBatchResultMode.SummaryOnly);

        OperationCanceledException actual;
        if (asynchronous)
        {
            actual = await Assert.ThrowsAsync<OperationCanceledException>(() =>
                importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                    (_, _) => Task.CompletedTask, cancellation.Token));
        }
        else
        {
            actual = Assert.Throws<OperationCanceledException>(() =>
                importer.ImportEntityBatches(source, CreateLayout(), options, _ => { },
                    cancellation.Token));
        }

        Assert.Equal(cancellation.Token, actual.CancellationToken);
        Assert.Equal(sentinel, File.ReadAllBytes(path));
        Assert.Equal(new[] { path }, Directory.GetFiles(directory.Path));
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证复杂布局唯一值预算耗尽时仍写出低保留失败工作簿。
    /// </summary>
    /// <param name="asynchronous">是否使用异步入口。</param>
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImportEntityBatches_Buffered_SummaryOnly_ComplexLayout_UniqueValueLimit(
        bool asynchronous)
    {
        using var provider = new ServiceCollection().AddBingOfficesClosedXml()
            .BuildServiceProvider();
        var importer = provider.GetRequiredService<IExcelEntityBatchImporter>();
        using var source = new MemoryStream(CreateWorkbook(includeSecondParent: true,
            invalidDynamicValue: false), writable: false);
        using var failure = new MemoryStream();
        var options = new ExcelEntityBatchImportOptions(1,
            CreateImportOptions(failure, new ExcelResourceLimits { MaxTrackedUniqueValues = 1 }),
            ExcelEntityBatchParsingMode.Buffered, ExcelEntityBatchResultMode.SummaryOnly);

        var summary = asynchronous
            ? await importer.ImportEntityBatchesAsync(source, CreateLayout(), options,
                (_, _) => Task.CompletedTask)
            : importer.ImportEntityBatches(source, CreateLayout(), options, _ => { });

        Assert.False(summary.IsSuccess);
        Assert.Contains(summary.Result.Errors, error =>
            error.Code == ExcelImportErrorCode.ResourceLimit
            && error.SheetName == "子项"
            && error.PropertyName == nameof(ComplexChild.Name)
            && error.RowIndex == 3);
        Assert.False(summary.HasCompleteResult);
        Assert.Empty(summary.Result.Entity.Parents);
        Assert.Empty(summary.Result.Entity.Children);
        Assert.True(source.CanRead);

        failure.Position = 0;
        using var workbook = new XLWorkbook(failure);
        Assert.NotNull(workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "子项"));
        Assert.NotNull(workbook.Worksheets.FirstOrDefault(sheet => sheet.Name == "_ImportErrors"));
    }

    /// <summary>
    /// 创建只输出错误行的实体导入选项。
    /// </summary>
    /// <param name="destination">失败工作簿目标流。</param>
    /// <param name="resourceLimits">资源限制。</param>
    /// <returns>实体导入选项。</returns>
    private static ExcelEntityImportOptions CreateImportOptions(Stream destination,
        ExcelResourceLimits resourceLimits = null) =>
        new(resourceLimits, ExcelValidationFailureMode.Continue, ExcelImportValidationMode.ConfiguredRules,
            ExcelUnsupportedFeaturePolicy.Fail, new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.ErrorRowsOnly,
                Destination = destination
            });

    /// <summary>
    /// 创建包含固定单元格、关系和动态列的实体布局。
    /// </summary>
    /// <returns>组合实体布局。</returns>
    private static ExcelEntityLayout<ComplexDocument> CreateLayout()
    {
        var definitions = new[]
        {
            new ExcelDynamicColumnDefinition
            {
                Key = "zoneValue",
                Title = "Zone",
                DataType = typeof(string),
                PhysicalColumnIndex = 2
            },
            new ExcelDynamicColumnDefinition
            {
                Key = "scoreValue",
                Title = "Score",
                DataType = typeof(decimal),
                PhysicalColumnIndex = 3
            }
        };
        return ExcelEntity.Layout<ComplexDocument>(builder => builder
            .Cell("单据", "A1", document => document.Number)
            .ListRegion("父项", "A1", document => document.Parents)
            .ListRegion("子项", "A1", document => document.Children,
                region => region.DynamicColumnGroup("attributes", item => item.Fields,
                    definitions))
            .HasMany(document => document.Parents, document => document.Children,
                parent => parent.Id, child => child.ParentId,
                parent => parent.Children, StringComparer.OrdinalIgnoreCase));
    }

    /// <summary>
    /// 创建固定字段和动态列均包含错误的 XLSX 工作簿。
    /// </summary>
    /// <param name="includeSecondParent">是否追加第二个父键以触发关系键预算。</param>
    /// <returns>工作簿字节。</returns>
    private static byte[] CreateWorkbook(bool includeSecondParent = false,
        bool invalidDynamicValue = true)
    {
        using var workbook = new XLWorkbook();
        var document = workbook.AddWorksheet("单据");
        document.Cell("A1").Value = "wrong";

        var parents = workbook.AddWorksheet("父项");
        parents.Cell("A1").Value = "Id";
        parents.Cell("A2").Value = "P-1";
        if (includeSecondParent)
            parents.Cell("A3").Value = "P-2";

        var children = workbook.AddWorksheet("子项");
        children.Cell("A1").Value = "ParentId";
        children.Cell("B1").Value = "Name";
        children.Cell("C1").Value = "Zone";
        children.Cell("D1").Value = "Score";
        children.Cell("A2").Value = "P-1";
        children.Cell("B2").Value = "detail";
        children.Cell("C2").Value = "North";
        children.Cell("D2").Value = invalidDynamicValue ? "wrong" : "9";
        if (includeSecondParent)
        {
            children.Cell("A3").Value = "P-2";
            children.Cell("B3").Value = "other";
            children.Cell("C3").Value = "South";
            children.Cell("D3").Value = 9;
        }

        using var output = new MemoryStream();
        workbook.SaveAs(output);
        return output.ToArray();
    }

    /// <summary>
    /// 提供不可定位但可读的工作簿输入流。
    /// </summary>
    private sealed class NonSeekableReadStream : Stream
    {
        /// <summary>
        /// 保存实际读取流。
        /// </summary>
        private readonly MemoryStream _inner;

        /// <summary>
        /// 初始化一个 <see cref="NonSeekableReadStream" /> 类型的实例。
        /// </summary>
        /// <param name="bytes">输入内容。</param>
        public NonSeekableReadStream(byte[] bytes)
        {
            _inner = new MemoryStream(bytes, writable: false);
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
        public override int Read(byte[] buffer, int offset, int count) =>
            _inner.Read(buffer, offset, count);

        /// <inheritdoc />
        public override int Read(Span<byte> buffer) => _inner.Read(buffer);

        /// <inheritdoc />
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken) => _inner.ReadAsync(buffer, offset, count,
            cancellationToken);

        /// <inheritdoc />
        public override ValueTask<int> ReadAsync(Memory<byte> buffer,
            CancellationToken cancellationToken = default) => _inner.ReadAsync(buffer,
            cancellationToken);

        /// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) =>
            throw new NotSupportedException();

        /// <inheritdoc />
        public override void SetLength(long value) => throw new NotSupportedException();

        /// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) =>
            throw new NotSupportedException();

        /// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _inner.Dispose();
            base.Dispose(disposing);
        }
    }

    /// <summary>
    /// 记录提交调用并委托默认原子提交实现。
    /// </summary>
    private sealed class CommitProbe : IFileExportCommitter
    {
        /// <summary>
        /// 默认文件提交实现。
        /// </summary>
        private readonly DefaultFileExportCommitter _inner = new();

        /// <summary>
        /// 获取或设置提交前抛出的异常。
        /// </summary>
        public Exception Failure { get; set; }

        /// <summary>
        /// 获取或设置写入完成后的回调。
        /// </summary>
        public Action AfterWrite { get; set; }

        /// <inheritdoc />
        public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken,
            string format)
        {
            ThrowIfConfigured();
            _inner.Commit(path, stream =>
            {
                write(stream);
                AfterWrite?.Invoke();
            }, cancellationToken, format);
        }

        /// <inheritdoc />
        public async Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
            CancellationToken cancellationToken, string format)
        {
            ThrowIfConfigured();
            await _inner.CommitAsync(path, async (stream, token) =>
            {
                await writeAsync(stream, token);
                AfterWrite?.Invoke();
            }, cancellationToken, format).ConfigureAwait(false);
        }

        /// <summary>
        /// 按测试配置抛出提交异常。
        /// </summary>
        private void ThrowIfConfigured()
        {
            if (Failure != null)
                throw Failure;
        }
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
            "closedxml-complex-layout-" + Guid.NewGuid().ToString("N"));

        /// <summary>
        /// 初始化一个 <see cref="TestDirectory" /> 类型的实例。
        /// </summary>
        public TestDirectory() => Directory.CreateDirectory(Path);

        /// <inheritdoc />
        public void Dispose() => Directory.Delete(Path, true);
    }

    /// <summary>
    /// 保存组合布局测试根对象。
    /// </summary>
    private sealed class ComplexDocument
    {
        /// <summary>
        /// 获取或设置单据编号。
        /// </summary>
        public int Number { get; set; }

        /// <summary>
        /// 获取父项集合。
        /// </summary>
        public List<ComplexParent> Parents { get; } = new();

        /// <summary>
        /// 获取子项集合。
        /// </summary>
        public List<ComplexChild> Children { get; } = new();
    }

    /// <summary>
    /// 保存组合布局测试父项。
    /// </summary>
    private sealed class ComplexParent
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 获取关联子项集合。
        /// </summary>
        public List<ComplexChild> Children { get; } = new();
    }

    /// <summary>
    /// 保存组合布局测试子项。
    /// </summary>
    private sealed class ComplexChild
    {
        /// <summary>
        /// 获取或设置父项标识。
        /// </summary>
        public string ParentId { get; set; }

        /// <summary>
        /// 获取或设置子项名称。
        /// </summary>
        [ExcelUnique]
        public string Name { get; set; }

        /// <summary>
        /// 获取动态字段字典。
        /// </summary>
        public Dictionary<string, object> Fields { get; } = new();
    }
}
