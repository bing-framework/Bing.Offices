using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Bing.Offices.Attributes;
using Bing.Offices.ExcelDataReader;
using Bing.Offices.ExcelDataReader.Extensions;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Microsoft.Extensions.DependencyInjection;
using Xunit;

namespace Bing.Offices.ExcelDataReader.Tests;

/// <summary>
/// ExcelDataReader Provider 的职责级测试。
/// </summary>
public sealed class ExcelDataReaderProviderTest
{
    /// <summary>
    /// 验证 Provider 声明 XLSB 只读和分批导入能力。
    /// </summary>
    [Fact]
    public void Capabilities_ShouldDescribeReadOnlyXlsbAndBatchSupport()
    {
        var importer = new ExcelDataReaderExcelImporter();

        Assert.Contains(ExcelFormat.Xlsb, importer.ReadFormats);
        Assert.Empty(importer.WriteFormats);
        Assert.True(importer.SupportsBatchImport);
        Assert.False(importer.SupportsCompleteWorkbookExport);
    }

    /// <summary>
    /// 真实 XLSB 文件应完成固定列完整导入和分批导入。
    /// </summary>
    [Fact]
    public void Import_ShouldReadRealXlsbFixtureAndDeliverBatches()
    {
        using var source = OpenFixture("xlsb/Issue635.xlsb");
        var importer = new ExcelDataReaderExcelImporter();
        var request = ExcelImport.Workbook<XlsbWorkbook>(workbook => workbook
            .Sheet<XlsbRow>("Лист2", root => root.Rows));

        var result = importer.Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var row = Assert.Single(result.Workbook.Rows);
        Assert.Equal("123.456", row.A);
        Assert.False(row.B);
        Assert.Equal(new DateTime(2023, 6, 9), row.C);
        Assert.Equal("Plaintext", row.D);
        Assert.True(source.CanRead);

        source.Position = 0;
        var batches = new List<ExcelImportBatch<XlsbRow>>();
        var summary = importer.ImportBatches(source, new ExcelBatchImportRequest<XlsbRow>("Лист2"),
            batches.Add);
        Assert.True(summary.IsSuccess);
        Assert.Equal("123.456", Assert.Single(Assert.Single(batches).Items).A);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证 XLSB 物理单元格预算在预检阶段被拒绝。
    /// </summary>
    [Fact]
    public void Import_XlsbPhysicalCellLimits_ShouldBeExplicitlyUnsupported()
    {
        using var source = OpenFixture("xlsb/Issue635.xlsb");
        var exception = Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ExcelDataReaderExcelImporter().Import(source,
                ExcelImport.Workbook<XlsbWorkbook>(workbook => workbook
                    .ResourceLimits(new ExcelResourceLimits { MaxCells = 1 })
                    .Sheet<XlsbRow>("Лист2", root => root.Rows))));

        Assert.Equal("ExcelDataReader", exception.Provider);
        Assert.Equal(BingOfficesStage.Preflight, exception.Stage);
        Assert.Contains("物理单元格", exception.Message);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证导入 NPOI 生成的 XLSX 并映射人员数据。
    /// </summary>
    [Fact]
    public void Import_ShouldReadNpoiGeneratedXlsxAndMapValues()
    {
        using var source = CreateWorkbook(ExcelFormat.Xlsx);
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "Alice", "Bob" }, result.Workbook.People.Select(item => item.Name));
        Assert.Equal(new[] { 30, 31 }, result.Workbook.People.Select(item => item.Age));
    }

    /// <summary>
    /// 验证导入 NPOI 生成的 XLS 人员数据。
    /// </summary>
    [Fact]
    public void Import_ShouldReadNpoiGeneratedXls()
    {
        using var source = CreateWorkbook(ExcelFormat.Xls);
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new[] { "Alice", "Bob" }, result.Workbook.People.Select(item => item.Name));
    }

    /// <summary>
    /// 固定列导入应保留中文、日期和空值。
    /// </summary>
    [Fact]
    public void Import_ShouldPreserveUnicodeDateAndEmptyValue()
    {
        var request = ExcelExport.Workbook(workbook => workbook
            .AddSheet("Data", new[]
            {
                new RichPerson { Name = "张三", RegisteredAt = new DateTime(2026, 9, 25), Note = null }
            }));
        using var source = new MemoryStream();
        new Bing.Offices.Exports.NpoiExcelExporter().Export(request, source);
        source.Position = 0;

        var importRequest = ExcelImport.Workbook<RichImportWorkbook>(workbook =>
            workbook.Sheet<RichPerson>("Data", root => root.Items));
        var result = new ExcelDataReaderExcelImporter().Import(source, importRequest);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        var item = Assert.Single(result.Workbook.Items);
        Assert.Equal("张三", item.Name);
        Assert.Equal(new DateTime(2026, 9, 25), item.RegisteredAt);
        Assert.Null(item.Note);
    }

    /// <summary>
    /// 1900 与 1904 日期系统都应还原为同一日期值。
    /// </summary>
    /// <param name="relativePath">相对于黄金样例目录的文件路径。</param>
    [Theory]
    [InlineData("date/date1900.xlsx")]
    [InlineData("date/date1904.xlsx")]
    public void Import_ShouldRespectWorkbookDateSystem(string relativePath)
    {
        using var source = OpenFixture(relativePath);
        var request = ExcelImport.Workbook<DateImportWorkbook>(workbook =>
            workbook.Sheet<DateValue>("Data", root => root.Items));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(new DateTime(2026, 9, 4), Assert.Single(result.Workbook.Items).Date);
    }

    /// <summary>
    /// 公式字段按 ExcelDataReader 提供的缓存值读取，不尝试重新计算。
    /// </summary>
    [Fact]
    public void Import_ShouldReadCachedFormulaValue()
    {
        using var source = OpenFixture("formula/cached-number.xlsx");
        var request = ExcelImport.Workbook<FormulaImportWorkbook>(workbook =>
            workbook.Sheet<FormulaValue>("Formula", root => root.Items));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.True(result.IsSuccess, string.Join(";", result.Errors.Select(error => error.Message)));
        Assert.Equal(2, Assert.Single(result.Workbook.Items).Value);
    }

    /// <summary>
    /// 验证工作簿行数超限时不返回部分实体。
    /// </summary>
    [Fact]
    public void Import_ShouldApplyWorkbookRowBudgetWithoutPartialEntities()
    {
        using var source = CreateWorkbook();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxRows = 1 })
            .Sheet<Person>("People", root => root.People));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.People);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);
    }

    /// <summary>
    /// 验证跨表共享行数预算且超限时清除先前导入实体。
    /// </summary>
    [Fact]
    public void Import_ShouldShareRowBudgetAcrossSheetsAndHideEarlierEntities()
    {
        using var source = CreateWorkbook(ExcelFormat.Xlsx, includeSecondSheet: true);
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .ResourceLimits(new ExcelResourceLimits { MaxRows = 2 })
            .Sheet<Person>("People", root => root.People)
            .Sheet<Person>("Other", root => root.Other));

        var result = new ExcelDataReaderExcelImporter().Import(source, request);

        Assert.False(result.IsSuccess);
        Assert.Empty(result.Workbook.People);
        Assert.Empty(result.Workbook.Other);
        Assert.Contains(result.Errors, error => error.Code == ExcelImportErrorCode.ResourceLimit);
    }

    /// <summary>
    /// 验证分批导入交付完整批次且保持源流可用。
    /// </summary>
    [Fact]
    public void ImportBatches_ShouldDeliverBatchesAndKeepSourceOwnedByCaller()
    {
        using var source = CreateWorkbook();
        var batches = new List<ExcelImportBatch<Person>>();
        var request = new ExcelBatchImportRequest<Person>("People") { BatchSize = 1 };

        var summary = new ExcelDataReaderExcelImporter().ImportBatches(source, request, batches.Add);

        Assert.True(summary.IsSuccess);
        Assert.Equal(2, summary.SucceededRows);
        Assert.Equal(2, batches.SelectMany(batch => batch.Items).Count());
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证异步导入等待每个批次回调完成。
    /// </summary>
    [Fact]
    public async Task ImportBatchesAsync_ShouldAwaitEachBatchAndHonorCancellation()
    {
        using var source = CreateWorkbook();
        var count = 0;
        var request = new ExcelBatchImportRequest<Person>("People") { BatchSize = 1 };

        var summary = await new ExcelDataReaderExcelImporter().ImportBatchesAsync(source, request,
            async (batch, cancellationToken) =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                count += batch.Items.Count;
                await Task.Yield();
            });

        Assert.True(summary.IsSuccess);
        Assert.Equal(2, count);
    }

    /// <summary>
    /// 批次大小应稳定分组，行数预算只允许交付预算内的完整实体。
    /// </summary>
    /// <param name="rowCount">生成的测试数据行数。</param>
    /// <param name="expectedResourceLimitExceeded">是否预期超出工作簿行数预算。</param>
    [Theory]
    [InlineData(4, false)]
    [InlineData(5, true)]
    public void ImportBatches_BatchSizeAndMaxRowsBoundary_ShouldDeliverOnlyBudgetedRows(
        int rowCount, bool expectedResourceLimitExceeded)
    {
        var rows = Enumerable.Range(1, rowCount)
            .Select(index => new Person { Name = $"Person-{index}", Age = 20 + index })
            .ToArray();
        using var source = CreateWorkbook(people: rows);
        var batches = new List<ExcelImportBatch<Person>>();
        var request = new ExcelBatchImportRequest<Person>("People")
        {
            BatchSize = 2,
            ResourceLimits = new ExcelResourceLimits { MaxRows = 4 }
        };

        var summary = new ExcelDataReaderExcelImporter().ImportBatches(source, request, batches.Add);

        Assert.Equal(new[] { 2, 2 }, batches.Select(batch => batch.Items.Count));
        Assert.Equal(new[] { "Person-1", "Person-2", "Person-3", "Person-4" },
            batches.SelectMany(batch => batch.Items).Select(item => item.Name));
        Assert.All(batches, batch => Assert.Empty(batch.Errors));
        Assert.Equal(4, summary.RowsRead);
        Assert.Equal(4, summary.SucceededRows);
        Assert.Equal(0, summary.FailedRows);
        Assert.Equal(expectedResourceLimitExceeded, summary.ResourceLimitExceeded);
        Assert.Equal(!expectedResourceLimitExceeded, summary.IsSuccess);
    }

    /// <summary>
    /// Continue 模式应交付错误批次，并继续交付错误行之后的有效实体。
    /// </summary>
    [Fact]
    public void ImportBatches_ContinueMode_ShouldReturnErrorRowWithItsBatch()
    {
        using var source = CreateWorkbook(people: new[]
        {
            new Person { Name = "Alice", Age = 30 },
            new Person { Name = string.Empty, Age = 31 },
            new Person { Name = "Bob", Age = 32 }
        });
        var batches = new List<ExcelImportBatch<Person>>();
        var request = new ExcelBatchImportRequest<Person>("People")
        {
            BatchSize = 1,
            ValidationFailureMode = ExcelValidationFailureMode.Continue
        };

        var summary = new ExcelDataReaderExcelImporter().ImportBatches(source, request, batches.Add);

        Assert.Equal(3, batches.Count);
        Assert.Equal("Alice", Assert.Single(batches[0].Items).Name);
        Assert.Empty(batches[0].Errors);
        Assert.Empty(batches[1].Items);
        var error = Assert.Single(batches[1].Errors);
        Assert.Equal(ExcelImportErrorCode.Validation, error.Code);
        Assert.Equal(3, error.RowIndex);
        Assert.Equal(nameof(Person.Name), error.PropertyName);
        Assert.Equal("Bob", Assert.Single(batches[2].Items).Name);
        Assert.Empty(batches[2].Errors);
        Assert.Equal(3, summary.RowsRead);
        Assert.Equal(2, summary.SucceededRows);
        Assert.Equal(1, summary.FailedRows);
        Assert.Equal(1, summary.ErrorCount);
    }

    /// <summary>
    /// 异步导入在真实输入读取后取消时，不得释放调用方拥有的源流。
    /// </summary>
    [Fact]
    public async Task ImportAsync_MidReadCancellation_ShouldKeepCallerSourceOpen()
    {
        using var generated = CreateWorkbook();
        using var cancellation = new CancellationTokenSource();
        using var source = new CancelAfterFirstReadStream(generated, cancellation);
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new ExcelDataReaderExcelImporter().ImportAsync(source, request, cancellation.Token));

        Assert.True(source.HasRead);
        Assert.True(cancellation.IsCancellationRequested);
        Assert.False(source.WasDisposed);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 损坏输入应转换为带 Provider 和阶段元数据的公共异常。
    /// </summary>
    [Fact]
    public void Import_MalformedWorkbook_ShouldReturnStructuredProviderException()
    {
        using var source = new MemoryStream(new byte[] { 0x42, 0x49, 0x4E, 0x47, 0x00, 0x01 });
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook =>
            workbook.Sheet<Person>("People", root => root.People));

        var exception = Assert.Throws<BingOfficesImportException>(() =>
            new ExcelDataReaderExcelImporter().Import(source, request));

        Assert.Equal("ExcelDataReader", exception.Provider);
        Assert.Equal(BingOfficesOperation.Import, exception.Operation);
        Assert.Equal(BingOfficesStage.Read, exception.Stage);
        Assert.Equal(BingOfficesErrorCode.ImportFailed, exception.Code);
        Assert.NotNull(exception.InnerException);
        Assert.True(source.CanRead);
    }

    /// <summary>
    /// 验证不支持的失败工作簿请求不会写入目标流。
    /// </summary>
    [Fact]
    public void Import_ShouldRejectFailureWorkbookBeforeCreatingOutput()
    {
        using var source = CreateWorkbook();
        using var destination = new MemoryStream();
        var request = ExcelImport.Workbook<ImportWorkbook>(workbook => workbook
            .FailureWorkbook(new ExcelImportFailureOptions
            {
                Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                Destination = destination
            })
            .Sheet<Person>("People", root => root.People));

        Assert.Throws<BingOfficesUnsupportedFeatureException>(() =>
            new ExcelDataReaderExcelImporter().Import(source, request));
        Assert.Equal(0, destination.Length);
    }

    /// <summary>
    /// 验证注册完整导入和分批导入服务。
    /// </summary>
    [Fact]
    public void AddBingOfficesExcelDataReader_ShouldRegisterReadAndBatchServices()
    {
        using var provider = new ServiceCollection()
            .AddBingOfficesExcelDataReader()
            .BuildServiceProvider();

        Assert.IsType<ExcelDataReaderExcelImporter>(provider.GetRequiredService<IExcelImporter>());
        Assert.IsType<ExcelDataReaderExcelImporter>(provider.GetRequiredService<IExcelBatchImporter>());
    }

    /// <summary>
    /// 创建用于完整导入和分批导入的人员工作簿。
    /// </summary>
    /// <param name="format">生成的工作簿格式。</param>
    /// <param name="includeSecondSheet">是否生成第二个工作表。</param>
    /// <param name="people">主要工作表的数据；null 时使用默认人员数据。</param>
    /// <returns>包含测试工作簿内容且位于起始位置的流。</returns>
    private static MemoryStream CreateWorkbook(ExcelFormat format = ExcelFormat.Xlsx,
        bool includeSecondSheet = false, IEnumerable<Person> people = null)
    {
        var request = ExcelExport.Workbook(workbook =>
        {
            workbook.Format(format).AddSheet("People", people ?? new[]
            {
                new Person { Name = "Alice", Age = 30 },
                new Person { Name = "Bob", Age = 31 }
            });
            if (includeSecondSheet)
                workbook.AddSheet("Other", new[]
                {
                    new Person { Name = "Carol", Age = 32 },
                    new Person { Name = "Dan", Age = 33 }
                });
        });
        var stream = new MemoryStream();
        new Bing.Offices.Exports.NpoiExcelExporter().Export(request, stream);
        stream.Position = 0;
        return stream;
    }

    /// <summary>
    /// 打开导入测试使用的黄金样例文件。
    /// </summary>
    /// <param name="relativePath">相对于黄金样例目录的文件路径。</param>
    /// <returns>已打开的只读样例文件流。</returns>
    private static FileStream OpenFixture(string relativePath)
        => File.OpenRead(Path.Combine(AppContext.BaseDirectory, "Resources", "Golden", relativePath));

    /// <summary>
    /// 承载人员导入结果的测试工作簿。
    /// </summary>
    private sealed class ImportWorkbook
    {
        /// <summary>
        /// 获取主要工作表导入的人员集合。
        /// </summary>
        public List<Person> People { get; } = new List<Person>();
        /// <summary>
        /// 获取第二个工作表导入的人员集合。
        /// </summary>
        public List<Person> Other { get; } = new List<Person>();
    }

    /// <summary>
    /// 用于必填姓名映射的人员测试模型。
    /// </summary>
    private sealed class Person
    {
        /// <summary>
        /// 获取或设置测试人员姓名。
        /// </summary>
        [ExcelRequired]
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置测试人员年龄。
        /// </summary>
        public int Age { get; set; }
    }

    /// <summary>
    /// 承载文本、日期和空值导入结果的测试工作簿。
    /// </summary>
    private sealed class RichImportWorkbook
    {
        /// <summary>
        /// 获取文本、日期和空值导入结果集合。
        /// </summary>
        public List<RichPerson> Items { get; } = new List<RichPerson>();
    }

    /// <summary>
    /// 用于文本、日期和空值映射的人员测试模型。
    /// </summary>
    private sealed class RichPerson
    {
        /// <summary>
        /// 获取或设置测试人员姓名。
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 获取或设置测试人员注册日期。
        /// </summary>
        public DateTime RegisteredAt { get; set; }
        /// <summary>
        /// 获取或设置测试备注。
        /// </summary>
        public string Note { get; set; }
    }

    /// <summary>
    /// 承载日期系统测试结果的工作簿。
    /// </summary>
    private sealed class DateImportWorkbook
    {
        /// <summary>
        /// 获取日期导入结果集合。
        /// </summary>
        public List<DateValue> Items { get; } = new List<DateValue>();
    }

    /// <summary>
    /// 用于日期系统还原的测试数据行。
    /// </summary>
    private sealed class DateValue
    {
        /// <summary>
        /// 获取或设置还原后的日期。
        /// </summary>
        public DateTime Date { get; set; }
    }

    /// <summary>
    /// 承载公式缓存读取结果的测试工作簿。
    /// </summary>
    private sealed class FormulaImportWorkbook
    {
        /// <summary>
        /// 获取公式缓存导入结果集合。
        /// </summary>
        public List<FormulaValue> Items { get; } = new List<FormulaValue>();
    }

    /// <summary>
    /// 用于公式缓存值映射的测试数据行。
    /// </summary>
    private sealed class FormulaValue
    {
        /// <summary>
        /// 获取或设置导入的公式缓存数值。
        /// </summary>
        public int Value { get; set; }
    }

    /// <summary>
    /// 承载真实 XLSB 样例导入结果的测试工作簿。
    /// </summary>
    private sealed class XlsbWorkbook
    {
        /// <summary>
        /// 获取工作表导入的数据行集合。
        /// </summary>
        public List<XlsbRow> Rows { get; } = new List<XlsbRow>();
    }

    /// <summary>
    /// 匹配真实 XLSB 样例表头的测试数据行。
    /// </summary>
    private sealed class XlsbRow
    {
        /// <summary>
        /// 获取或设置 XLSB 样例 A 列的文本值。
        /// </summary>
        public string A { get; set; }
        /// <summary>
        /// 获取或设置 XLSB 样例 B 列的布尔值。
        /// </summary>
        public bool B { get; set; }
        /// <summary>
        /// 获取或设置 XLSB 样例 C 列的日期值。
        /// </summary>
        public DateTime C { get; set; }
        /// <summary>
        /// 获取或设置 XLSB 样例 D 列的文本值。
        /// </summary>
        public string D { get; set; }
        /// <summary>
        /// 获取或设置 XLSB 样例 E 列的文本值。
        /// </summary>
        public string E { get; set; }
        /// <summary>
        /// 获取或设置 XLSB 样例 F 列的文本值。
        /// </summary>
        public string F { get; set; }
    }

    /// <summary>
    /// 首次异步读取到数据后触发取消的测试流。
    /// </summary>
    /// <remarks>不接管底层流的所有权。</remarks>
    private sealed class CancelAfterFirstReadStream : Stream
    {
        /// <summary>
        /// 提供测试输入内容的底层流，其所有权保留给调用方。
        /// </summary>
        private readonly Stream _inner;
        /// <summary>
        /// 在首次读取到数据后触发取消的令牌源。
        /// </summary>
        private readonly CancellationTokenSource _cancellation;

        /// <summary>
        /// 初始化一个 <see cref="CancelAfterFirstReadStream"/> 类型的实例。
        /// </summary>
        /// <param name="inner">提供测试输入内容的底层流。</param>
        /// <param name="cancellation">首次读取到数据后需要触发的取消令牌源。</param>
        internal CancelAfterFirstReadStream(Stream inner, CancellationTokenSource cancellation)
        {
            _inner = inner;
            _cancellation = cancellation;
        }

        /// <summary>
        /// 获取是否已读取到输入数据。
        /// </summary>
        internal bool HasRead { get; private set; }
        /// <summary>
        /// 获取包装流是否已被释放。
        /// </summary>
        internal bool WasDisposed { get; private set; }
/// <inheritdoc />
        public override bool CanRead => !WasDisposed && _inner.CanRead;
/// <inheritdoc />
        public override bool CanSeek => !WasDisposed && _inner.CanSeek;
/// <inheritdoc />
        public override bool CanWrite => false;
/// <inheritdoc />
        public override long Length => _inner.Length;
/// <inheritdoc />
        public override long Position { get => _inner.Position; set => _inner.Position = value; }
/// <inheritdoc />
        public override void Flush() => _inner.Flush();
/// <inheritdoc />
        public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
/// <inheritdoc />
        public override void SetLength(long value) => throw new System.NotSupportedException();
/// <inheritdoc />
        public override void Write(byte[] buffer, int offset, int count) => throw new System.NotSupportedException();
/// <inheritdoc />
        public override int Read(byte[] buffer, int offset, int count) =>
            throw new System.NotSupportedException("ImportAsync must use asynchronous input IO.");
/// <inheritdoc />
        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count,
            CancellationToken cancellationToken)
        {
            var read = await _inner.ReadAsync(buffer, offset, count, cancellationToken);
            if (read > 0 && !HasRead)
            {
                HasRead = true;
                _cancellation.Cancel();
            }
            return read;
        }
/// <inheritdoc />
        protected override void Dispose(bool disposing)
        {
            WasDisposed = true;
            base.Dispose(disposing);
        }
    }
}
