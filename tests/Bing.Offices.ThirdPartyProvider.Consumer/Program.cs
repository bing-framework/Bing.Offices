using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Providers;
using Bing.Offices.Attributes;
using Bing.Offices.IO;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.Testing.Entities;

namespace Bing.Offices.ThirdPartyProvider.Consumer;

/// <summary>
/// 只引用已发布包的第三方 Provider 合同验证程序。
/// </summary>
internal static class Program
{
    /// <summary>
    /// 执行公开 Provider 与文件提交契约验证。
    /// </summary>
    /// <returns>异步操作；完成后返回验证成功时的进程退出码 0。</returns>
    public static async Task<int> Main()
    {
        await VerifyPublicFileCommitters();
        await VerifyWorkbookExportStrategy();
        await VerifyEntityLayoutPublicSurface();
        var layout = ExcelEntity.Layout<FixtureEntity>(builder =>
            builder.Cell("Data", "A1", entity => entity.Value));
        var provider = new FixtureProvider();

        using var output = new MemoryStream();
        ((IExcelExporter)provider).ExportEntity(new FixtureEntity { Value = "sync" }, layout, output);
        Ensure(provider.EntityExportCalls == 1 && output.Length > 0,
            "公开同步 Entity exporter 分派未执行。");

        output.SetLength(0);
        await ((IExcelExporter)provider).ExportEntityAsync(new FixtureEntity { Value = "async" }, layout, output);
        Ensure(provider.EntityExportAsyncCalls == 1 && output.Length > 0,
            "公开异步 Entity exporter 分派未执行。");

        using var templateStream = new MemoryStream(new byte[] { 1, 2, 3 });
        using var templateOutput = new MemoryStream();
        ((IExcelExporter)provider).ExportForTemplate(new FixtureEntity { Value = "template" }, layout,
            new ExcelEntityTemplateOptions(templateStream, leaveOpen: true), templateOutput);
        Ensure(provider.TemplateExportCalls == 1 && templateStream.CanRead,
            "公开 Template exporter 分派或流所有权合同失败。");

        templateOutput.SetLength(0);
        using var asyncTemplateStream = new MemoryStream(new byte[] { 7, 8, 9 });
        await ((IExcelExporter)provider).ExportForTemplateAsync(new FixtureEntity { Value = "template-async" },
            layout, new ExcelEntityTemplateOptions(asyncTemplateStream, leaveOpen: true), templateOutput);
        Ensure(provider.TemplateExportAsyncCalls == 1 && asyncTemplateStream.CanRead,
            "公开异步 Template exporter 分派或流所有权合同失败。");

        using var source = new MemoryStream(new byte[] { 4, 5, 6 });
        var imported = ((IExcelImporter)provider).ImportEntity(source, layout);
        Ensure(provider.EntityImportCalls == 1 && imported.IsSuccess,
            "公开同步 Entity importer 分派未执行。");
        var importedAsync = await ((IExcelImporter)provider).ImportEntityAsync(source, layout);
        Ensure(provider.EntityImportAsyncCalls == 1 && importedAsync.IsSuccess,
            "公开异步 Entity importer 分派未执行。");
        using var importTemplate = new MemoryStream(new byte[] { 10, 11, 12 });
        var templateImported = ((IExcelImporter)provider).ImportForTemplate(source, layout,
            new ExcelEntityTemplateOptions(importTemplate, leaveOpen: true));
        Ensure(provider.TemplateImportCalls == 1 && templateImported.IsSuccess && importTemplate.CanRead,
            "公开 Template importer 分派或流所有权合同失败。");
        using var asyncImportTemplate = new MemoryStream(new byte[] { 13, 14, 15 });
        var templateImportedAsync = await ((IExcelImporter)provider).ImportForTemplateAsync(source, layout,
            new ExcelEntityTemplateOptions(asyncImportTemplate, leaveOpen: true));
        Ensure(provider.TemplateImportAsyncCalls == 1 && templateImportedAsync.IsSuccess && asyncImportTemplate.CanRead,
            "公开异步 Template importer 分派或流所有权合同失败。");
        Ensure(source.CanRead && output.CanWrite,
            "Provider 不得关闭调用方拥有的 source/destination 流。");

        var path = Path.Combine(Path.GetTempPath(), "bing-offices-third-party-"
            + Guid.NewGuid().ToString("N") + ".xlsx");
        try
        {
            ((IExcelExporter)provider).ExportEntityToFile(new FixtureEntity { Value = "file" }, layout, path);
            Ensure(File.Exists(path) && new FileInfo(path).Length > 0,
                "公开 Entity file exporter 未产生完整文件。");
            await ((IExcelExporter)provider).ExportEntityToFileAsync(
                new FixtureEntity { Value = "file-async" }, layout, path);
            Ensure(File.Exists(path) && new FileInfo(path).Length > 0,
                "公开异步 Entity file exporter 未产生完整文件。");
        }
        finally
        {
            if (File.Exists(path))
                File.Delete(path);
        }

        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        Expect<OperationCanceledException>(() => ((IExcelExporter)provider).ExportEntityAsync(
            new FixtureEntity(), layout, output, canceled.Token).GetAwaiter().GetResult(),
            "异步 Entity exporter 未原样观察取消令牌。");
        Expect<OperationCanceledException>(() => ((IExcelImporter)provider).ImportEntityAsync(
            source, layout, canceled.Token).GetAwaiter().GetResult(),
            "异步 Entity importer 未原样观察取消令牌。");

        var mergeLayout = ExcelEntity.Layout<FixtureEntity>(builder => builder
            .Merge("Data", "A1:B1")
            .Cell("Data", "A1", entity => entity.Value));
        var limited = new FixtureProvider(ExcelProviderCapabilities.Entity |
            ExcelProviderCapabilities.Template | ExcelProviderCapabilities.Async |
            ExcelProviderCapabilities.Xlsx);
        var mergeException = Expect<BingOfficesUnsupportedFeatureException>(() => ((IExcelExporter)limited).ExportEntity(
            new FixtureEntity(), mergeLayout, output),
            "缺少 Merge 能力时必须在底层 Entity exporter 前 fail-fast。");
        Ensure(mergeException.Code == BingOfficesErrorCode.UnsupportedFeature &&
            mergeException.Operation == BingOfficesOperation.Export &&
            mergeException.Stage == BingOfficesStage.Preflight &&
            mergeException.Provider == "ThirdPartyFixture",
            "Merge 能力缺失时的结构化异常上下文不完整。");
        Ensure(limited.EntityExportCalls == 0,
            "能力 preflight 失败后不应调用第三方 Provider。");

        var noAsync = new FixtureProvider(ExcelProviderCapabilities.Entity | ExcelProviderCapabilities.Xlsx);
        var asyncException = Expect<BingOfficesUnsupportedFeatureException>(() => ((IExcelExporter)noAsync).ExportEntityAsync(
            new FixtureEntity(), layout, output).GetAwaiter().GetResult(),
            "缺少 Async 能力时必须拒绝异步入口。");
        Ensure(asyncException.Code == BingOfficesErrorCode.UnsupportedFeature &&
            asyncException.Stage == BingOfficesStage.Preflight,
            "Async 能力缺失时的结构化异常上下文不完整。");
        var noTemplate = new FixtureProvider(ExcelProviderCapabilities.Entity |
            ExcelProviderCapabilities.Async | ExcelProviderCapabilities.Xlsx);
        var templateException = Expect<BingOfficesUnsupportedFeatureException>(() => ((IExcelExporter)noTemplate).ExportForTemplate(
            new FixtureEntity(), layout, new ExcelEntityTemplateOptions(new MemoryStream()), output),
            "缺少 Template 能力时必须拒绝模板入口。");
        Ensure(templateException.Code == BingOfficesErrorCode.UnsupportedFeature &&
            templateException.Stage == BingOfficesStage.Preflight,
            "Template 能力缺失时的结构化异常上下文不完整。");

        var miniExcelExporter = new MiniExcelExcelExporter();
        try
        {
            miniExcelExporter.ExportEntity(new FixtureEntity(), layout, new MemoryStream());
            throw new InvalidOperationException("MiniExcel Entity unsupported 未 fail-fast。");
        }
        catch (BingOfficesUnsupportedFeatureException exception)
        {
            Ensure(exception.Code == BingOfficesErrorCode.UnsupportedFeature &&
                exception.Operation == BingOfficesOperation.Export &&
                exception.Provider == "MiniExcel" && exception.Stage == BingOfficesStage.Preflight,
                "MiniExcel unsupported 异常缺少稳定 Provider/Stage 上下文。");
        }

        Ensure(provider.Supports(ExcelProviderCapabilities.Entity | ExcelProviderCapabilities.Async),
            "第三方 Provider capability 组合判断失败。");
        Console.WriteLine("third-party-public-only-provider-ok");
        return 0;
    }

    /// <summary>
    /// 验证包消费者可显式选择完整与前向流式导出。
    /// </summary>
    private static async Task VerifyWorkbookExportStrategy()
    {
        var strategy = new ExcelWorkbookExportStrategy(new ClosedXmlExcelExporter(),
            new SpreadCheetahStreamingExcelExporter());
        var request = ExcelExport.Workbook(book => book.AddSheet("Rows",
            new[] { new FixtureEntity { Value = "strategy" } }));
        using var output = new MemoryStream();
        strategy.Export(request, output, ExcelWorkbookExportMode.CompleteWorkbook);
        Ensure(output.Length > 0 && output.CanWrite, "完整工作簿策略未写入目标流。");

        output.SetLength(0);
        await strategy.ExportAsync(request, output, ExcelWorkbookExportMode.ForwardStreaming,
            new ExcelStreamingExportOptions { BatchSize = 1 });
        Ensure(output.Length > 0 && output.CanWrite, "前向流式策略未写入目标流。");

        output.SetLength(0);
        using var template = new MemoryStream(new byte[] { 1, 2, 3 });
        var templateRequest = ExcelExport.Workbook(book => book.UseTemplate(template, leaveOpen: true)
            .AddSheet("Rows", new[] { new FixtureEntity { Value = "strategy" } }));
        var unsupported = Expect<BingOfficesUnsupportedFeatureException>(() =>
            strategy.Export(templateRequest, output, ExcelWorkbookExportMode.ForwardStreaming),
            "前向流式策略必须在写入前拒绝模板。");
        Ensure(unsupported.Stage == BingOfficesStage.Preflight && output.Length == 0 && template.CanRead,
            "前向流式预检或流所有权合同失败。");
    }

    /// <summary>
    /// 通过包公开 API 验证属性布局、动态列分组、分组小计、分页小计、尾部聚合和分页符。
    /// </summary>
    private static async Task VerifyEntityLayoutPublicSurface()
    {
        var layout = ExcelEntity.LayoutFromAttributes<ConsumerEntity>(builder => builder
            .ListRegion("Order", "A4", entity => entity.Lines, region => region
                .DynamicColumnGroup("商品组", item => item.Goods, new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "goods-color",
                        Title = "商品颜色",
                        DataType = typeof(string)
                    }
                })
                .DynamicColumnGroup("产品组", item => item.Product, new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "product-weight",
                        Title = "产品重量",
                        DataType = typeof(decimal)
                    }
                })
                .Footer("TOTAL", footer => footer
                    .Cell("B1", items => items.Sum(item => item.Quantity)))
                .GroupSubtotal(item => item.Name, "SUBTOTAL", footer => footer
                    .Cell("B1", items => items.Sum(item => item.Quantity)))));
        var pageBreakLayout = ExcelEntity.Layout<ConsumerEntity>(builder => builder
            .ListRegion("Order", "A4", entity => entity.Lines, region => region
                .DynamicColumnGroup("商品组", item => item.Goods, new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "goods-color",
                        Title = "商品颜色",
                        DataType = typeof(string)
                    }
                })
                .DynamicColumnGroup("产品组", item => item.Product, new[]
                {
                    new ExcelDynamicColumnDefinition
                    {
                        Key = "product-weight",
                        Title = "产品重量",
                        DataType = typeof(decimal)
                    }
                })
                .PageBreak(1)));
        var pageSubtotalLayout = ExcelEntity.Layout<ConsumerEntity>(builder => builder
            .ListRegion("Order", "A4", entity => entity.Lines, region => region
                .PageBreak(1)
                .PageSubtotal("PAGE", footer => footer
                    .Cell("B1", items => items.Sum(item => item.Quantity)))
                .FooterNamed("OrderTotal", "TOTAL", footer => footer
                    .Cell("B1", items => items.Sum(item => item.Quantity))
                    .Formula("C1", "=1+1", numberFormat: "0.00")
                    .FormulaSumDetailRowsAbove("C2", "B", numberFormat: "0.00"))));
        var contiguousSumLayout = ExcelEntity.Layout<ConsumerEntity>(builder => builder
            .ListRegion("Order", "A4", entity => entity.Lines, region => region
                .Footer("TOTAL", footer => footer
                    .FormulaSumContiguousRowsAbove("C2", "B", numberFormat: "0.00")
                    .GapRows(1))));
        var source = new ConsumerEntity
        {
            Code = "consumer-order",
            Lines = new List<ConsumerLine>
            {
                new ConsumerLine
                {
                    Name = "A",
                    Quantity = 2,
                    Goods = new Dictionary<string, object> { ["goods-color"] = "red" },
                    Product = new Dictionary<string, object> { ["product-weight"] = 1.5m }
                },
                new ConsumerLine
                {
                    Name = "B",
                    Quantity = 3,
                    Goods = new Dictionary<string, object> { ["goods-color"] = "blue" },
                    Product = new Dictionary<string, object> { ["product-weight"] = 2.5m }
                }
            }
        };
        var providers = new (IExcelExporter Exporter, IExcelImporter Importer)[]
        {
            (new NpoiExcelExporter(), new NpoiExcelImporter()),
            (new ClosedXmlExcelExporter(), new ClosedXmlExcelImporter())
        };
        foreach (var provider in providers)
        {
            var entityExporter = (IExcelEntityExporter)provider.Exporter;
            var entityImporter = (IExcelEntityImporter)provider.Importer;
            EntityLayoutProviderContractSuite.VerifyCore(entityExporter, entityImporter);
            EntityLayoutProviderContractSuite.VerifyPageSubtotal(entityExporter, entityImporter);
            EntityLayoutProviderContractSuite.VerifyGroupSubtotal(entityExporter, entityImporter);
            EntityLayoutProviderContractSuite.VerifyFooterFormulas(entityExporter, entityImporter);
            await EntityLayoutProviderContractSuite.VerifyAsync(entityExporter, entityImporter);
            EntityLayoutProviderContractSuite.VerifyAttributeAndNamedAnchors(entityExporter, entityImporter);
            EntityLayoutProviderContractSuite.VerifyNamedListFixedCellCollision(entityExporter, entityImporter);
            using var output = new MemoryStream();
            provider.Exporter.ExportEntity(source, layout, output);
            using var input = new MemoryStream(output.ToArray(), writable: false);
            var result = provider.Importer.ImportEntity(input, layout);
            Ensure(result.IsSuccess && result.Entity.Code == source.Code
                && result.Entity.Lines.Count == 2
                && (string)result.Entity.Lines[0].Goods["goods-color"] == "red"
                && Convert.ToDecimal(result.Entity.Lines[1].Product["product-weight"]) == 2.5m,
                "包公开 Entity Layout API 往返验证失败。");

            using var pageBreakOutput = new MemoryStream();
            provider.Exporter.ExportEntity(source, pageBreakLayout, pageBreakOutput);
            Ensure(pageBreakOutput.Length > 0,
                "包公开 Entity Layout 分页符配置未能完成导出。");

            using var pageSubtotalOutput = new MemoryStream();
            provider.Exporter.ExportEntity(source, pageSubtotalLayout, pageSubtotalOutput);
            using var pageSubtotalInput = new MemoryStream(pageSubtotalOutput.ToArray(), writable: false);
            var pageSubtotalResult = provider.Importer.ImportEntity(pageSubtotalInput, pageSubtotalLayout);
            Ensure(pageSubtotalResult.IsSuccess && pageSubtotalResult.Entity.Lines.Count == source.Lines.Count,
                "包公开 Entity Layout 分页小计往返或导入边界验证失败。");
            Ensure(pageSubtotalLayout.ListRegions[0].FooterAnchorName == "OrderTotal",
                "包公开 Entity Layout 最终尾部命名锚点未保留。");

            using var contiguousSumOutput = new MemoryStream();
            provider.Exporter.ExportEntity(source, contiguousSumLayout, contiguousSumOutput);
            using var contiguousSumInput = new MemoryStream(contiguousSumOutput.ToArray(), writable: false);
            var contiguousSumResult = provider.Importer.ImportEntity(contiguousSumInput, contiguousSumLayout);
            Ensure(contiguousSumResult.IsSuccess && contiguousSumResult.Entity.Lines.Count == source.Lines.Count,
                "包公开 Entity Layout 连续明细求和公式未正确导出或导入。");
        }
    }

    /// <summary>
    /// 使用包内公开契约验证文件提交器和导入器扩展入口。
    /// </summary>
    /// <remarks>
    /// 同时覆盖默认提交器、构造注入和同步、异步失败工作簿输出。
    /// </remarks>
    private static async Task VerifyPublicFileCommitters()
    {
        var directory = Path.Combine(Path.GetTempPath(), "bing-public-committer-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var path = Path.Combine(directory, "default.bin");
            var expected = new byte[] { 1, 2, 3, 4 };
            var defaultCommitter = new DefaultFileExportCommitter();
            defaultCommitter.Commit(path, stream => stream.Write(expected, 0, expected.Length), default, "Consumer");
            Ensure(File.ReadAllBytes(path).SequenceEqual(expected), "默认同步提交器未写入完整内容。");
            await defaultCommitter.CommitAsync(path, (stream, token) => stream.WriteAsync(expected, 0, expected.Length, token), default, "Consumer");
            Ensure((await File.ReadAllBytesAsync(path)).SequenceEqual(expected), "默认异步提交器未写入完整内容。");
            File.Delete(path);

            var committer = new ConsumerFileCommitter();
            IExcelImporter[] importers =
            {
                new NpoiExcelImporter(null, null, null, null, null, committer),
                new ClosedXmlExcelImporter(null, null, null, null, null, committer)
            };
            using var source = new MemoryStream();
            new NpoiExcelExporter().Export(ExcelExport.Workbook(builder => builder
                .AddSheet("Data", new[] { new InvalidRow { Code = "", Quantity = 1 } })), source);
            var request = ExcelImport.Workbook<InvalidWorkbook>(builder => builder
                .FailureWorkbook(new ExcelImportFailureOptions
                {
                    Mode = ExcelImportFailureWorkbookMode.AnnotatedOriginal,
                    DestinationPath = Path.Combine(directory, "failure.xlsx"),
                    TemporaryDirectory = directory
                })
                .Sheet<InvalidRow>("Data", workbook => workbook.Rows,
                    sheet => sheet.Validate(ExcelValidationFailureMode.Continue)));
            foreach (var importer in importers)
            foreach (var asynchronous in new[] { false, true })
            {
                source.Position = 0;
                var beforeSync = committer.SyncCalls;
                var beforeAsync = committer.AsyncCalls;
                var result = asynchronous ? await importer.ImportAsync(source, request) : importer.Import(source, request);
                Ensure(!result.IsSuccess && result.Errors.Count == 1, "包内导入器未生成预期校验错误。");
                Ensure(committer.SyncCalls - beforeSync == (asynchronous ? 0 : 1)
                    && committer.AsyncCalls - beforeAsync == (asynchronous ? 1 : 0), "失败工作簿未使用指定提交器或重复提交。");
                // 再次通过公开导入入口读取完整失败产物，验证其不是空文件或损坏文件。
                using var output = File.OpenRead(request.FailureOptions.DestinationPath);
                var reread = importer.Import(output, ExcelImport.Workbook<InvalidWorkbook>(builder => builder
                    .Sheet<InvalidRow>("Data", workbook => workbook.Rows,
                        sheet => sheet.Validate(ExcelValidationFailureMode.Continue))));
                Ensure(reread.Errors.Count == 1
                    && reread.Errors[0].Code == result.Errors[0].Code
                    && reread.Errors[0].SheetName == "Data"
                    && reread.Errors[0].RowIndex == 2
                    && reread.Errors[0].ColumnIndex == 1
                    && reread.Errors[0].PropertyName == "Code", "失败工作簿重读后的完整错误位置不一致。");
            }
            Ensure(Directory.GetFiles(directory).SequenceEqual(new[] { request.FailureOptions.DestinationPath }),
                "公开提交器遗留临时文件。");
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    /// <summary>
    /// 通过公开接口装饰默认文件提交器。
    /// </summary>
    private sealed class ConsumerFileCommitter : IFileExportCommitter
    {
        /// <summary>
        /// 默认文件提交实现。
        /// </summary>
        private readonly DefaultFileExportCommitter _inner = new();

        /// <summary>
        /// 获取同步提交调用次数。
        /// </summary>
        public int SyncCalls { get; private set; }

        /// <summary>
        /// 获取异步提交调用次数。
        /// </summary>
        public int AsyncCalls { get; private set; }

        /// <inheritdoc />
        public void Commit(string path, Action<Stream> write, CancellationToken cancellationToken, string format)
        {
            SyncCalls++;
            _inner.Commit(path, write, cancellationToken, format);
        }
        /// <inheritdoc />
        public Task CommitAsync(string path, Func<Stream, CancellationToken, Task> writeAsync,
            CancellationToken cancellationToken, string format)
        {
            AsyncCalls++;
            return _inner.CommitAsync(path, writeAsync, cancellationToken, format);
        }
    }

    /// <summary>
    /// 失败工作簿验证使用的行模型。
    /// </summary>
    private sealed class InvalidRow
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
    /// 失败工作簿验证使用的根模型。
    /// </summary>
    private sealed class InvalidWorkbook
    {
        /// <summary>
        /// 获取或设置数据行集合。
        /// </summary>
        public List<InvalidRow> Rows { get; set; } = new();
    }

    /// <summary>
    /// 验证第三方 fixture 的断言。
    /// </summary>
    /// <param name="condition">待验证条件。</param>
    /// <param name="message">失败消息。</param>
    private static void Ensure(bool condition, string message)
    {
        if (!condition)
            throw new InvalidOperationException(message);
    }

    /// <summary>
    /// 断言指定操作抛出预期异常类型。
    /// </summary>
    /// <typeparam name="TException">预期异常类型。</typeparam>
    /// <param name="action">待执行操作。</param>
    /// <param name="message">未抛出预期异常时的失败消息。</param>
    private static TException Expect<TException>(Action action, string message)
        where TException : Exception
    {
        try
        {
            action();
        }
        catch (TException exception)
        {
            return exception;
        }
        throw new InvalidOperationException(message);
    }

    /// <summary>
    /// fixture 使用的最小实体类型。
    /// </summary>
    private sealed class FixtureEntity
    {
        /// <summary>
        /// 获取或设置固定单元格值。
        /// </summary>
        public string Value { get; set; }
    }

    /// <summary>
    /// 只依赖公开 Bing.Offices 契约的第三方 Provider 替身。
    /// </summary>
    private sealed class FixtureProvider : IExcelImporter, IExcelExporter,
        IExcelEntityImporter, IExcelEntityExporter, IExcelProviderCapabilities
    {
        /// <summary>
        /// 初始化一个 <see cref="FixtureProvider" /> 类型的实例。
        /// </summary>
        /// <param name="capabilities">Provider 声明的能力。</param>
        public FixtureProvider(ExcelProviderCapabilities? capabilities = null)
        {
            Capabilities = capabilities ?? (ExcelProviderCapabilities.List |
                ExcelProviderCapabilities.Workbook | ExcelProviderCapabilities.Entity |
                ExcelProviderCapabilities.Template | ExcelProviderCapabilities.Merge |
                ExcelProviderCapabilities.Async | ExcelProviderCapabilities.Xlsx);
        }

        /// <summary>
        /// 获取同步实体导入调用次数。
        /// </summary>
        public int EntityImportCalls { get; private set; }
        /// <summary>
        /// 获取异步实体导入调用次数。
        /// </summary>
        public int EntityImportAsyncCalls { get; private set; }
        /// <summary>
        /// 获取同步实体导出调用次数。
        /// </summary>
        public int EntityExportCalls { get; private set; }
        /// <summary>
        /// 获取异步实体导出调用次数。
        /// </summary>
        public int EntityExportAsyncCalls { get; private set; }
        /// <summary>
        /// 获取模板导出调用次数。
        /// </summary>
        public int TemplateExportCalls { get; private set; }
        /// <summary>
        /// 获取异步模板导出调用次数。
        /// </summary>
        public int TemplateExportAsyncCalls { get; private set; }
        /// <summary>
        /// 获取模板导入调用次数。
        /// </summary>
        public int TemplateImportCalls { get; private set; }
        /// <summary>
        /// 获取异步模板导入调用次数。
        /// </summary>
        public int TemplateImportAsyncCalls { get; private set; }

        /// <inheritdoc />
        public string ProviderName => "ThirdPartyFixture";
        /// <inheritdoc />
        public ExcelProviderCapabilities Capabilities { get; }
        /// <inheritdoc />
        public bool Supports(ExcelProviderCapabilities capabilities) =>
            (Capabilities & capabilities) == capabilities;

        /// <inheritdoc />
        public ExcelWorkbookImportResult<TWorkbook> Import<TWorkbook>(Stream source,
            ExcelWorkbookImportRequest<TWorkbook> request,
            CancellationToken cancellationToken = default) where TWorkbook : class, new() =>
            CreateWorkbookResult<TWorkbook>();

        /// <inheritdoc />
        public Task<ExcelWorkbookImportResult<TWorkbook>> ImportAsync<TWorkbook>(Stream source,
            ExcelWorkbookImportRequest<TWorkbook> request,
            CancellationToken cancellationToken = default) where TWorkbook : class, new() =>
            Task.FromResult(CreateWorkbookResult<TWorkbook>());

        /// <inheritdoc />
        public void Export(ExcelWorkbookExportRequest request, Stream destination,
            CancellationToken cancellationToken = default) => destination.WriteByte(1);

        /// <inheritdoc />
        public Task ExportAsync(ExcelWorkbookExportRequest request, Stream destination,
            CancellationToken cancellationToken = default)
        {
            destination.WriteByte(2);
            return Task.CompletedTask;
        }

        /// <inheritdoc />
        public void ExportToFile(ExcelWorkbookExportRequest request, string path,
            CancellationToken cancellationToken = default) => File.WriteAllBytes(path, new byte[] { 3 });

        /// <inheritdoc />
        public Task ExportToFileAsync(ExcelWorkbookExportRequest request, string path,
            CancellationToken cancellationToken = default)
        {
            File.WriteAllBytes(path, new byte[] { 4 });
            return Task.CompletedTask;
        }

        /// <inheritdoc />
        public ExcelEntityImportResult<TEntity> ImportEntity<TEntity>(Stream source,
            ExcelEntityLayout<TEntity> layout, CancellationToken cancellationToken = default)
            where TEntity : class, new()
        {
            cancellationToken.ThrowIfCancellationRequested();
            EntityImportCalls++;
            return CreateEntityResult<TEntity>();
        }

        /// <inheritdoc />
        public Task<ExcelEntityImportResult<TEntity>> ImportEntityAsync<TEntity>(Stream source,
            ExcelEntityLayout<TEntity> layout, CancellationToken cancellationToken = default)
            where TEntity : class, new()
        {
            cancellationToken.ThrowIfCancellationRequested();
            EntityImportAsyncCalls++;
            return Task.FromResult(CreateEntityResult<TEntity>());
        }

        /// <inheritdoc />
        public ExcelEntityImportResult<TEntity> ImportForTemplate<TEntity>(Stream source,
            ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
            CancellationToken cancellationToken = default) where TEntity : class, new() =>
            ImportTemplate<TEntity>(cancellationToken, false);

        /// <inheritdoc />
        public Task<ExcelEntityImportResult<TEntity>> ImportForTemplateAsync<TEntity>(Stream source,
            ExcelEntityLayout<TEntity> layout, ExcelEntityTemplateOptions template,
            CancellationToken cancellationToken = default) where TEntity : class, new() =>
            Task.FromResult(ImportTemplate<TEntity>(cancellationToken, true));

        /// <inheritdoc />
        public void ExportEntity<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout,
            Stream destination, CancellationToken cancellationToken = default)
            where TEntity : class, new()
        {
            cancellationToken.ThrowIfCancellationRequested();
            EntityExportCalls++;
            destination.WriteByte(5);
        }

        /// <inheritdoc />
        public Task ExportEntityAsync<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout,
            Stream destination, CancellationToken cancellationToken = default)
            where TEntity : class, new()
        {
            cancellationToken.ThrowIfCancellationRequested();
            EntityExportAsyncCalls++;
            destination.WriteByte(6);
            return Task.CompletedTask;
        }

        /// <inheritdoc />
        public void ExportEntityToFile<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout,
            string path, CancellationToken cancellationToken = default) where TEntity : class, new() =>
            File.WriteAllBytes(path, new byte[] { 7 });

        /// <inheritdoc />
        public Task ExportEntityToFileAsync<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout,
            string path, CancellationToken cancellationToken = default) where TEntity : class, new()
        {
            File.WriteAllBytes(path, new byte[] { 8 });
            return Task.CompletedTask;
        }

        /// <inheritdoc />
        public void ExportForTemplate<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout,
            ExcelEntityTemplateOptions template, Stream destination,
            CancellationToken cancellationToken = default) where TEntity : class, new()
        {
            TemplateExportCalls++;
            destination.WriteByte(9);
        }

        /// <inheritdoc />
        public Task ExportForTemplateAsync<TEntity>(TEntity entity, ExcelEntityLayout<TEntity> layout,
            ExcelEntityTemplateOptions template, Stream destination,
            CancellationToken cancellationToken = default) where TEntity : class, new()
        {
            cancellationToken.ThrowIfCancellationRequested();
            TemplateExportAsyncCalls++;
            destination.WriteByte(10);
            return Task.CompletedTask;
        }

        /// <summary>
        /// 创建模板导入结果并观察取消令牌。
        /// </summary>
        /// <typeparam name="TEntity">实体类型。</typeparam>
        /// <param name="cancellationToken">取消令牌。</param>
        /// <param name="async">是否记录异步调用。</param>
        /// <returns>空的实体导入结果。</returns>
        private ExcelEntityImportResult<TEntity> ImportTemplate<TEntity>(CancellationToken cancellationToken,
            bool async) where TEntity : class, new()
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (async)
                TemplateImportAsyncCalls++;
            else
                TemplateImportCalls++;
            return CreateEntityResult<TEntity>();
        }

        /// <summary>
        /// 创建空的公开 Workbook 结果。
        /// </summary>
        /// <typeparam name="TWorkbook">Workbook 类型。</typeparam>
        /// <returns>空 Workbook 结果。</returns>
        private static ExcelWorkbookImportResult<TWorkbook> CreateWorkbookResult<TWorkbook>()
            where TWorkbook : class, new() => new ExcelWorkbookImportResult<TWorkbook>(new TWorkbook(),
                Array.Empty<ExcelSheetImportResult>(), Array.Empty<ExcelImportError>(), false, null);

        /// <summary>
        /// 创建空的公开 Entity 结果。
        /// </summary>
        /// <typeparam name="TEntity">Entity 类型。</typeparam>
        /// <returns>空 Entity 结果。</returns>
        private static ExcelEntityImportResult<TEntity> CreateEntityResult<TEntity>()
            where TEntity : class, new() => new ExcelEntityImportResult<TEntity>(new TEntity(),
                Array.Empty<ExcelImportError>(), Array.Empty<ExcelSheetImportResult>());
    }
}

/// <summary>
/// 包消费者用于验证属性式单据布局的根实体。
/// </summary>
public sealed class ConsumerEntity
{
    /// <summary>
    /// 获取或设置单据编号。
    /// </summary>
    [ExcelEntityCell("Order", "B2")]
    public string Code { get; set; }

    /// <summary>
    /// 获取或设置单据明细集合。
    /// </summary>
    public List<ConsumerLine> Lines { get; set; } = new();
}

/// <summary>
/// 包消费者用于验证多个动态字段字典的明细实体。
/// </summary>
public sealed class ConsumerLine
{
    /// <summary>
    /// 获取或设置明细名称。
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// 获取或设置明细数量。
    /// </summary>
    public int Quantity { get; set; }

    /// <summary>
    /// 获取或设置商品动态字段。
    /// </summary>
    public IDictionary<string, object> Goods { get; set; }

    /// <summary>
    /// 获取或设置产品动态字段。
    /// </summary>
    public IDictionary<string, object> Product { get; set; }
}
