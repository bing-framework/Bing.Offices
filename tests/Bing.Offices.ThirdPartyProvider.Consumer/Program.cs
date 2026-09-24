using Bing.Offices.Entities;
using Bing.Offices.Exceptions;
using Bing.Offices.Exports;
using Bing.Offices.Extensions;
using Bing.Offices.Imports;
using Bing.Offices.Providers;

namespace Bing.Offices.ThirdPartyProvider.Consumer;

/// <summary>
/// 只引用已发布包的第三方 Provider 合同验证程序。
/// </summary>
internal static class Program
{
    /// <summary>
    /// 执行公开 Entity/Template/SPI 分派验证。
    /// </summary>
    /// <returns>验证成功时返回零。</returns>
    public static async Task<int> Main()
    {
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
        /// 使用完整或指定能力集合创建 fixture Provider。
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
