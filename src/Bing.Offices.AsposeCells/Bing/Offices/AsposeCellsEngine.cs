using System.Collections;
using System.Reflection;
using Bing.Offices.Conversions;
using Bing.Offices.Exceptions;
using Bing.Offices.Formula;
using Bing.Offices.Rendering;
using Bing.Offices.Providers;

namespace Bing.Offices.AsposeCells;

/// <summary>
/// Aspose.Cells 可选商业能力适配器。
/// </summary>
/// <remarks>
/// 通过反射隔离商业引擎的具体 API，避免 Core 和 Abstractions 传递商业依赖；
/// 运行时仍要求宿主部署 Aspose.Cells 26.8.0 及有效许可证。
/// </remarks>
public sealed class AsposeCellsEngine : IExcelDocumentRenderer, IExcelPageRenderer,
    IExcelDocumentConverter, IExcelFormulaProcessor, IExcelProviderFeatureDescriptor
{
    /// <summary>
    /// 用于标识错误和能力描述所属提供程序的名称。
    /// </summary>
    private const string Provider = "Aspose.Cells";
    /// <summary>
    /// 当前引擎独立持有的宿主配置副本。
    /// </summary>
    private readonly AsposeCellsProviderOptions _options;
    /// <summary>
    /// 保护许可证和字体初始化状态的同步锁。
    /// </summary>
    private readonly object _gate = new();
    /// <summary>
    /// 指示许可证及字体配置是否已完成初始化。
    /// </summary>
    private bool _licenseConfigured;

    /// <summary>
    /// 初始化一个 <see cref="AsposeCellsEngine"/> 类型的实例。
    /// </summary>
    /// <param name="options">宿主配置。</param>
    public AsposeCellsEngine(AsposeCellsProviderOptions options = null)
    {
        _options = options?.Clone() ?? new AsposeCellsProviderOptions();
    }

    /// <inheritdoc />
    public string ProviderName => Provider;
    /// <inheritdoc />
    public ExcelProviderCapabilities Capabilities => ExcelProviderCapabilities.Workbook
        | ExcelProviderCapabilities.Async | ExcelProviderCapabilities.Xls
        | ExcelProviderCapabilities.Xlsx | ExcelProviderCapabilities.Xlsm
        | ExcelProviderCapabilities.Ods;
    /// <inheritdoc />
    public bool Supports(ExcelProviderCapabilities capabilities) => (Capabilities & capabilities) == capabilities;
    /// <inheritdoc />
    public IReadOnlyList<ExcelFormat> ReadFormats { get; } = new[]
        { ExcelFormat.Xls, ExcelFormat.Xlsx, ExcelFormat.Xlsm, ExcelFormat.Ods };
    /// <inheritdoc />
    public IReadOnlyList<ExcelFormat> WriteFormats { get; } = new[]
        { ExcelFormat.Xls, ExcelFormat.Xlsx, ExcelFormat.Xlsm, ExcelFormat.Ods };
    /// <inheritdoc />
    public bool SupportsCompleteWorkbookImport => false;
    /// <inheritdoc />
    public bool SupportsBatchImport => false;
    /// <inheritdoc />
    public bool SupportsCompleteWorkbookExport => false;
    /// <inheritdoc />
    public bool SupportsTrueAsyncIo => true;
    /// <inheritdoc />
    public IReadOnlyList<string> Limitations { get; } = new[]
    {
        "Aspose.Cells 许可证、字体目录和商业部署授权由宿主负责。",
        "宏只允许保留或显式剥离，不执行、不编辑 VBA。",
        "ODS 与 XLSX 转换可能产生格式损失，结果通过 warning 返回。"
    };
    /// <inheritdoc />
    public ExcelProviderFeatures Features => ExcelProviderFeatures.WorkbookEditing
        | ExcelProviderFeatures.TemplateEditing | ExcelProviderFeatures.Tables
        | ExcelProviderFeatures.AutoFilter | ExcelProviderFeatures.FreezePanes
        | ExcelProviderFeatures.ConditionalFormatting | ExcelProviderFeatures.NamedRanges
        | ExcelProviderFeatures.PrintLayout | ExcelProviderFeatures.FormulaText
        | ExcelProviderFeatures.FormulaCachedValues | ExcelProviderFeatures.FormulaRecalculation
        | ExcelProviderFeatures.PdfRendering
        | ExcelProviderFeatures.PageImageRendering | ExcelProviderFeatures.MacroPreservation
        | ExcelProviderFeatures.Ods | ExcelProviderFeatures.EncryptionRead
        | ExcelProviderFeatures.EncryptionWrite;
    /// <inheritdoc />
    public bool Supports(ExcelProviderFeatures features) => (Features & features) == features;

    /// <inheritdoc />
    public ExcelRenderResult Render(Stream source, Stream destination, ExcelRenderRequest request,
        CancellationToken cancellationToken = default)
    {
        ValidateStreams(source, destination);
        if (request == null) throw new ArgumentNullException(nameof(request));
        ValidateRenderRequest(request, renderPages: false);
        cancellationToken.ThrowIfCancellationRequested();
        EnsureReady(request.StrictFonts);
        var workbook = LoadWorkbook(source, request.InputFormat, null);
        try
        {
            using var scope = ApplyRenderScope(workbook, request);
            var saveOptions = CreatePdfSaveOptions(request);
            cancellationToken.ThrowIfCancellationRequested();
            Invoke(workbook, "Save", destination, saveOptions);
            return new ExcelRenderResult
            {
                PageCount = CountRenderedPages(workbook, request),
                Warnings = BuildFontWarnings(request.StrictFonts)
            };
        }
        finally
        {
            DisposeIfNeeded(workbook);
        }
    }

    /// <inheritdoc />
    public async Task<ExcelRenderResult> RenderAsync(Stream source, Stream destination,
        ExcelRenderRequest request, CancellationToken cancellationToken = default)
    {
        ValidateStreams(source, destination);
        await using var staged = new MemoryStream();
        await source.CopyToAsync(staged, 81920, cancellationToken).ConfigureAwait(false);
        staged.Position = 0;
        await using var rendered = new MemoryStream();
        var result = Render(staged, rendered, request, cancellationToken);
        rendered.Position = 0;
        await rendered.CopyToAsync(destination, 81920, cancellationToken).ConfigureAwait(false);
        await destination.FlushAsync(cancellationToken).ConfigureAwait(false);
        return result;
    }

    /// <inheritdoc />
    public ExcelRenderResult RenderPages(Stream source, ExcelRenderRequest request,
        Action<int, Stream, string> pageSink, CancellationToken cancellationToken = default)
    {
        if (pageSink == null) throw new ArgumentNullException(nameof(pageSink));
        ValidateStreams(source, null);
        if (request == null) throw new ArgumentNullException(nameof(request));
        ValidateRenderRequest(request, renderPages: true);
        cancellationToken.ThrowIfCancellationRequested();
        EnsureReady(request.StrictFonts);
        var workbook = LoadWorkbook(source, request.InputFormat, null);
        try
        {
            using var scope = ApplyRenderScope(workbook, request);
            var pageOffset = 0;
            var emittedPages = 0;
            foreach (var sheet in EnumerateWorksheets(workbook, request.SheetName))
            {
                cancellationToken.ThrowIfCancellationRequested();
                pageOffset += RenderSheetPages(sheet, request, pageSink, pageOffset,
                    ref emittedPages, cancellationToken);
            }
            return new ExcelRenderResult
            {
                PageCount = emittedPages,
                Warnings = BuildFontWarnings(request.StrictFonts)
            };
        }
        finally
        {
            DisposeIfNeeded(workbook);
        }
    }

    /// <inheritdoc />
    public async Task<ExcelRenderResult> RenderPagesAsync(Stream source, ExcelRenderRequest request,
        Func<int, Stream, string, Task> pageSink, CancellationToken cancellationToken = default)
    {
        if (pageSink == null) throw new ArgumentNullException(nameof(pageSink));
        ValidateStreams(source, null);
        await using var staged = new MemoryStream();
        await source.CopyToAsync(staged, 81920, cancellationToken).ConfigureAwait(false);
        staged.Position = 0;
        return await RenderPagesAsyncCore(staged, request, pageSink, cancellationToken).ConfigureAwait(false);
    }

    /// <inheritdoc />
    public ExcelDocumentConversionResult Convert(Stream source, Stream destination,
        ExcelFormat sourceFormat, ExcelFormat targetFormat, ExcelWorkbookOpenOptions openOptions = null,
        ExcelWorkbookSaveOptions saveOptions = null, CancellationToken cancellationToken = default)
    {
        ValidateStreams(source, destination);
        ValidateFormat(sourceFormat);
        ValidateFormat(targetFormat);
        openOptions ??= new ExcelWorkbookOpenOptions();
        saveOptions ??= new ExcelWorkbookSaveOptions();
        ValidateMacroPolicy(sourceFormat, targetFormat, openOptions.MacroPolicy, saveOptions.MacroPolicy);
        cancellationToken.ThrowIfCancellationRequested();
        EnsureReady(_options.StrictFonts);
        var workbook = LoadWorkbook(source, sourceFormat, openOptions.Password);
        try
        {
            ApplyMacroPolicy(workbook, saveOptions.MacroPolicy, targetFormat);
            SaveWorkbook(workbook, destination, ToSaveFormat(targetFormat), saveOptions.Password, cancellationToken);
            return new ExcelDocumentConversionResult
            {
                SourceFormat = sourceFormat,
                TargetFormat = targetFormat,
                Warnings = sourceFormat == ExcelFormat.Ods || targetFormat == ExcelFormat.Ods
                    ? new[] { "ODS 与 OOXML/BIFF 转换不承诺完全无损。" }
                    : Array.Empty<string>()
            };
        }
        finally
        {
            DisposeIfNeeded(workbook);
        }
    }

    /// <inheritdoc />
    public async Task<ExcelDocumentConversionResult> ConvertAsync(Stream source, Stream destination,
        ExcelFormat sourceFormat, ExcelFormat targetFormat, ExcelWorkbookOpenOptions openOptions = null,
        ExcelWorkbookSaveOptions saveOptions = null, CancellationToken cancellationToken = default)
    {
        ValidateStreams(source, destination);
        await using var staged = new MemoryStream();
        await source.CopyToAsync(staged, 81920, cancellationToken).ConfigureAwait(false);
        staged.Position = 0;
        await using var converted = new MemoryStream();
        var result = Convert(staged, converted, sourceFormat, targetFormat, openOptions, saveOptions, cancellationToken);
        converted.Position = 0;
        await converted.CopyToAsync(destination, 81920, cancellationToken).ConfigureAwait(false);
        await destination.FlushAsync(cancellationToken).ConfigureAwait(false);
        return result;
    }

    /// <inheritdoc />
    public ExcelFormulaResult Process(Stream source, ExcelFormat format, ExcelFormulaRequest request,
        CancellationToken cancellationToken = default)
    {
        ValidateStreams(source, null);
        ValidateFormat(format);
        if (request == null) throw new ArgumentNullException(nameof(request));
        ValidateFormulaRequest(request);
        cancellationToken.ThrowIfCancellationRequested();
        EnsureReady(_options.StrictFonts);
        var workbook = LoadWorkbook(source, format, null);
        try
        {
            if (request.CalculationMode == ExcelFormulaCalculationMode.CalculateOnOpen)
                throw Unsupported("CalculateOnOpen 需要返回可保存的工作簿；公式处理接口只返回分析结果。",
                    BingOfficesOperation.Import);
            if (request.CalculationMode != ExcelFormulaCalculationMode.DoNotCalculate)
            {
                if (request.CalculationMode == ExcelFormulaCalculationMode.MustRecalculate
                    && ContainsUnsupportedFormulaReference(workbook, request.SheetName))
                    throw Unsupported("公式包含外部引用、失效引用或未支持的引用语义，无法满足 MustRecalculate。",
                        BingOfficesOperation.Import);
                try
                {
                    Invoke(workbook, "CalculateFormula");
                }
                catch (TargetInvocationException exception)
                {
                    throw Unsupported("Aspose.Cells 公式重新计算失败，无法满足请求的计算语义。",
                        BingOfficesOperation.Import, exception.InnerException ?? exception);
                }
            }
            var cells = ReadFormulaCells(workbook, request, cancellationToken);
            if (request.CalculationMode == ExcelFormulaCalculationMode.MustRecalculate
                && cells.Any(cell => cell.ErrorCode != null))
                throw Unsupported("公式计算产生错误值，无法满足 MustRecalculate。", BingOfficesOperation.Import);
            return new ExcelFormulaResult
            {
                Cells = cells,
                CalculationCompleted = request.CalculationMode != ExcelFormulaCalculationMode.MustRecalculate
                    || cells.All(cell => cell.ErrorCode == null)
            };
        }
        finally
        {
            DisposeIfNeeded(workbook);
        }
    }

    /// <inheritdoc />
    public Task<ExcelFormulaResult> ProcessAsync(Stream source, ExcelFormat format,
        ExcelFormulaRequest request, CancellationToken cancellationToken = default)
    {
        if (source == null) throw new ArgumentNullException(nameof(source));
        return ProcessAsyncCore(source, format, request, cancellationToken);
    }

    /// <summary>
    /// 异步暂存输入并处理工作簿公式。
    /// </summary>
    /// <param name="source">输入工作簿流。</param>
    /// <param name="format">目标格式。</param>
    /// <param name="request">本次处理请求。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>公式处理结果。</returns>
    private async Task<ExcelFormulaResult> ProcessAsyncCore(Stream source, ExcelFormat format,
        ExcelFormulaRequest request, CancellationToken cancellationToken)
    {
        await using var staged = new MemoryStream();
        await source.CopyToAsync(staged, 81920, cancellationToken).ConfigureAwait(false);
        staged.Position = 0;
        return Process(staged, format, request, cancellationToken);
    }

    /// <summary>
    /// 异步输出选定工作表的页面图片。
    /// </summary>
    /// <param name="source">输入工作簿流。</param>
    /// <param name="request">本次处理请求。</param>
    /// <param name="pageSink">异步接收页码、图片流和扩展名的回调。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>已输出页数和字体提示。</returns>
    private async Task<ExcelRenderResult> RenderPagesAsyncCore(Stream source, ExcelRenderRequest request,
        Func<int, Stream, string, Task> pageSink, CancellationToken cancellationToken)
    {
        ValidateStreams(source, null);
        if (request == null) throw new ArgumentNullException(nameof(request));
        ValidateRenderRequest(request, renderPages: true);
        EnsureReady(request.StrictFonts);
        var workbook = LoadWorkbook(source, request.InputFormat, null);
        try
        {
            using var scope = ApplyRenderScope(workbook, request);
            var pageOffset = 0;
            var emittedPages = 0;
            foreach (var sheet in EnumerateWorksheets(workbook, request.SheetName))
            {
                var imageOptions = CreateImageOptions(request);
                var render = CreateSheetRender(sheet, imageOptions);
                var pageCount = System.Convert.ToInt32(GetProperty(render, "PageCount") ?? 0);
                try
                {
                    for (var index = 0; index < pageCount; index++)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var pageNumber = pageOffset + index + 1;
                        if (!IsPageSelected(pageNumber, request))
                            continue;
                        using var image = (IDisposable)Invoke(render, "ToImage", index);
                        await using var buffer = new MemoryStream();
                        SaveImage(image, buffer, request.ImageFormat);
                        buffer.Position = 0;
                        await pageSink(pageNumber, buffer, GetImageExtension(request.ImageFormat))
                            .ConfigureAwait(false);
                        emittedPages++;
                    }
                }
                finally
                {
                    DisposeIfNeeded(render);
                }
                pageOffset += pageCount;
            }
            return new ExcelRenderResult
            {
                PageCount = emittedPages,
                Warnings = BuildFontWarnings(request.StrictFonts)
            };
        }
        finally
        {
            DisposeIfNeeded(workbook);
        }
    }

    /// <summary>
    /// 输出工作表中选定的页面图片。
    /// </summary>
    /// <param name="sheet">待处理的工作表。</param>
    /// <param name="request">本次处理请求。</param>
    /// <param name="sink">接收页码、图片流和扩展名的回调。</param>
    /// <param name="pageOffset">此前工作表的累计页数。</param>
    /// <param name="emittedPages">已输出页面的累计数量。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>工作表的总页数，包含未被选中的页面。</returns>
    private int RenderSheetPages(object sheet, ExcelRenderRequest request,
        Action<int, Stream, string> sink, int pageOffset, ref int emittedPages,
        CancellationToken cancellationToken)
    {
        var imageOptions = CreateImageOptions(request);
        var render = CreateSheetRender(sheet, imageOptions);
        try
        {
            var pageCount = System.Convert.ToInt32(GetProperty(render, "PageCount") ?? 0);
            for (var index = 0; index < pageCount; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var pageNumber = pageOffset + index + 1;
                if (!IsPageSelected(pageNumber, request))
                    continue;
                using var image = (IDisposable)Invoke(render, "ToImage", index);
                using var stream = new MemoryStream();
                SaveImage(image, stream, request.ImageFormat);
                stream.Position = 0;
                sink(pageNumber, new MemoryStream(stream.ToArray()), GetImageExtension(request.ImageFormat));
                emittedPages++;
            }
            return pageCount;
        }
        finally
        {
            DisposeIfNeeded(render);
        }
    }

    /// <summary>
    /// 创建页面图片渲染选项。
    /// </summary>
    /// <param name="request">本次处理请求。</param>
    /// <returns>配置完成的图片渲染选项。</returns>
    private static object CreateImageOptions(ExcelRenderRequest request)
    {
        var type = RequireType("Aspose.Cells.Rendering.ImageOrPrintOptions");
        var options = Activator.CreateInstance(type);
        SetProperty(options, "HorizontalResolution", request.Dpi);
        SetProperty(options, "VerticalResolution", request.Dpi);
        SetProperty(options, "OnePagePerSheet", false);
        if (!string.IsNullOrWhiteSpace(request.Area))
            SetProperty(options, "OnlyArea", true);
        SetProperty(options, "ImageType", ParseImageType(request.ImageFormat));
        return options;
    }

    /// <summary>
    /// 创建 PDF 保存选项。
    /// </summary>
    /// <param name="request">本次处理请求。</param>
    /// <returns>配置完成的 PDF 保存选项。</returns>
    private static object CreatePdfSaveOptions(ExcelRenderRequest request)
    {
        var options = Activator.CreateInstance(RequireType("Aspose.Cells.PdfSaveOptions"));
        if (request.PdfCompliance != ExcelPdfCompliance.None)
        {
            var compliance = Enum.Parse(RequireType("Aspose.Cells.Rendering.PdfCompliance"),
                request.PdfCompliance.ToString(), ignoreCase: false);
            if (!SetProperty(options, "Compliance", compliance))
                throw Unsupported("Aspose.Cells 当前版本不支持 PDF/A 兼容级别。", BingOfficesOperation.Export);
        }

        if (!string.IsNullOrWhiteSpace(request.SheetName))
        {
            var sheetSet = CreateSheetSet(request.SheetName);
            if (!SetProperty(options, "SheetSet", sheetSet))
                throw Unsupported("Aspose.Cells 当前版本不支持按 Sheet 渲染 PDF。", BingOfficesOperation.Export);
        }

        if (request.StartPage.HasValue
            && !SetProperty(options, "PageIndex", request.StartPage.Value - 1))
            throw Unsupported("Aspose.Cells 当前版本不支持 PDF 起始页范围。", BingOfficesOperation.Export);
        if (request.EndPage.HasValue)
        {
            var first = request.StartPage ?? 1;
            if (!SetProperty(options, "PageCount", request.EndPage.Value - first + 1))
                throw Unsupported("Aspose.Cells 当前版本不支持 PDF 页数范围。", BingOfficesOperation.Export);
        }
        return options;
    }

    /// <summary>
    /// 创建指定工作表的渲染集合。
    /// </summary>
    /// <param name="sheetName">待选择的工作表名称。</param>
    /// <returns>只包含指定工作表的渲染集合。</returns>
    private static object CreateSheetSet(string sheetName)
    {
        var type = RequireType("Aspose.Cells.Rendering.SheetSet");
        var constructor = type.GetConstructor(new[] { typeof(string[]) });
        if (constructor == null)
            throw Unsupported("Aspose.Cells 当前版本不支持按 Sheet 渲染。", BingOfficesOperation.Export);
        return constructor.Invoke(new object[] { new[] { sheetName } });
    }

    /// <summary>
    /// 解析 Aspose.Cells 图片格式。
    /// </summary>
    /// <param name="format">目标格式。</param>
    /// <returns>Aspose.Cells 图片格式枚举值。</returns>
    private static object ParseImageType(ExcelRenderFormat format)
    {
        var name = format switch
        {
            ExcelRenderFormat.Jpeg => "Jpeg",
            ExcelRenderFormat.Tiff => "Tiff",
            _ => "Png"
        };
        return Enum.Parse(RequireType("Aspose.Cells.Drawing.ImageType"), name, ignoreCase: false);
    }

    /// <summary>
    /// 计算请求范围内的渲染页数。
    /// </summary>
    /// <param name="workbook">待处理的工作簿。</param>
    /// <param name="request">本次处理请求。</param>
    /// <returns>请求范围内的页数；范围超出文档时为 0。</returns>
    private static int CountRenderedPages(object workbook, ExcelRenderRequest request)
    {
        var render = Activator.CreateInstance(RequireType("Aspose.Cells.Rendering.WorkbookRender"),
            workbook, CreateImageOptions(request));
        try
        {
            var total = System.Convert.ToInt32(GetProperty(render, "PageCount") ?? 0);
            var first = request.StartPage ?? 1;
            var last = request.EndPage ?? total;
            if (last < first || first > total)
                return 0;
            return Math.Min(total, last) - first + 1;
        }
        finally
        {
            DisposeIfNeeded(render);
        }
    }

    /// <summary>
    /// 判断页面是否位于请求的页码范围内。
    /// </summary>
    /// <param name="pageNumber">从 1 开始的页码。</param>
    /// <param name="request">本次处理请求。</param>
    /// <returns>页面位于请求范围内时为 true，否则为 false。</returns>
    private static bool IsPageSelected(int pageNumber, ExcelRenderRequest request)
    {
        var first = request.StartPage ?? 1;
        var last = request.EndPage ?? int.MaxValue;
        return pageNumber >= first && pageNumber <= last;
    }

    /// <summary>
    /// 验证渲染请求的格式和页码范围。
    /// </summary>
    /// <param name="request">本次处理请求。</param>
    /// <param name="renderPages">为 true 时验证页面图片格式；否则验证文档渲染请求。</param>
    private static void ValidateRenderRequest(ExcelRenderRequest request, bool renderPages)
    {
        if (!Enum.IsDefined(typeof(ExcelFormat), request.InputFormat))
            throw new ArgumentOutOfRangeException(nameof(request.InputFormat));
        if (!Enum.IsDefined(typeof(ExcelRenderFormat), request.ImageFormat))
            throw new ArgumentOutOfRangeException(nameof(request.ImageFormat));
        if (!Enum.IsDefined(typeof(ExcelPdfCompliance), request.PdfCompliance))
            throw new ArgumentOutOfRangeException(nameof(request.PdfCompliance));
        if (request.Dpi <= 0)
            throw new ArgumentOutOfRangeException(nameof(request.Dpi));
        if (request.StartPage.HasValue && request.StartPage.Value <= 0)
            throw new ArgumentOutOfRangeException(nameof(request.StartPage));
        if (request.EndPage.HasValue && request.EndPage.Value <= 0)
            throw new ArgumentOutOfRangeException(nameof(request.EndPage));
        if (request.StartPage.HasValue && request.EndPage.HasValue
            && request.EndPage.Value < request.StartPage.Value)
            throw new ArgumentException("结束页不能小于起始页。", nameof(request.EndPage));
        if (renderPages && request.ImageFormat == ExcelRenderFormat.Pdf)
            throw Unsupported("页面渲染必须选择 PNG、JPEG 或 TIFF。", BingOfficesOperation.Export);
    }

    /// <summary>
    /// 临时应用工作表可见性和打印区域。
    /// </summary>
    /// <param name="workbook">待处理的工作簿。</param>
    /// <param name="request">本次处理请求。</param>
    /// <returns>释放时恢复工作表原有状态的作用域。</returns>
    private static RenderScope ApplyRenderScope(object workbook, ExcelRenderRequest request)
    {
        var sheets = EnumerateWorksheets(workbook, null).ToList();
        var selected = string.IsNullOrWhiteSpace(request.SheetName)
            ? sheets
            : sheets.Where(sheet => string.Equals(
                    System.Convert.ToString(GetProperty(sheet, "Name")), request.SheetName,
                    StringComparison.OrdinalIgnoreCase)).ToList();
        if (selected.Count == 0)
            throw Unsupported($"找不到指定 Sheet: {request.SheetName}", BingOfficesOperation.Export);

        var entries = new List<RenderScopeEntry>();
        foreach (var sheet in sheets)
        {
            var selectedSheet = selected.Contains(sheet);
            var entry = new RenderScopeEntry(sheet, GetProperty(sheet, "IsVisible"));
            if (!string.IsNullOrWhiteSpace(request.SheetName))
            {
                if (!SetProperty(sheet, "IsVisible", selectedSheet))
                    throw Unsupported("Aspose.Cells 当前版本不支持按 Sheet 渲染。", BingOfficesOperation.Export);
            }

            if (!string.IsNullOrWhiteSpace(request.Area) && selectedSheet)
            {
                var pageSetup = GetProperty(sheet, "PageSetup");
                entry.PageSetup = pageSetup;
                entry.PrintArea = System.Convert.ToString(GetProperty(pageSetup, "PrintArea"));
                if (!SetProperty(pageSetup, "PrintArea", request.Area))
                    throw Unsupported("Aspose.Cells 当前版本不支持按区域渲染。", BingOfficesOperation.Export);
            }
            entries.Add(entry);
        }
        return new RenderScope(entries);
    }

    /// <summary>
    /// 创建工作表页面渲染器。
    /// </summary>
    /// <param name="sheet">待处理的工作表。</param>
    /// <param name="imageOptions">页面图片渲染选项。</param>
    /// <returns>工作表页面渲染器。</returns>
    private static object CreateSheetRender(object sheet, object imageOptions) =>
        Activator.CreateInstance(RequireType("Aspose.Cells.Rendering.SheetRender"), sheet, imageOptions);

    /// <summary>
    /// 按指定格式保存页面图片。
    /// </summary>
    /// <param name="image">待保存的页面图片。</param>
    /// <param name="destination">接收输出的目标流。</param>
    /// <param name="format">目标格式。</param>
    private static void SaveImage(object image, Stream destination, ExcelRenderFormat format)
    {
        var imageFormatType = Type.GetType("System.Drawing.Imaging.ImageFormat, System.Drawing.Common")
            ?? Type.GetType("System.Drawing.Imaging.ImageFormat, System.Drawing");
        if (imageFormatType == null)
            throw Unsupported("当前运行时未提供 System.Drawing.ImageFormat，无法输出页面图片。", BingOfficesOperation.Export);
        var imageFormat = imageFormatType.GetProperty(ToImageProperty(format), BindingFlags.Public | BindingFlags.Static)
            ?.GetValue(null);
        var method = image.GetType().GetMethod("Save", new[] { typeof(Stream), imageFormatType });
        if (method == null || imageFormat == null)
            throw Unsupported("Aspose.Cells 当前运行时不支持所请求的图片编码。", BingOfficesOperation.Export);
        method.Invoke(image, new[] { destination, imageFormat });
    }

    /// <summary>
    /// 获取图片格式对应的编码属性名。
    /// </summary>
    /// <param name="format">目标格式。</param>
    /// <returns>图片编码对应的静态属性名。</returns>
    private static string ToImageProperty(ExcelRenderFormat format) => format switch
    {
        ExcelRenderFormat.Jpeg => "Jpeg",
        ExcelRenderFormat.Tiff => "Tiff",
        _ => "Png"
    };

    /// <summary>
    /// 获取图片格式对应的文件扩展名。
    /// </summary>
    /// <param name="format">目标格式。</param>
    /// <returns>包含前导句点的图片扩展名。</returns>
    private static string GetImageExtension(ExcelRenderFormat format) => format switch
    {
        ExcelRenderFormat.Jpeg => ".jpg",
        ExcelRenderFormat.Tiff => ".tiff",
        _ => ".png"
    };

    /// <summary>
    /// 读取选定工作表的公式及计算结果。
    /// </summary>
    /// <param name="workbook">待处理的工作簿。</param>
    /// <param name="request">本次处理请求。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    /// <returns>公式单元格的读取及计算结果集合。</returns>
    private IReadOnlyList<ExcelFormulaCellResult> ReadFormulaCells(object workbook,
        ExcelFormulaRequest request, CancellationToken cancellationToken)
    {
        var result = new List<ExcelFormulaCellResult>();
        foreach (var sheet in EnumerateWorksheets(workbook, request.SheetName))
        {
            var cells = GetProperty(sheet, "Cells");
            var maxRow = System.Convert.ToInt32(GetProperty(cells, "MaxDataRow") ?? -1);
            var maxColumn = System.Convert.ToInt32(GetProperty(cells, "MaxDataColumn") ?? -1);
            for (var row = 0; row <= maxRow; row++)
                for (var column = 0; column <= maxColumn; column++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    var cell = GetIndexed(cells, row, column);
                    var formula = System.Convert.ToString(GetProperty(cell, "Formula"));
                    if (string.IsNullOrWhiteSpace(formula)) continue;
                    var isError = System.Convert.ToBoolean(GetProperty(cell, "IsErrorValue") ?? false);
                    var value = GetProperty(cell, "Value");
                    result.Add(new ExcelFormulaCellResult
                    {
                        SheetName = System.Convert.ToString(GetProperty(sheet, "Name")),
                        RowIndex = row,
                        ColumnIndex = column,
                        Formula = request.ReadMode == ExcelFormulaReadMode.CachedValue ? null : formula,
                        CachedValue = request.ReadMode == ExcelFormulaReadMode.FormulaText ? null : value,
                        CalculatedValue = request.CalculationMode == ExcelFormulaCalculationMode.DoNotCalculate
                            ? null : value,
                        ErrorCode = isError ? System.Convert.ToString(value) ?? "FormulaError" : null
                    });
                }
        }
        return result;
    }

    /// <summary>
    /// 枚举符合名称条件的工作表。
    /// </summary>
    /// <param name="workbook">待处理的工作簿。</param>
    /// <param name="requestedSheet">限定的工作表名称；为空时枚举全部工作表。</param>
    /// <returns>符合名称条件的工作表序列。</returns>
    private static IEnumerable<object> EnumerateWorksheets(object workbook, string requestedSheet)
    {
        var worksheets = GetProperty(workbook, "Worksheets");
        var count = System.Convert.ToInt32(GetProperty(worksheets, "Count") ?? 0);
        for (var index = 0; index < count; index++)
        {
            var sheet = GetIndexed(worksheets, index);
            if (string.IsNullOrWhiteSpace(requestedSheet)
                || string.Equals(System.Convert.ToString(GetProperty(sheet, "Name")), requestedSheet,
                    StringComparison.OrdinalIgnoreCase))
                yield return sheet;
        }
    }

    /// <summary>
    /// 检查工作表公式是否包含外部或失效引用。
    /// </summary>
    /// <param name="workbook">待处理的工作簿。</param>
    /// <param name="requestedSheet">限定的工作表名称；为空时枚举全部工作表。</param>
    /// <returns>发现外部或失效引用时为 true，否则为 false。</returns>
    private static bool ContainsUnsupportedFormulaReference(object workbook, string requestedSheet)
    {
        foreach (var sheet in EnumerateWorksheets(workbook, requestedSheet))
        {
            var cells = GetProperty(sheet, "Cells");
            var maxRow = System.Convert.ToInt32(GetProperty(cells, "MaxDataRow") ?? -1);
            var maxColumn = System.Convert.ToInt32(GetProperty(cells, "MaxDataColumn") ?? -1);
            for (var row = 0; row <= maxRow; row++)
                for (var column = 0; column <= maxColumn; column++)
                {
                    var formula = System.Convert.ToString(GetProperty(GetIndexed(cells, row, column), "Formula"));
                    if (string.IsNullOrWhiteSpace(formula)) continue;
                    if (formula.IndexOf('[') >= 0
                        || formula.IndexOf(']') >= 0
                        || formula.IndexOf("#REF!", StringComparison.OrdinalIgnoreCase) >= 0)
                        return true;
                }
        }
        return false;
    }

    /// <summary>
    /// 加载工作簿并应用打开密码。
    /// </summary>
    /// <param name="source">输入工作簿流。</param>
    /// <param name="format">输入工作簿格式。</param>
    /// <param name="password">工作簿密码；为空时不设置密码。</param>
    /// <returns>加载完成的工作簿。</returns>
    private object LoadWorkbook(Stream source, ExcelFormat format, string password)
    {
        var workbookType = RequireType("Aspose.Cells.Workbook");
        if (string.IsNullOrEmpty(password))
            return Activator.CreateInstance(workbookType, source);
        var loadOptionsType = RequireType("Aspose.Cells.LoadOptions");
        var loadOptions = Activator.CreateInstance(loadOptionsType);
        SetProperty(loadOptions, "Password", password);
        SetProperty(loadOptions, "LoadFormat", ParseLoadFormat(format));
        return Activator.CreateInstance(workbookType, source, loadOptions);
    }

    /// <summary>
    /// 按指定格式和密码保存工作簿。
    /// </summary>
    /// <param name="workbook">待处理的工作簿。</param>
    /// <param name="destination">接收输出的目标流。</param>
    /// <param name="format">目标格式。</param>
    /// <param name="password">工作簿密码；为空时不设置密码。</param>
    /// <param name="cancellationToken">取消操作的令牌。</param>
    private void SaveWorkbook(object workbook, Stream destination, string format,
        string password, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (string.IsNullOrEmpty(password))
        {
            Invoke(workbook, "Save", destination, ParseSaveFormat(format));
            return;
        }
        var saveOptions = CreateSaveOptions(format, password);
        Invoke(workbook, "Save", destination, saveOptions);
    }

    /// <summary>
    /// 创建包含保存密码的格式选项。
    /// </summary>
    /// <param name="format">目标格式。</param>
    /// <param name="password">工作簿密码；为空时不设置密码。</param>
    /// <returns>包含保存密码的格式选项。</returns>
    private static object CreateSaveOptions(string format, string password)
    {
        var typeName = format switch
        {
            "Xls" => "Aspose.Cells.XlsSaveOptions",
            "Ods" => "Aspose.Cells.OdsSaveOptions",
            _ => "Aspose.Cells.OoxmlSaveOptions"
        };
        var options = Activator.CreateInstance(RequireType(typeName));
        if (!SetProperty(options, "Password", password))
            throw Unsupported("Aspose.Cells 当前保存格式不支持密码写入。", BingOfficesOperation.Export);
        return options;
    }

    /// <summary>
    /// 确保许可证和字体配置可用于当前操作。
    /// </summary>
    /// <param name="strictFonts">是否要求配置至少一个字体目录。</param>
    private void EnsureReady(bool strictFonts)
    {
        lock (_gate)
        {
            if (!_licenseConfigured)
            {
                if (string.IsNullOrWhiteSpace(_options.LicensePath)
                    || !File.Exists(_options.LicensePath))
                    throw Unsupported("Aspose.Cells 需要宿主配置有效许可证文件。", BingOfficesOperation.Export);
                var license = Activator.CreateInstance(RequireType("Aspose.Cells.License"));
                Invoke(license, "SetLicense", _options.LicensePath);
                ConfigureFonts();
                _licenseConfigured = true;
            }
        }
        if (strictFonts || _options.StrictFonts)
        {
            if ((_options.FontDirectories?.Count ?? 0) == 0)
                throw Unsupported("严格字体模式必须配置至少一个字体目录。", BingOfficesOperation.Export);
        }
    }

    /// <summary>
    /// 注册宿主配置的字体目录和替代映射。
    /// </summary>
    private void ConfigureFonts()
    {
        var fontConfigs = RequireType("Aspose.Cells.FontConfigs");
        foreach (var directory in _options.FontDirectories ?? Array.Empty<string>())
        {
            if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
                throw new DirectoryNotFoundException(directory);
            InvokeStatic(fontConfigs, "SetFontFolder", directory, true);
        }
        foreach (var pair in _options.FontSubstitutions ?? new Dictionary<string, string>())
            InvokeStatic(fontConfigs, "SetFontSubstitutes", pair.Key, new[] { pair.Value });
    }

    /// <summary>
    /// 生成已配置字体替代的提示。
    /// </summary>
    /// <param name="strict">是否启用严格字体模式。</param>
    /// <returns>字体替代提示集合；严格模式或无替代配置时为空。</returns>
    private IReadOnlyList<ExcelRenderWarning> BuildFontWarnings(bool strict)
    {
        if (strict || _options.StrictFonts || (_options.FontSubstitutions?.Count ?? 0) == 0)
            return Array.Empty<ExcelRenderWarning>();
        return _options.FontSubstitutions.Select(pair => new ExcelRenderWarning
        {
            Code = "FontSubstitutionConfigured",
            Message = $"字体 {pair.Key} 使用宿主配置的替代字体 {pair.Value}。"
        }).ToArray();
    }

    /// <summary>
    /// 应用工作簿宏保存策略。
    /// </summary>
    /// <param name="workbook">待处理的工作簿。</param>
    /// <param name="policy">宏保存策略。</param>
    /// <param name="targetFormat">输出工作簿格式。</param>
    private static void ApplyMacroPolicy(object workbook, ExcelMacroPolicy policy, ExcelFormat targetFormat)
    {
        if (policy == ExcelMacroPolicy.Preserve && targetFormat != ExcelFormat.Xlsm)
            throw Unsupported("宏保留只允许输出 XLSM。", BingOfficesOperation.Export);
        if (policy != ExcelMacroPolicy.Strip) return;
        var method = workbook.GetType().GetMethod("RemoveMacro", BindingFlags.Public | BindingFlags.Instance,
            binder: null, types: Type.EmptyTypes, modifiers: null);
        if (method == null)
            throw Unsupported("Aspose.Cells 当前版本无法按请求剥离 VBA 项目。", BingOfficesOperation.Export);
        try
        {
            method.Invoke(workbook, null);
        }
        catch (TargetInvocationException exception)
        {
            throw Unsupported("Aspose.Cells 剥离 VBA 项目失败。", BingOfficesOperation.Export,
                exception.InnerException ?? exception);
        }
    }

    /// <summary>
    /// 验证源格式和目标格式的宏策略。
    /// </summary>
    /// <param name="source">输入工作簿格式。</param>
    /// <param name="target">输出工作簿格式。</param>
    /// <param name="openPolicy">打开工作簿时的宏策略。</param>
    /// <param name="savePolicy">保存工作簿时的宏策略。</param>
    private static void ValidateMacroPolicy(ExcelFormat source, ExcelFormat target,
        ExcelMacroPolicy openPolicy, ExcelMacroPolicy savePolicy)
    {
        if (!Enum.IsDefined(typeof(ExcelMacroPolicy), openPolicy)
            || !Enum.IsDefined(typeof(ExcelMacroPolicy), savePolicy))
            throw new ArgumentOutOfRangeException(nameof(openPolicy));
        if (source == ExcelFormat.Xlsm && openPolicy == ExcelMacroPolicy.Reject)
            throw Unsupported("打开 XLSM 必须显式选择 Preserve 或 Strip。", BingOfficesOperation.Import);
        if (savePolicy == ExcelMacroPolicy.Preserve && target != ExcelFormat.Xlsm)
            throw Unsupported("宏保留只允许输出 XLSM。", BingOfficesOperation.Export);
    }

    /// <summary>
    /// 验证提供程序是否支持指定工作簿格式。
    /// </summary>
    /// <param name="format">目标格式。</param>
    private static void ValidateFormat(ExcelFormat format)
    {
        if (format != ExcelFormat.Xls && format != ExcelFormat.Xlsx && format != ExcelFormat.Xlsm
            && format != ExcelFormat.Ods)
            throw Unsupported("Aspose.Cells Provider 不支持请求的 Excel 格式。", BingOfficesOperation.Import);
    }

    /// <summary>
    /// 验证公式读取和计算模式。
    /// </summary>
    /// <param name="request">本次处理请求。</param>
    private static void ValidateFormulaRequest(ExcelFormulaRequest request)
    {
        if (!Enum.IsDefined(typeof(ExcelFormulaReadMode), request.ReadMode))
            throw new ArgumentOutOfRangeException(nameof(request.ReadMode));
        if (!Enum.IsDefined(typeof(ExcelFormulaCalculationMode), request.CalculationMode))
            throw new ArgumentOutOfRangeException(nameof(request.CalculationMode));
    }

    /// <summary>
    /// 获取工作簿格式对应的 Aspose.Cells 格式名。
    /// </summary>
    /// <param name="format">目标格式。</param>
    /// <returns>Aspose.Cells 格式名称。</returns>
    private static string ToSaveFormat(ExcelFormat format) => format switch
    {
        ExcelFormat.Xls => "Xls",
        ExcelFormat.Xlsm => "Xlsm",
        ExcelFormat.Ods => "Ods",
        _ => "Xlsx"
    };

    /// <summary>
    /// 解析 Aspose.Cells 保存格式。
    /// </summary>
    /// <param name="format">目标格式。</param>
    /// <returns>Aspose.Cells 保存格式枚举值。</returns>
    private static object ParseSaveFormat(ExcelFormat format) => ParseSaveFormat(ToSaveFormat(format));
    /// <summary>
    /// 解析 Aspose.Cells 保存格式。
    /// </summary>
    /// <param name="format">目标格式。</param>
    /// <returns>Aspose.Cells 保存格式枚举值。</returns>
    private static object ParseSaveFormat(string format) => Enum.Parse(RequireType("Aspose.Cells.SaveFormat"), format, true);
    /// <summary>
    /// 解析 Aspose.Cells 加载格式。
    /// </summary>
    /// <param name="format">输入工作簿格式。</param>
    /// <returns>Aspose.Cells 加载格式枚举值。</returns>
    private static object ParseLoadFormat(ExcelFormat format) =>
        Enum.Parse(RequireType("Aspose.Cells.LoadFormat"), ToSaveFormat(format), true);

    /// <summary>
    /// 获取工作簿的工作表数量。
    /// </summary>
    /// <param name="workbook">待处理的工作簿。</param>
    /// <returns>工作表数量。</returns>
    private static int GetWorksheetCount(object workbook) =>
        System.Convert.ToInt32(GetProperty(GetProperty(workbook, "Worksheets"), "Count") ?? 0);

    /// <summary>
    /// 验证输入流和可选目标流的读写能力。
    /// </summary>
    /// <param name="source">输入工作簿流。</param>
    /// <param name="destination">接收输出的目标流。</param>
    private static void ValidateStreams(Stream source, Stream destination)
    {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (!source.CanRead) throw new ArgumentException("输入流不可读取。", nameof(source));
        if (destination != null && !destination.CanWrite)
            throw new ArgumentException("目标流不可写入。", nameof(destination));
    }

    /// <summary>
    /// 加载所需的 Aspose.Cells 运行时类型。
    /// </summary>
    /// <param name="typeName">不含程序集名称的类型全名。</param>
    /// <returns>已加载的运行时类型。</returns>
    private static Type RequireType(string typeName)
    {
        var type = Type.GetType(typeName + ", Aspose.Cells", throwOnError: false);
        if (type == null)
            throw Unsupported("未加载 Aspose.Cells 26.8.0 运行时程序集。", BingOfficesOperation.Import);
        return type;
    }

    /// <summary>
    /// 读取对象的公共实例属性。
    /// </summary>
    /// <param name="target">目标对象。</param>
    /// <param name="name">成员名称。</param>
    /// <returns>属性值；对象或属性不存在、或属性本身为空时为 null。</returns>
    private static object GetProperty(object target, string name) =>
        target?.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.Instance)?.GetValue(target);

    /// <summary>
    /// 读取对象的索引属性值。
    /// </summary>
    /// <param name="target">目标对象。</param>
    /// <param name="indexes">索引属性的参数值。</param>
    /// <returns>索引属性值；目标值为空时为 null。</returns>
    private static object GetIndexed(object target, params object[] indexes)
    {
        var property = target.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .FirstOrDefault(item => item.Name == "Item" && item.GetIndexParameters().Length == indexes.Length);
        if (property == null) throw new MissingMemberException(target.GetType().FullName, "Item");
        return property.GetValue(target, indexes);
    }

    /// <summary>
    /// 调用参数类型匹配的公共实例方法。
    /// </summary>
    /// <param name="target">目标对象。</param>
    /// <param name="name">成员名称。</param>
    /// <param name="arguments">方法调用的实参。</param>
    /// <returns>调用结果；无返回值或返回值为空时为 null。</returns>
    private static object Invoke(object target, string name, params object[] arguments)
    {
        var methods = target.GetType().GetMethods(BindingFlags.Public | BindingFlags.Instance)
            .Where(method => method.Name == name && method.GetParameters().Length == arguments.Length);
        foreach (var method in methods)
        {
            var parameters = method.GetParameters();
            if (parameters.Where((parameter, index) => arguments[index] != null
                    && !parameter.ParameterType.IsInstanceOfType(arguments[index]))
                .Any()) continue;
            return method.Invoke(target, arguments);
        }
        throw new MissingMethodException(target.GetType().FullName, name);
    }

    /// <summary>
    /// 调用参数类型匹配的公共静态方法。
    /// </summary>
    /// <param name="type">包含目标静态方法的类型。</param>
    /// <param name="name">成员名称。</param>
    /// <param name="arguments">方法调用的实参。</param>
    private static void InvokeStatic(Type type, string name, params object[] arguments)
    {
        var methods = type.GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Where(method => method.Name == name && method.GetParameters().Length == arguments.Length);
        foreach (var method in methods)
        {
            var parameters = method.GetParameters();
            if (parameters.Where((parameter, index) => arguments[index] != null
                    && !parameter.ParameterType.IsInstanceOfType(arguments[index]))
                .Any()) continue;
            method.Invoke(null, arguments);
            return;
        }
        throw new MissingMethodException(type.FullName, name);
    }

    /// <summary>
    /// 写入对象的可写公共实例属性。
    /// </summary>
    /// <param name="target">目标对象。</param>
    /// <param name="name">成员名称。</param>
    /// <param name="value">待处理的对象值。</param>
    /// <returns>成功写入时为 true；属性不存在或不可写时为 false。</returns>
    private static bool SetProperty(object target, string name, object value)
    {
        var property = target?.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
        if (property == null || !property.CanWrite) return false;
        property.SetValue(target, value);
        return true;
    }

    /// <summary>
    /// 释放实现可释放接口的对象。
    /// </summary>
    /// <param name="value">待处理的对象值。</param>
    private static void DisposeIfNeeded(object value) => (value as IDisposable)?.Dispose();

    /// <summary>
    /// 用于恢复渲染前工作表状态的作用域。
    /// </summary>
    private sealed class RenderScope : IDisposable
    {
        /// <summary>
        /// 渲染前保存的工作表状态集合。
        /// </summary>
        private readonly IReadOnlyList<RenderScopeEntry> _entries;
        /// <summary>
        /// 指示作用域是否已恢复工作表状态。
        /// </summary>
        private bool _disposed;

        /// <summary>
        /// 初始化一个 <see cref="RenderScope"/> 类型的实例。
        /// </summary>
        /// <param name="entries">渲染前的工作表状态集合。</param>
        public RenderScope(IReadOnlyList<RenderScopeEntry> entries) => _entries = entries;

        /// <inheritdoc />
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            foreach (var entry in _entries)
            {
                if (entry.PageSetup != null)
                    SetProperty(entry.PageSetup, "PrintArea", entry.PrintArea);
                if (entry.Visible != null)
                    SetProperty(entry.Sheet, "IsVisible", entry.Visible);
            }
        }
    }

    /// <summary>
    /// 渲染前工作表可见性和打印区域的快照。
    /// </summary>
    private sealed class RenderScopeEntry
    {
        /// <summary>
        /// 初始化一个 <see cref="RenderScopeEntry"/> 类型的实例。
        /// </summary>
        /// <param name="sheet">待处理的工作表。</param>
        /// <param name="visible">工作表原有的可见性。</param>
        public RenderScopeEntry(object sheet, object visible)
        {
            Sheet = sheet;
            Visible = visible;
        }

        /// <summary>
        /// 获取待恢复状态的工作表。
        /// </summary>
        public object Sheet { get; }
        /// <summary>
        /// 获取渲染前的工作表可见性。
        /// </summary>
        public object Visible { get; }
        /// <summary>
        /// 获取或设置需要恢复的页面设置对象。
        /// </summary>
        public object PageSetup { get; set; }
        /// <summary>
        /// 获取或设置渲染前的打印区域。
        /// </summary>
        public string PrintArea { get; set; }
    }

    /// <summary>
    /// 创建提供程序不支持指定能力的异常。
    /// </summary>
    /// <param name="message">错误说明。</param>
    /// <param name="operation">发生错误的操作类型。</param>
    /// <param name="innerException">原始异常；没有原始异常时为 null。</param>
    /// <returns>包含提供程序和预检阶段信息的异常。</returns>
    private static BingOfficesUnsupportedFeatureException Unsupported(string message, BingOfficesOperation operation,
        Exception innerException = null) =>
        new(message, provider: Provider, operation: operation, stage: BingOfficesStage.Preflight,
            innerException: innerException);
}
