namespace Bing.Offices.Exports;

/// <summary>
/// Workbook 导出请求构建入口。
/// </summary>
public static class ExcelExport
{
    /// <summary>
    /// 创建 Workbook 导出请求。
    /// </summary>
    /// <param name="configure">用于配置 Workbook 导出选项的委托。</param>
    /// <returns>已完成配置的 Workbook 导出请求。</returns>
    public static ExcelWorkbookExportRequest Workbook(Action<ExcelWorkbookExportBuilder> configure)
    {
        if (configure == null)
            throw new ArgumentNullException(nameof(configure));
        var builder = new ExcelWorkbookExportBuilder();
        configure(builder);
        return builder.Build();
    }
}

/// <summary>
/// Workbook 导出构建器。
/// </summary>
public sealed class ExcelWorkbookExportBuilder
{
    /// <summary>
    /// 按配置顺序保存待导出的工作表请求。
    /// </summary>
    private readonly List<ExcelSheetExportRequest> _sheets = new List<ExcelSheetExportRequest>();
    /// <summary>
    /// 导出时使用的模板流；未设置时从空 Workbook 创建。
    /// </summary>
    private Stream _template;
    /// <summary>
    /// 导出完成后是否保持模板流打开。
    /// </summary>
    private bool _leaveTemplateOpen;
    /// <summary>
    /// 目标 Excel 文件格式，默认为 Xlsx。
    /// </summary>
    private ExcelFormat _format = ExcelFormat.Xlsx;
    /// <summary>
    /// 待写入 Workbook 的元数据；未设置时不覆盖元数据。
    /// </summary>
    private ExcelWorkbookMetadataOptions _metadata;

    /// <summary>
    /// 设置输出格式。
    /// </summary>
    /// <param name="format">要生成的 Excel 文件格式。</param>
    /// <returns>当前构建器，用于继续配置 Workbook。</returns>
    public ExcelWorkbookExportBuilder Format(ExcelFormat format)
    {
        _format = format;
        return this;
    }

    /// <summary>
    /// 设置 Workbook 元数据。
    /// </summary>
    /// <param name="metadata">要写入 Workbook 的元数据；不能为 null。</param>
    /// <returns>当前构建器，用于继续配置 Workbook。</returns>
    public ExcelWorkbookExportBuilder Metadata(ExcelWorkbookMetadataOptions metadata)
    {
        _metadata = metadata ?? throw new ArgumentNullException(nameof(metadata));
        return this;
    }

    /// <summary>
    /// 使用已有模板作为 Workbook 来源。
    /// </summary>
    /// <param name="templateStream">可读取的模板流。</param>
    /// <param name="leaveOpen">完成导出后是否保持模板流打开。</param>
    /// <returns>当前构建器，用于继续配置 Workbook。</returns>
    public ExcelWorkbookExportBuilder UseTemplate(Stream templateStream, bool leaveOpen = false)
    {
        if (templateStream == null)
            throw new ArgumentNullException(nameof(templateStream));
        if (!templateStream.CanRead)
            throw new ArgumentException("模板流不可读取。", nameof(templateStream));
        _template = templateStream;
        _leaveTemplateOpen = leaveOpen;
        return this;
    }

    /// <summary>
    /// 添加一个强类型 Sheet。
    /// </summary>
    /// <typeparam name="T">Sheet 数据项类型。</typeparam>
    /// <param name="name">工作表名称。</param>
    /// <param name="data">要写入工作表的数据集合。</param>
    /// <param name="configure">用于配置当前 Sheet 的可选委托。</param>
    /// <returns>当前构建器，用于继续配置 Workbook。</returns>
    public ExcelWorkbookExportBuilder AddSheet<T>(string name, IEnumerable<T> data,
        Action<ExcelSheetExportBuilder<T>> configure = null) where T : class, new()
    {
        if (data == null)
            throw new ArgumentNullException(nameof(data));
        var builder = new ExcelSheetExportBuilder<T>(name, data);
        configure?.Invoke(builder);
        _sheets.Add(builder.Build());
        return this;
    }

    /// <summary>
    /// 添加导航集合 Sheet。
    /// </summary>
    /// <remarks>该方法等价于一次 SelectMany 后的强类型 AddSheet。</remarks>
    /// <typeparam name="TParent">父实体类型。</typeparam>
    /// <typeparam name="TChild">导航集合中的子实体类型。</typeparam>
    /// <param name="name">工作表名称。</param>
    /// <param name="parents">父实体集合。</param>
    /// <param name="navigation">从父实体获取子实体集合的函数。</param>
    /// <param name="configure">用于配置当前 Sheet 的可选委托。</param>
    /// <returns>当前构建器，用于继续配置 Workbook。</returns>
    public ExcelWorkbookExportBuilder AddNavigationSheet<TParent, TChild>(string name, IEnumerable<TParent> parents,
        Func<TParent, IEnumerable<TChild>> navigation,
        Action<ExcelSheetExportBuilder<TChild>> configure = null)
        where TParent : class
        where TChild : class, new()
    {
        if (parents == null)
            throw new ArgumentNullException(nameof(parents));
        if (navigation == null)
            throw new ArgumentNullException(nameof(navigation));
        var data = parents.SelectMany(parent => navigation(parent) ?? Array.Empty<TChild>());
        return AddSheet(name, data, configure);
    }

    /// <summary>
    /// 验证并生成不可变 Workbook 导出请求。
    /// </summary>
    /// <returns>已完成校验的 Workbook 导出请求。</returns>
    internal ExcelWorkbookExportRequest Build()
    {
        if (_sheets.Count == 0)
            throw new InvalidOperationException("Workbook 至少需要一个 Sheet。");
        if (!Enum.IsDefined(typeof(ExcelFormat), _format))
            throw new ArgumentOutOfRangeException(nameof(_format));
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in _sheets)
        {
            if (!names.Add(sheet.Name))
                throw new ArgumentException($"Workbook 包含重复 Sheet 名称: {sheet.Name}");
        }
        return new ExcelWorkbookExportRequest(_sheets.AsReadOnly(), _template, _leaveTemplateOpen, _format, _metadata,
            _metadata != null);
    }
}
