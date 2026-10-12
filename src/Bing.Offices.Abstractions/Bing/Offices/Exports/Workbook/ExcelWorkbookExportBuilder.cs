namespace Bing.Offices.Exports;

/// <summary>
/// 工作簿导出构建器。
/// </summary>
public sealed class ExcelWorkbookExportBuilder
{
    /// <summary>
    /// 按配置顺序保存待导出的工作表请求。
    /// </summary>
    private readonly List<ExcelSheetExportRequest> _sheets = new List<ExcelSheetExportRequest>();
    /// <summary>
    /// 按配置顺序保存文本字典。
    /// </summary>
    private readonly List<ExcelLookupListDefinition> _lookupLists = new List<ExcelLookupListDefinition>();
    /// <summary>
    /// 按配置顺序保存级联字典。
    /// </summary>
    private readonly List<ExcelCascadeListDefinition> _cascadeLists = new List<ExcelCascadeListDefinition>();
    /// <summary>
    /// 按配置顺序保存透视表定义。
    /// </summary>
    private readonly List<ExcelPivotTableDefinition> _pivotTables = new();
    /// <summary>
    /// 按配置顺序保存已有透视的更新定义。
    /// </summary>
    private readonly List<ExcelPivotTableDefinition> _pivotTableUpdates = new();
    /// <summary>
    /// 是否显式保留模板中的已有透视。
    /// </summary>
    private bool _preserveTemplatePivotTables;
    /// <summary>
    /// 导出时使用的模板流；未设置时从空工作簿创建。
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
    /// 待写入工作簿的元数据；未设置时不覆盖元数据。
    /// </summary>
    private ExcelWorkbookMetadataOptions _metadata;
    /// <summary>
    /// 模板预检使用的资源限制；未设置时沿用默认值。
    /// </summary>
    private Bing.Offices.Imports.ExcelResourceLimits _resourceLimits;

    /// <summary>
    /// 设置输出格式。
    /// </summary>
    /// <param name="format">要生成的 Excel 文件格式。</param>
    /// <returns>当前构建器，用于继续配置工作簿。</returns>
    public ExcelWorkbookExportBuilder Format(ExcelFormat format)
    {
        _format = format;
        return this;
    }

    /// <summary>
    /// 设置工作簿元数据。
    /// </summary>
    /// <param name="metadata">要写入工作簿的元数据；不能为 null。</param>
    /// <returns>当前构建器，用于继续配置工作簿。</returns>
    public ExcelWorkbookExportBuilder Metadata(ExcelWorkbookMetadataOptions metadata)
    {
        _metadata = metadata ?? throw new ArgumentNullException(nameof(metadata));
        return this;
    }

    /// <summary>
    /// 设置导出模板预检使用的资源限制。
    /// </summary>
    /// <param name="limits">模板输入、ZIP 和 XML 处理使用的资源限制；传入 null 时使用默认值。</param>
    /// <returns>当前构建器，用于继续配置工作簿。</returns>
    /// <remarks>该配置主要约束透视模板保留和更新路径，普通无模板导出不额外执行模板预检。</remarks>
    public ExcelWorkbookExportBuilder ResourceLimits(Bing.Offices.Imports.ExcelResourceLimits limits)
    {
        _resourceLimits = limits;
        return this;
    }

    /// <summary>
    /// 使用已有模板作为工作簿来源。
    /// </summary>
    /// <param name="templateStream">可读取的模板流。</param>
    /// <param name="leaveOpen">完成导出后是否保持模板流打开。</param>
    /// <returns>当前构建器，用于继续配置工作簿。</returns>
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
    /// 添加一个强类型工作表。
    /// </summary>
    /// <typeparam name="T">Sheet 数据项类型。</typeparam>
    /// <param name="name">工作表名称。</param>
    /// <param name="data">要写入工作表的数据集合。</param>
    /// <param name="configure">用于配置当前工作表的可选委托。</param>
    /// <returns>当前构建器，用于继续配置工作簿。</returns>
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
    /// 添加工作簿文本字典。
    /// </summary>
    /// <param name="name">供列表校验引用的工作簿级名称。</param>
    /// <param name="values">按显示顺序排列的文本值。</param>
    /// <returns>当前构建器。</returns>
    public ExcelWorkbookExportBuilder LookupList(string name, IReadOnlyList<string> values)
    {
        var definition = new ExcelLookupListDefinition(name, values);
        if (_lookupLists.Any(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase))
            || _cascadeLists.Any(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException("工作簿包含重复字典名称。", nameof(name));
        _lookupLists.Add(definition);
        return this;
    }

    /// <summary>
    /// 添加两级联动文本字典。
    /// </summary>
    /// <param name="name">上级列表的工作簿名称。</param>
    /// <param name="values">上级文本与下级列表。</param>
    /// <returns>当前构建器。</returns>
    public ExcelWorkbookExportBuilder CascadeList(string name,
        IReadOnlyDictionary<string, IReadOnlyList<string>> values)
    {
        var definition = new ExcelCascadeListDefinition(name, values);
        if (_lookupLists.Any(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase))
            || _cascadeLists.Any(item => string.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException("工作簿包含重复字典名称。", nameof(name));
        _cascadeLists.Add(definition);
        return this;
    }

    /// <summary>
    /// 添加工作簿透视表。
    /// </summary>
    /// <param name="definition">透视表定义。</param>
    /// <returns>当前构建器。</returns>
    public ExcelWorkbookExportBuilder PivotTable(ExcelPivotTableDefinition definition)
    {
        if (definition == null)
            throw new ArgumentNullException(nameof(definition));
        definition.Validate();
        if (_pivotTableUpdates.Any(item => string.Equals(item.Name, definition.Name,
                StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException("新建与更新透视表不能同名。", nameof(definition));
        _pivotTables.Add(definition.Clone());
        return this;
    }

    /// <summary>
    /// 保留模板中的已有透视表。
    /// </summary>
    /// <returns>当前构建器。</returns>
    /// <remarks>要求 XLSX 模板；保留模式不自动更新缓存，刷新来源应使用更新入口。</remarks>
    public ExcelWorkbookExportBuilder PreserveTemplatePivotTables()
    {
        _preserveTemplatePivotTables = true;
        return this;
    }

    /// <summary>
    /// 更新模板中的已有透视表。
    /// </summary>
    /// <param name="definition">按名称和目标工作表定位的更新定义。</param>
    /// <returns>当前构建器。</returns>
    /// <remarks>目标存在性在模板预检阶段确认，不隐式新建；未指定更新的模板透视保持原样。</remarks>
    public ExcelWorkbookExportBuilder UpdatePivotTable(ExcelPivotTableDefinition definition)
    {
        if (definition == null)
            throw new ArgumentNullException(nameof(definition));
        definition.Validate();
        var snapshot = definition.Clone();
        if (_pivotTables.Concat(_pivotTableUpdates).Any(item => string.Equals(item.Name,
                snapshot.Name, StringComparison.OrdinalIgnoreCase)))
            throw new ArgumentException("新建或更新透视表名称重复。", nameof(definition));
        _pivotTableUpdates.Add(snapshot);
        _preserveTemplatePivotTables = true;
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
    /// <returns>当前构建器，用于继续配置工作簿。</returns>
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
    /// 验证并生成不可变工作簿导出请求。
    /// </summary>
    /// <returns>已完成校验的工作簿导出请求。</returns>
    internal ExcelWorkbookExportRequest Build()
    {
        if (_preserveTemplatePivotTables && _template == null)
            throw new InvalidOperationException("保留或更新透视表必须使用模板。");
        if (_sheets.Count == 0 && !_preserveTemplatePivotTables)
            throw new InvalidOperationException("Workbook 至少需要一个 Sheet。");
        if (!Enum.IsDefined(typeof(ExcelFormat), _format))
            throw new ArgumentOutOfRangeException(nameof(_format));
        _resourceLimits?.Validate();
        if (_preserveTemplatePivotTables && _format != ExcelFormat.Xlsx)
            throw new ArgumentException("保留或更新透视模板仅支持 XLSX。", nameof(_format));
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in _sheets)
        {
            if (!names.Add(sheet.Name))
                throw new ArgumentException($"Workbook 包含重复 Sheet 名称: {sheet.Name}");
        }
        var lookupNames = new HashSet<string>(_lookupLists.Select(item => item.Name)
                .Concat(_cascadeLists.Select(item => item.Name)),
            StringComparer.OrdinalIgnoreCase);
        foreach (var namedRange in _sheets.SelectMany(sheet => sheet.NamedRanges))
            if (lookupNames.Contains(namedRange.Name))
                throw new ArgumentException($"字典与名称范围重名: {namedRange.Name}");
        var sourceNames = new HashSet<string>(names, StringComparer.OrdinalIgnoreCase);
        var pivotNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var pivot in _pivotTables)
        {
            if (!sourceNames.Contains(pivot.SourceSheet))
                throw new ArgumentException($"透视表来源 Sheet 不存在: {pivot.SourceSheet}");
            if (!names.Add(pivot.TargetSheet))
                throw new ArgumentException($"透视表目标 Sheet 重名: {pivot.TargetSheet}");
            if (!pivotNames.Add(pivot.Name))
                throw new ArgumentException($"透视表名称重复: {pivot.Name}");
        }
        return new ExcelWorkbookExportRequest(_sheets.AsReadOnly(), _template, _leaveTemplateOpen, _format, _metadata,
            _metadata != null, _lookupLists.AsReadOnly(), _cascadeLists.AsReadOnly(), _pivotTables.AsReadOnly(),
            _preserveTemplatePivotTables, _pivotTableUpdates.AsReadOnly(), _resourceLimits);
    }
}
