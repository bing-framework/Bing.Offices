using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Bing.Offices.Exports;

/// <summary>
/// Workbook 导出请求构建入口。
/// </summary>
public static class ExcelExport
{
    /// <summary>
    /// 创建 Workbook 导出请求。
    /// </summary>
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
    private readonly List<ExcelSheetExportRequest> _sheets = new List<ExcelSheetExportRequest>();
    private Stream _template;
    private bool _leaveTemplateOpen;
    private ExcelFormat _format = ExcelFormat.Xlsx;

    /// <summary>
    /// 设置输出格式。
    /// </summary>
    public ExcelWorkbookExportBuilder Format(ExcelFormat format)
    {
        _format = format;
        return this;
    }

    /// <summary>
    /// 使用已有模板作为 Workbook 来源。
    /// </summary>
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
    /// 添加导航集合 Sheet。该方法等价于一次 SelectMany 后的强类型 AddSheet。
    /// </summary>
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
        return new ExcelWorkbookExportRequest(_sheets.AsReadOnly(), _template, _leaveTemplateOpen, _format);
    }
}
