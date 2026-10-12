using Bing.Offices.Exceptions;
using Bing.Offices.Exports;

namespace Bing.Offices.Internals;

/// <summary>
/// 定位透视模板更新并校验工作簿名称空间。
/// </summary>
internal sealed class ExcelPivotTemplatePlan
{
    /// <summary>
    /// 初始化一个 <see cref="ExcelPivotTemplatePlan" /> 类型的实例。
    /// </summary>
    /// <param name="manifest">模板依赖清单。</param>
    /// <param name="updates">已定位的更新。</param>
    /// <param name="newPivots">本次请求中新建的透视定义。</param>
    private ExcelPivotTemplatePlan(ExcelPivotTemplateManifest manifest,
        IEnumerable<ExcelPivotTemplateUpdate> updates,
        IEnumerable<ExcelPivotTableDefinition> newPivots)
    {
        Manifest = manifest;
        Updates = Array.AsReadOnly(updates.ToArray());
        NewPivots = Array.AsReadOnly(newPivots.ToArray());
        var updated = new HashSet<string>(Updates.Select(item => item.Entry.PartName), StringComparer.OrdinalIgnoreCase);
        PreservedEntries = Array.AsReadOnly(manifest.Entries.Where(item => !updated.Contains(item.PartName)).ToArray());
    }

    /// <summary>
    /// 获取模板清单。
    /// </summary>
    internal ExcelPivotTemplateManifest Manifest { get; }

    /// <summary>
    /// 获取已定位的更新。
    /// </summary>
    internal IReadOnlyList<ExcelPivotTemplateUpdate> Updates { get; }

    /// <summary>
    /// 获取本次请求中新建的透视名称。
    /// </summary>
    internal IReadOnlyList<ExcelPivotTableDefinition> NewPivots { get; }

    /// <summary>
    /// 获取保持原定义的透视表。
    /// </summary>
    internal IReadOnlyList<ExcelPivotTemplateEntry> PreservedEntries { get; }

    /// <summary>
    /// 生成透视模板执行计划。
    /// </summary>
    /// <param name="request">不可变工作簿请求。</param>
    /// <param name="manifest">实际模板清单。</param>
    /// <param name="provider">提供程序名称。</param>
    /// <param name="token">取消令牌。</param>
    /// <returns>全部名称与目标定位通过后的计划。</returns>
    /// <remarks>实际来源类型、展开容量及目标占用由 Provider 在业务数据写入后继续预检。</remarks>
    internal static ExcelPivotTemplatePlan Create(ExcelWorkbookExportRequest request,
        ExcelPivotTemplateManifest manifest, string provider, CancellationToken token = default)
    {
        if (request == null)
            throw new ArgumentNullException(nameof(request));
        if (manifest == null)
            throw new ArgumentNullException(nameof(manifest));
        token.ThrowIfCancellationRequested();
        if (!request.PreserveTemplatePivotTables || request.Template == null || request.Format != ExcelFormat.Xlsx)
            throw Configuration("透视模板执行计划要求显式 XLSX 模板模式。", provider);
        var sheets = new HashSet<string>(manifest.SheetNames, StringComparer.OrdinalIgnoreCase);
        var sources = new HashSet<string>(sheets, StringComparer.OrdinalIgnoreCase);
        var names = new HashSet<string>(manifest.Entries.Select(item => item.Name), StringComparer.OrdinalIgnoreCase);
        foreach (var sheet in request.Sheets)
        {
            token.ThrowIfCancellationRequested();
            if (!sheets.Contains(sheet.Name))
                throw Configuration($"模板缺少请求的工作表: {sheet.Name}", provider);
        }
        foreach (var definition in request.PivotTables)
        {
            token.ThrowIfCancellationRequested();
            ValidateSource(definition, sources, provider);
            if (!names.Add(definition.Name) || !sheets.Add(definition.TargetSheet))
                throw Configuration($"新建透视与模板名称或目标冲突: {definition.Name}/{definition.TargetSheet}", provider);
        }
        if (request.PivotTableUpdates.Count > 0 && manifest.RelatedFeatures.Count > 0)
            throw Configuration($"模板包含当前更新路径不支持的关联功能: {string.Join("、", manifest.RelatedFeatures)}。",
                provider);
        var updates = new List<ExcelPivotTemplateUpdate>();
        var updateNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var definition in request.PivotTableUpdates)
        {
            token.ThrowIfCancellationRequested();
            ValidateSource(definition, sources, provider);
            var entry = manifest.Entries.SingleOrDefault(item => string.Equals(item.Name,
                definition.Name, StringComparison.OrdinalIgnoreCase));
            if (entry == null || !string.Equals(entry.SheetName, definition.TargetSheet, StringComparison.OrdinalIgnoreCase)
                || !updateNames.Add(definition.Name))
                throw Configuration($"更新透视的名称或目标工作表不存在或不唯一: {definition.Name}/{definition.TargetSheet}", provider);
            if (entry.HasUnverifiedDependencies || entry.Cache.HasUnverifiedDependencies)
                throw Configuration($"模板透视包含未验证的扩展或依赖，当前不允许更新: {definition.Name}", provider);
            updates.Add(new ExcelPivotTemplateUpdate(entry, definition));
        }
        return new ExcelPivotTemplatePlan(manifest, updates, request.PivotTables);
    }

    /// <summary>
    /// 校验模板来源与格式边界。
    /// </summary>
    /// <param name="definition">透视定义。</param>
    /// <param name="sheets">实际来源工作表。</param>
    /// <param name="provider">提供程序名称。</param>
    private static void ValidateSource(ExcelPivotTableDefinition definition, ISet<string> sheets, string provider)
    {
        if (!sheets.Contains(definition.SourceSheet))
            throw Configuration($"透视来源工作表不存在: {definition.SourceSheet}", provider);
        if (definition.SourceRange.EndRow >= 1048576 || definition.SourceRange.EndColumn >= 16384
            || definition.TargetRow >= 1048576 || definition.TargetColumn >= 16384)
            throw Configuration($"透视来源或目标超出 XLSX 坐标边界: {definition.Name}", provider);
    }

    /// <summary>
    /// 创建模板配置错误。
    /// </summary>
    /// <param name="message">错误说明。</param>
    /// <param name="provider">提供程序名称。</param>
    /// <returns>包含预检阶段信息的异常。</returns>
    private static BingOfficesConfigurationException Configuration(string message, string provider) =>
        new($"{provider}: {message}", stage: BingOfficesStage.Preflight);
}
