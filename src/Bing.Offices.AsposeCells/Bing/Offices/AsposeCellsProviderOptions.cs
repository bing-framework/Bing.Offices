using System.Collections.ObjectModel;

namespace Bing.Offices.AsposeCells;

/// <summary>
/// Aspose.Cells Provider 的宿主配置。
/// </summary>
public sealed class AsposeCellsProviderOptions
{
    /// <summary>
    /// 获取或设置许可证文件路径。
    /// </summary>
    /// <remarks>未设置有效许可证路径时，商业能力在预检阶段拒绝执行。</remarks>
    public string LicensePath { get; set; }
    /// <summary>
    /// 获取允许 Aspose.Cells 使用的字体目录集合。
    /// </summary>
    public IList<string> FontDirectories { get; } = new List<string>();
    /// <summary>
    /// 获取字体替代映射。
    /// </summary>
    /// <remarks>键为工作簿字体名，值为替代字体名。</remarks>
    public IDictionary<string, string> FontSubstitutions { get; } =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
    /// <summary>
    /// 获取或设置是否启用严格字体配置检查。
    /// </summary>
    /// <remarks>启用时必须配置至少一个字体目录，并且不生成已配置字体替代的提示。</remarks>
    public bool StrictFonts { get; set; }

    /// <summary>
    /// 复制宿主配置及其字体集合。
    /// </summary>
    /// <returns>拥有独立字体集合的配置副本。</returns>
    internal AsposeCellsProviderOptions Clone() => new AsposeCellsProviderOptions
    {
        LicensePath = LicensePath,
        StrictFonts = StrictFonts
    }.CopyCollectionsFrom(this);

    /// <summary>
    /// 复制指定配置中的字体集合。
    /// </summary>
    /// <param name="source">字体集合的来源配置。</param>
    /// <returns>已复制字体集合的当前配置实例。</returns>
    private AsposeCellsProviderOptions CopyCollectionsFrom(AsposeCellsProviderOptions source)
    {
        foreach (var directory in source.FontDirectories ?? Array.Empty<string>())
            FontDirectories.Add(directory);
        foreach (var pair in source.FontSubstitutions ?? new Dictionary<string, string>())
            FontSubstitutions[pair.Key] = pair.Value;
        return this;
    }
}

/// <summary>
/// 字体配置的只读快照。
/// </summary>
public sealed class AsposeCellsFontConfiguration
{
    /// <summary>
    /// 获取字体目录集合。
    /// </summary>
    public IReadOnlyList<string> Directories { get; }
    /// <summary>
    /// 获取字体替代映射。
    /// </summary>
    public IReadOnlyDictionary<string, string> Substitutions { get; }
    /// <summary>
    /// 获取是否启用严格字体配置检查。
    /// </summary>
    public bool Strict { get; }

    /// <summary>
    /// 初始化一个 <see cref="AsposeCellsFontConfiguration"/> 类型的实例。
    /// </summary>
    /// <param name="options">待创建快照的宿主配置。</param>
    internal AsposeCellsFontConfiguration(AsposeCellsProviderOptions options)
    {
        Directories = new ReadOnlyCollection<string>((options.FontDirectories ?? Array.Empty<string>()).ToList());
        Substitutions = new ReadOnlyDictionary<string, string>(
            new Dictionary<string, string>(options.FontSubstitutions ?? new Dictionary<string, string>(),
                StringComparer.OrdinalIgnoreCase));
        Strict = options.StrictFonts;
    }
}
