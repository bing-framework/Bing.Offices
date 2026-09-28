using System.Collections.ObjectModel;

namespace Bing.Offices.AsposeCells;

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
