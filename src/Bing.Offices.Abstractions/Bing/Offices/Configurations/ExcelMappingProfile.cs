namespace Bing.Offices.Configurations;

/// <summary>
/// 实体类型的不可变 Fluent 映射配置文件。
/// </summary>
/// <typeparam name="T">实体类型。</typeparam>
[Obsolete("请改用 ExcelMappingProfile<TImport, TExport>。")]
public sealed class ExcelMappingProfile<T> where T : class, new()
{
    private readonly ExcelMappingConfiguration _configuration;

    /// <summary>
    /// 初始化一个<see cref="ExcelMappingProfile{T}"/>类型的实例。
    /// </summary>
    /// <param name="configuration">已构建的映射配置。</param>
    public ExcelMappingProfile(ExcelMappingConfiguration configuration)
    {
        _configuration = Clone(configuration ?? throw new ArgumentNullException(nameof(configuration)));
    }

    /// <summary>
    /// 使用同模型方向 Profile 构建兼容 Profile。
    /// </summary>
    /// <param name="profile">方向 Profile。</param>
    public ExcelMappingProfile(IMappingProfile<T> profile)
        : this(new ExcelMappingProfile<T, T>(profile).ImportConfiguration)
    {
    }

    /// <summary>
    /// 使用同模型方向 Fluent 设置构建兼容 Profile。
    /// </summary>
    /// <param name="configure">方向配置委托。</param>
    public ExcelMappingProfile(Action<FluentSetting<T, T>> configure)
        : this(new ExcelMappingProfile<T, T>(configure).ImportConfiguration)
    {
    }

    /// <summary>
    /// 获取当前 Profile 的独立配置快照。
    /// </summary>
    public ExcelMappingConfiguration Configuration => Clone(_configuration);

    private static ExcelMappingConfiguration Clone(ExcelMappingConfiguration configuration) => new()
    {
        Columns = (configuration.Columns ?? new List<ExcelColumnConfiguration>()).Select(column => new ExcelColumnConfiguration
        {
            PropertyName = column.PropertyName,
            Title = column.Title,
            Aliases = (column.Aliases ?? new List<string>()).ToList(),
            ColumnIndex = column.ColumnIndex,
            Ignored = column.Ignored,
            Formatter = column.Formatter,
            DecimalScale = column.DecimalScale,
            ConverterName = column.ConverterName,
            ImportWhitespace = column.ImportWhitespace,
            ValidationRuleNames = (column.ValidationRuleNames ?? new List<string>()).ToList(),
            ValidationRuleNamesToRemove = (column.ValidationRuleNamesToRemove ?? new List<string>()).ToList(),
            ClearValidationRules = column.ClearValidationRules,
            ValidationRuleMergeMode = column.ValidationRuleMergeMode,
            ValueMappings = (column.ValueMappings ?? new List<ExcelValueMappingConfiguration>()).Select(mapping =>
                new ExcelValueMappingConfiguration { Text = mapping.Text, Value = mapping.Value }).ToList(),
            ValueMappingMergeMode = column.ValueMappingMergeMode,
            ImageMultiplicity = column.ImageMultiplicity
        }).ToList()
    };
}
