namespace Bing.Offices.Csv;

/// <summary>
/// CSV 字段的公式注入处理策略。
/// </summary>
public enum CsvFormulaInjectionPolicy
{
    /// <summary>
    /// 在潜在公式字段前添加单引号，使表格软件将其作为文本处理。
    /// </summary>
    Escape = 0,

    /// <summary>
    /// 保持字段原始内容。
    /// </summary>
    None = 1
}
