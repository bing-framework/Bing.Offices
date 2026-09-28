namespace Bing.Offices.Imports;

/// <summary>
/// 导入规则来源组合策略。
/// </summary>
public enum ExcelImportValidationMode
{
    /// <summary>
    /// 禁用配置和工作簿原生规则。
    /// </summary>
    Disabled,
    /// <summary>
    /// 只执行配置和属性规则。
    /// </summary>
    ConfiguredRules,
    /// <summary>
    /// 只执行工作簿原生规则。
    /// </summary>
    WorkbookRules,
    /// <summary>
    /// 同时执行配置和工作簿规则。
    /// </summary>
    ConfiguredAndWorkbook
}
