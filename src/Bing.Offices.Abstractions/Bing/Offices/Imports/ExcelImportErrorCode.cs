namespace Bing.Offices.Imports;

/// <summary>
/// Excel 导入错误代码。
/// </summary>
public enum ExcelImportErrorCode
{
    /// <summary>
    /// 输入数据格式无效。
    /// </summary>
    InvalidInput,

    /// <summary>
    /// 表头无效。
    /// </summary>
    InvalidHeader,

    /// <summary>
    /// 单元格值无法转换。
    /// </summary>
    ValueConversion,

    /// <summary>
    /// 单元格数据未通过业务校验。
    /// </summary>
    Validation,

    /// <summary>
    /// 父子导航关系绑定失败。
    /// </summary>
    Relationship,

    /// <summary>
    /// Workbook 原生 Data Validation 校验失败或规则不受支持。
    /// </summary>
    WorkbookValidation,

    /// <summary>
    /// 输入工作簿超过资源限制。
    /// </summary>
    ResourceLimit,
}
