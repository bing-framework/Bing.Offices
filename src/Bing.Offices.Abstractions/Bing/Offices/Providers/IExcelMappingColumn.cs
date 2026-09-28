using System.ComponentModel;
using Bing.Offices.Conversions;
using Bing.Offices.Imports;
using Bing.Offices.Validations;

namespace Bing.Offices.Providers;

/// <summary>
/// Provider 使用的只读列映射视图。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelMappingColumn
{
    /// <summary>
    /// 获取用于实体属性绑定和错误定位的属性名称。
    /// </summary>
    string Name { get; }
    /// <summary>
    /// 获取列标题。
    /// </summary>
    string Title { get; }
    /// <summary>
    /// 获取标题别名。
    /// </summary>
    IReadOnlyList<string> Aliases { get; }
    /// <summary>
    /// 获取格式化字符串。
    /// </summary>
    string Formatter { get; }
    /// <summary>
    /// 获取是否忽略。
    /// </summary>
    bool Ignored { get; }
    /// <summary>
    /// 获取是否为动态列。
    /// </summary>
    bool IsDynamicColumn { get; }
    /// <summary>
    /// 获取导入空白策略。
    /// </summary>
    ExcelWhitespacePolicy? ImportWhitespace { get; }
    /// <summary>
    /// 获取小数精度。
    /// </summary>
    byte? DecimalScale { get; }
    /// <summary>
    /// 获取转换器名称。
    /// </summary>
    string ConverterName { get; }
    /// <summary>
    /// 获取命名校验规则。
    /// </summary>
    IReadOnlyList<string> ValidationRuleNames { get; }
    /// <summary>
    /// 获取显示文本到配置值文本的映射。
    /// </summary>
    IReadOnlyDictionary<string, string> ValueMap { get; }
    /// <summary>
    /// 获取图片多值策略。
    /// </summary>
    ExcelImageMultiplicityPolicy ImageMultiplicity { get; }
    /// <summary>
    /// 获取是否启用唯一性校验。
    /// </summary>
    bool IsUnique { get; }
    /// <summary>
    /// 获取是否忽略空值参与唯一性校验。
    /// </summary>
    bool UniqueIgnoreEmpty { get; }
    /// <summary>
    /// 获取构建阶段绑定的值转换器。
    /// </summary>
    IReadOnlyList<IExcelValueConverter> ValueConverters { get; }
    /// <summary>
    /// 获取构建阶段绑定的校验规则。
    /// </summary>
    IReadOnlyList<IExcelValidationBinding> ValidationBindings { get; }
}
