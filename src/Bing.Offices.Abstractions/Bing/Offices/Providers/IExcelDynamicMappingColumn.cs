using System.ComponentModel;
using Bing.Offices.Conversions;
using Bing.Offices.Imports;
using Bing.Offices.Validations;

namespace Bing.Offices.Providers;

/// <summary>
/// Provider 使用的只读动态列映射。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelDynamicMappingColumn
{
    /// <summary>
    /// 获取动态列的稳定标识键。
    /// </summary>
    string Key { get; }
    /// <summary>
    /// 获取导出表头或导入匹配使用的显示标题。
    /// </summary>
    string Title { get; }
    /// <summary>
    /// 获取可用于导入表头匹配的历史标题别名。
    /// </summary>
    IReadOnlyList<string> Aliases { get; }
    /// <summary>
    /// 获取动态列值的逻辑数据类型名称。
    /// </summary>
    string DataTypeName { get; }
    /// <summary>
    /// 获取动态列在未指定位置时采用的排序序号。
    /// </summary>
    int Order { get; }
    /// <summary>
    /// 获取显式指定的值转换器名称。
    /// </summary>
    string ConverterName { get; }
    /// <summary>
    /// 获取兼容的单个命名校验规则名称。
    /// </summary>
    string ValidatorName { get; }
    /// <summary>
    /// 获取应用于动态列的命名校验规则名称集合。
    /// </summary>
    IReadOnlyList<string> ValidationRuleNames { get; }
    /// <summary>
    /// 获取导出值使用的数字格式。
    /// </summary>
    string NumberFormat { get; }
    /// <summary>
    /// 获取动态列的零基物理列索引。
    /// </summary>
    int? ColumnIndex { get; }
    /// <summary>
    /// 获取相对其他动态列定位的键。
    /// </summary>
    string PlacementKey { get; }
    /// <summary>
    /// 获取单元格包含多个图片时采用的处理策略。
    /// </summary>
    ExcelImageMultiplicityPolicy ImageMultiplicity { get; }
    /// <summary>
    /// 获取构建阶段已绑定的值转换器。
    /// </summary>
    IReadOnlyList<IExcelValueConverter> ValueConverters { get; }
    /// <summary>
    /// 获取构建阶段已绑定的校验规则。
    /// </summary>
    IReadOnlyList<IExcelValidationBinding> ValidationBindings { get; }
    /// <summary>
    /// 获取是否对该动态列启用唯一性校验。
    /// </summary>
    bool IsUnique { get; }
    /// <summary>
    /// 获取唯一性校验是否忽略空值。
    /// </summary>
    bool UniqueIgnoreEmpty { get; }
}
