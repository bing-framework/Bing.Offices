using System;
using System.Collections.Generic;
using System.ComponentModel;
using Bing.Offices.Conversions;
using Bing.Offices.Imports;
using Bing.Offices.Validations;

namespace Bing.Offices.Providers;

/// <summary>
/// Provider 使用的不可变映射计划。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelMappingPlan
{
    /// <summary>获取编译时绑定的业务 Profile 名称。</summary>
    string ProfileName { get; }
    /// <summary>获取编译时绑定的业务模型别名。</summary>
    string ModelAlias { get; }
    /// <summary>
    /// 获取固定列映射。
    /// </summary>
    IReadOnlyList<IExcelMappingColumn> Columns { get; }
    /// <summary>获取已预绑定的动态列映射。</summary>
    IReadOnlyList<IExcelDynamicMappingColumn> DynamicColumns { get; }
    /// <summary>获取 provider-neutral 样式配置。</summary>
    IExcelMappingStyle Style { get; }
    /// <summary>获取 provider-neutral 布局配置。</summary>
    IExcelMappingLayout Layout { get; }
}

/// <summary>Provider 使用的只读动态列映射。</summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelDynamicMappingColumn
{
    string Key { get; }
    string Title { get; }
    IReadOnlyList<string> Aliases { get; }
    string DataTypeName { get; }
    int Order { get; }
    string ConverterName { get; }
    string ValidatorName { get; }
    IReadOnlyList<string> ValidationRuleNames { get; }
    string NumberFormat { get; }
    int? ColumnIndex { get; }
    string PlacementKey { get; }
    ExcelImageMultiplicityPolicy ImageMultiplicity { get; }
    IReadOnlyList<IExcelValueConverter> ValueConverters { get; }
    IReadOnlyList<IExcelValidationBinding> ValidationBindings { get; }
    bool IsUnique { get; }
    bool UniqueIgnoreEmpty { get; }
}

/// <summary>Provider 使用的只读样式配置。</summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelMappingStyle
{
    string HeaderStyleKey { get; }
    string BodyStyleKey { get; }
    string NumberFormat { get; }
}

/// <summary>Provider 使用的只读布局配置。</summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelMappingLayout
{
    int? ColumnIndex { get; }
    string PlacementKey { get; }
}

/// <summary>
/// Provider 使用的只读列映射视图。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelMappingColumn
{
    /// <summary>获取属性名称。</summary>
    string Name { get; }
    /// <summary>获取列标题。</summary>
    string Title { get; }
    /// <summary>获取标题别名。</summary>
    IReadOnlyList<string> Aliases { get; }
    /// <summary>获取格式化字符串。</summary>
    string Formatter { get; }
    /// <summary>获取是否忽略。</summary>
    bool Ignored { get; }
    /// <summary>获取是否为动态列。</summary>
    bool IsDynamicColumn { get; }
    /// <summary>获取导入空白策略。</summary>
    ExcelWhitespacePolicy? ImportWhitespace { get; }
    /// <summary>获取小数精度。</summary>
    byte? DecimalScale { get; }
    /// <summary>获取转换器名称。</summary>
    string ConverterName { get; }
    /// <summary>获取命名校验规则。</summary>
    IReadOnlyList<string> ValidationRuleNames { get; }
    /// <summary>获取显示文本到配置值文本的映射。</summary>
    IReadOnlyDictionary<string, string> ValueMap { get; }
    /// <summary>获取图片多值策略。</summary>
    ExcelImageMultiplicityPolicy ImageMultiplicity { get; }
    /// <summary>获取是否启用唯一性校验。</summary>
    bool IsUnique { get; }
    /// <summary>获取是否忽略空值参与唯一性校验。</summary>
    bool UniqueIgnoreEmpty { get; }
    /// <summary>获取构建阶段绑定的值转换器。</summary>
    IReadOnlyList<IExcelValueConverter> ValueConverters { get; }
    /// <summary>获取构建阶段绑定的校验规则。</summary>
    IReadOnlyList<IExcelValidationBinding> ValidationBindings { get; }
}

/// <summary>
/// Provider-neutral 映射计划工厂。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelMappingPlanFactory
{
    /// <summary>从规范化文档构建指定方向的映射计划。</summary>
    IExcelMappingPlan Create<T>(Configurations.ExcelMappingDocument document,
        Configurations.MappingDirection direction) where T : class, new();

    /// <summary>从规范化文档构建包含 Sheet 视图的 Workbook 计划。</summary>
    IExcelMappingWorkbookPlan CreateWorkbook<T>(Configurations.ExcelMappingDocument document,
        Configurations.MappingDirection direction, IReadOnlyList<string> sheetNames) where T : class, new();
}
