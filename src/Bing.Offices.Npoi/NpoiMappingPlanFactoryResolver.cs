using System.Collections.Generic;
using Bing.Offices.Conversions;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Validations;

namespace Bing.Offices.Npoi;

/// <summary>
/// 为直接构造器路径创建默认映射计划工厂；生产环境的依赖组合由 Core 注册入口负责。
/// </summary>
internal static class NpoiMappingPlanFactoryResolver
{
    /// <summary>创建直接构造器路径使用的默认映射计划工厂。</summary>
    /// <param name="valueConverters">可选值转换器集合。</param>
    /// <param name="validationRules">可选特性校验规则集合。</param>
    /// <param name="namedValidationRules">可选命名校验规则集合。</param>
    /// <returns>已配置的映射计划工厂。</returns>
    internal static IExcelMappingPlanFactory CreateDefault(
        IEnumerable<IExcelValueConverter> valueConverters = null,
        IEnumerable<IExcelValidationRule> validationRules = null,
        IEnumerable<INamedExcelValidationRule> namedValidationRules = null) =>
        ExcelMappingPlanFactoryProvider.CreateDefault(valueConverters: valueConverters,
            validationRules: validationRules, namedValidationRules: namedValidationRules);
}
