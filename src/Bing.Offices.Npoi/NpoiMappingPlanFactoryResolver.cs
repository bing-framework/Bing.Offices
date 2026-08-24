using System.Collections.Generic;
using Bing.Offices.Conversions;
using Bing.Offices.Mappings;
using Bing.Offices.Providers;
using Bing.Offices.Validations;

namespace Bing.Offices.Npoi;

/// <summary>
/// 直接构造器的 current-major 兼容解析器；DI 生产组合由 Core 注册入口负责。
/// </summary>
internal static class NpoiMappingPlanFactoryResolver
{
    internal static IExcelMappingPlanFactory CreateDefault(
        IEnumerable<IExcelValueConverter> valueConverters = null,
        IEnumerable<IExcelValidationRule> validationRules = null,
        IEnumerable<INamedExcelValidationRule> namedValidationRules = null) =>
        ExcelMappingPlanFactoryProvider.CreateDefault(valueConverters, validationRules, namedValidationRules);
}
