using System.Collections.Generic;
using System.Globalization;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Npoi.Imports;

/// <summary>
/// NPOI 导入器内部执行选项，不作为公开兼容 API 暴露。
/// </summary>
internal sealed class ExcelImportExecutionOptions<T> where T : class, new()
{
    internal int SheetIndex { get; set; }
    internal bool MultiSheet { get; set; }
    internal int HeaderRowIndex { get; set; }
    internal int DataRowIndex { get; set; } = 1;
    internal int MaxColumnLength { get; set; } = 100;
    internal ExcelReadColumnRange ReadColumnRange { get; set; }
    internal ExcelNameComparison HeaderComparison { get; set; } = ExcelNameComparison.OrdinalIgnoreCase;
    internal ExcelWhitespacePolicy HeaderWhitespace { get; set; } = ExcelWhitespacePolicy.Trim;
    internal ExcelWhitespacePolicy BodyWhitespace { get; set; } = ExcelWhitespacePolicy.Trim;
    internal ExcelImportValidationMode ValidationMode { get; set; } = ExcelImportValidationMode.ConfiguredRules;
    internal ExcelUnsupportedFeaturePolicy UnsupportedFeaturePolicy { get; set; } = ExcelUnsupportedFeaturePolicy.Fail;
    internal Func<object, object> DynamicTargetGetter { get; set; }
    internal bool HeaderMatch { get; set; } = true;
    internal IReadOnlyList<ExcelDynamicColumnDefinition> DynamicColumns { get; set; } =
        System.Array.Empty<ExcelDynamicColumnDefinition>();
    internal bool FailOnUnknownDynamicColumns { get; set; }
    internal ValidateMode ValidateMode { get; set; } = ValidateMode.StopOnFirstFailure;
    internal bool EnabledEmptyLine { get; set; }
    internal bool IgnoreEmptyLineAfterData { get; set; }
    internal ExcelMappingConfiguration MappingConfiguration { get; set; }
    internal ExcelMappingProfile<T> MappingProfile { get; set; }
    internal CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;
}
