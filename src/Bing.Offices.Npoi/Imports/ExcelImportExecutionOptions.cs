using System.Collections.Generic;
using System.Globalization;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Providers;

namespace Bing.Offices.Npoi.Imports;

/// <summary>
/// NPOI 导入器内部执行选项，不作为公开兼容 API 暴露。
/// </summary>
internal sealed class ExcelImportExecutionOptions<T> where T : class, new()
{
    /// <summary>获取或设置工作表在工作簿中的零基索引。</summary>
    internal int SheetIndex { get; set; }
    /// <summary>获取或设置当前导入是否属于多工作表请求。</summary>
    internal bool MultiSheet { get; set; }
    /// <summary>获取或设置表头所在的零基行索引。</summary>
    internal int HeaderRowIndex { get; set; }
    /// <summary>获取或设置数据开始所在的零基行索引。</summary>
    internal int DataRowIndex { get; set; } = 1;
    /// <summary>获取或设置允许读取的最大列数。</summary>
    internal int MaxReadColumns { get; set; } = 100;
    /// <summary>获取或设置允许读取的列索引范围。</summary>
    internal ExcelReadColumnRange ReadColumnRange { get; set; }
    /// <summary>获取或设置表头名称的比较规则。</summary>
    internal ExcelNameComparison HeaderComparison { get; set; } = ExcelNameComparison.OrdinalIgnoreCase;
    /// <summary>获取或设置表头文本的空白处理策略。</summary>
    internal ExcelWhitespacePolicy HeaderWhitespace { get; set; } = ExcelWhitespacePolicy.Trim;
    /// <summary>获取或设置数据单元格文本的空白处理策略。</summary>
    internal ExcelWhitespacePolicy BodyWhitespace { get; set; } = ExcelWhitespacePolicy.Trim;
    /// <summary>获取或设置启用的导入校验范围。</summary>
    internal ExcelImportValidationMode ValidationMode { get; set; } = ExcelImportValidationMode.ConfiguredRules;
    /// <summary>获取或设置遇到不支持的工作簿功能时的处理策略。</summary>
    internal ExcelUnsupportedFeaturePolicy UnsupportedFeaturePolicy { get; set; } = ExcelUnsupportedFeaturePolicy.Fail;
    /// <summary>获取或设置从导入实体取得动态列字典的委托。</summary>
    internal Func<object, object> DynamicTargetGetter { get; set; }
    /// <summary>获取或设置是否要求全部固定映射列均出现在表头中。</summary>
    internal bool RequireExpectedHeaders { get; set; } = true;
    /// <summary>获取或设置请求声明的动态列定义。</summary>
    internal IReadOnlyList<ExcelDynamicColumnDefinition> DynamicColumns { get; set; } =
        System.Array.Empty<ExcelDynamicColumnDefinition>();
    /// <summary>获取或设置是否将未知动态列表头作为结构错误处理。</summary>
    internal bool FailOnUnknownDynamicColumns { get; set; }
    /// <summary>获取或设置单行校验失败后的继续策略。</summary>
    internal ValidateMode ValidateMode { get; set; } = ValidateMode.StopOnFirstFailure;
    /// <summary>获取或设置是否将数据区空行记录为错误。</summary>
    internal bool ReportEmptyRows { get; set; }
    /// <summary>获取或设置遇到首个数据区空行后是否停止读取。</summary>
    internal bool StopAtFirstEmptyRow { get; set; }
    /// <summary>获取或设置请求级映射配置。</summary>
    internal ExcelMappingConfiguration MappingConfiguration { get; set; }
    /// <summary>获取或设置映射计划的来源文档。</summary>
    internal ExcelMappingDocument MappingDocument { get; set; }
    /// <summary>获取或设置预先生成的不可变映射计划。</summary>
    internal IExcelMappingPlan MappingPlan { get; set; }
    /// <summary>获取或设置唯一值跟踪器可保留的最大不同值数量。</summary>
    internal int? MaxTrackedUniqueValues { get; set; }
    /// <summary>获取或设置唯一值比较使用的字符串比较规则。</summary>
    internal StringComparison UniqueComparison { get; set; } = StringComparison.OrdinalIgnoreCase;
    /// <summary>获取或设置文本转换和校验使用的区域性。</summary>
    internal CultureInfo Culture { get; set; } = CultureInfo.InvariantCulture;
    /// <summary>获取或设置当前工作簿是否使用 1904 日期系统。</summary>
    internal bool IsDate1904 { get; set; }
}
