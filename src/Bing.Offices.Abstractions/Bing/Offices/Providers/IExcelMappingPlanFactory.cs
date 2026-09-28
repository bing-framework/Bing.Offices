using System.ComponentModel;
using Bing.Offices.Conversions;
using Bing.Offices.Imports;
using Bing.Offices.Validations;

namespace Bing.Offices.Providers;

/// <summary>
/// Provider-neutral 映射计划工厂。
/// </summary>
[EditorBrowsable(EditorBrowsableState.Never)]
public interface IExcelMappingPlanFactory
{
    /// <summary>
    /// 从规范化文档构建指定方向的映射计划。
    /// </summary>
    /// <typeparam name="T">用于编译映射计划的实体类型。</typeparam>
    /// <param name="document">规范化映射文档。</param>
    /// <param name="direction">映射方向。</param>
    /// <returns>编译后的不可变映射计划。</returns>
    IExcelMappingPlan Create<T>(Configurations.ExcelMappingDocument document,
        Configurations.MappingDirection direction) where T : class, new();

    /// <summary>
    /// 从文档和请求级配置构建指定方向的最终映射计划。
    /// </summary>
    /// <typeparam name="T">用于编译映射计划的实体类型。</typeparam>
    /// <param name="document">规范化映射文档。</param>
    /// <param name="requestConfiguration">可选的请求级覆盖配置。</param>
    /// <param name="direction">映射方向。</param>
    /// <returns>合并请求级配置后编译的不可变映射计划。</returns>
    IExcelMappingPlan Create<T>(Configurations.ExcelMappingDocument document,
        Configurations.ExcelMappingConfiguration requestConfiguration,
        Configurations.MappingDirection direction) where T : class, new();

    /// <summary>
    /// 从规范化文档构建包含 Sheet 视图的 Workbook 计划。
    /// </summary>
    /// <typeparam name="T">用于编译映射计划的实体类型。</typeparam>
    /// <param name="document">规范化映射文档。</param>
    /// <param name="direction">映射方向。</param>
    /// <param name="sheetNames">需要生成计划的工作表名称，顺序决定结果顺序。</param>
    /// <returns>按指定工作表顺序组织的 Workbook 映射计划。</returns>
    IExcelMappingWorkbookPlan CreateWorkbook<T>(Configurations.ExcelMappingDocument document,
        Configurations.MappingDirection direction, IReadOnlyList<string> sheetNames) where T : class, new();

    /// <summary>
    /// 从文档和请求级配置构建包含 Sheet 视图的最终 Workbook 计划。
    /// </summary>
    /// <typeparam name="T">用于编译映射计划的实体类型。</typeparam>
    /// <param name="document">规范化映射文档。</param>
    /// <param name="requestConfiguration">可选的请求级覆盖配置。</param>
    /// <param name="direction">映射方向。</param>
    /// <param name="sheetNames">需要生成计划的工作表名称，顺序决定结果顺序。</param>
    /// <returns>按指定工作表顺序组织且应用请求级配置的 Workbook 映射计划。</returns>
    IExcelMappingWorkbookPlan CreateWorkbook<T>(Configurations.ExcelMappingDocument document,
        Configurations.ExcelMappingConfiguration requestConfiguration,
        Configurations.MappingDirection direction, IReadOnlyList<string> sheetNames) where T : class, new();
}
