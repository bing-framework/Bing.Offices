using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Providers;

using System.Collections.Generic;
using Bing.Offices.Conversions;

namespace Bing.Offices.ProviderContract.Tests.Drivers;

/// <summary>
/// 通过公共 API 创建一个 Provider 实例的合同驱动器。
/// </summary>
public interface IProviderContractDriver
{
    /// <summary>
    /// 获取稳定 Provider 名称。
    /// </summary>
    string Name { get; }
    /// <summary>
    /// 获取独立于合同预期的生产能力声明。
    /// </summary>
    IExcelProviderCapabilities DeclaredCapabilities { get; }
    /// <summary>
    /// 创建导出器。
    /// </summary>
    /// <returns>用于合同测试的导出器实例。</returns>
    IExcelExporter CreateExporter();
    /// <summary>
    /// 创建导出器。
    /// </summary>
    /// <param name="valueConverters">注册到导出器的值转换器。</param>
    /// <returns>配置值转换器的导出器实例。</returns>
    IExcelExporter CreateExporter(IEnumerable<IExcelValueConverter> valueConverters);
    /// <summary>
    /// 创建导入器。
    /// </summary>
    /// <returns>用于合同测试的导入器实例。</returns>
    IExcelImporter CreateImporter();
    /// <summary>
    /// 创建导入器。
    /// </summary>
    /// <param name="valueConverters">注册到导入器的值转换器。</param>
    /// <returns>配置值转换器的导入器实例。</returns>
    IExcelImporter CreateImporter(IEnumerable<IExcelValueConverter> valueConverters);
}

/// <summary>
/// 只读 Provider 的合同驱动器。
/// </summary>
/// <remarks>只暴露实际可用的导入服务，不提供导出器。</remarks>
public interface IProviderImportContractDriver
{
    /// <summary>
    /// 获取稳定 Provider 名称。
    /// </summary>
    string Name { get; }
    /// <summary>
    /// 获取独立于合同预期的生产能力声明。
    /// </summary>
    IExcelProviderCapabilities DeclaredCapabilities { get; }
    /// <summary>
    /// 创建导入器。
    /// </summary>
    /// <returns>用于合同测试的导入器实例。</returns>
    IExcelImporter CreateImporter();
    /// <summary>
    /// 创建分批导入器。
    /// </summary>
    /// <returns>用于合同测试的分批导入器实例。</returns>
    IExcelBatchImporter CreateBatchImporter();
}
