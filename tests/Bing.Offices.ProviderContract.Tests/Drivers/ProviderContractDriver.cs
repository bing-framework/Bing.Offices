using System;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Providers;

using System.Collections.Generic;
using Bing.Offices.Conversions;

namespace Bing.Offices.ProviderContract.Tests.Drivers;

/// <summary>
/// 封装 Provider 构造函数和生产能力声明的测试驱动器。
/// </summary>
public sealed class ProviderContractDriver : IProviderContractDriver
{
    /// <summary>
    /// 创建默认导出器的工厂。
    /// </summary>
    private readonly Func<IExcelExporter> _exporterFactory;
    /// <summary>
    /// 创建默认导入器的工厂。
    /// </summary>
    private readonly Func<IExcelImporter> _importerFactory;
    /// <summary>
    /// 创建带指定值转换器的导出器的工厂。
    /// </summary>
    private readonly Func<IEnumerable<IExcelValueConverter>, IExcelExporter> _configuredExporterFactory;
    /// <summary>
    /// 创建带指定值转换器的导入器的工厂。
    /// </summary>
    private readonly Func<IEnumerable<IExcelValueConverter>, IExcelImporter> _configuredImporterFactory;

    /// <summary>
    /// 初始化一个 <see cref="ProviderContractDriver"/> 类型的实例。
    /// </summary>
    /// <param name="name">Provider 稳定名称。</param>
    /// <param name="exporterFactory">默认导出器工厂。</param>
    /// <param name="importerFactory">默认导入器工厂。</param>
    /// <param name="configuredExporterFactory">带值转换器的导出器工厂；为空时使用默认工厂。</param>
    /// <param name="configuredImporterFactory">带值转换器的导入器工厂；为空时使用默认工厂。</param>
    public ProviderContractDriver(string name, Func<IExcelExporter> exporterFactory,
        Func<IExcelImporter> importerFactory,
        Func<IEnumerable<IExcelValueConverter>, IExcelExporter> configuredExporterFactory = null,
        Func<IEnumerable<IExcelValueConverter>, IExcelImporter> configuredImporterFactory = null)
    {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        _exporterFactory = exporterFactory ?? throw new ArgumentNullException(nameof(exporterFactory));
        _importerFactory = importerFactory ?? throw new ArgumentNullException(nameof(importerFactory));
        _configuredExporterFactory = configuredExporterFactory ?? (_ => _exporterFactory());
        _configuredImporterFactory = configuredImporterFactory ?? (_ => _importerFactory());
        var exporter = _exporterFactory();
        var importer = _importerFactory();
        if (exporter is not IExcelProviderCapabilities exporterCapabilities)
            throw new InvalidOperationException($"{name} exporter does not expose provider capabilities.");
        if (importer is not IExcelProviderCapabilities importerCapabilities)
            throw new InvalidOperationException($"{name} importer does not expose provider capabilities.");
        if (!string.Equals(exporterCapabilities.ProviderName, importerCapabilities.ProviderName,
                StringComparison.Ordinal))
            throw new InvalidOperationException($"{name} exporter/importer capability names differ.");
        DeclaredCapabilities = exporterCapabilities;
    }

    /// <inheritdoc />
    public string Name { get; }
    /// <inheritdoc />
    public IExcelProviderCapabilities DeclaredCapabilities { get; }
    /// <inheritdoc />
    public IExcelExporter CreateExporter() => _exporterFactory();
    /// <inheritdoc />
    public IExcelExporter CreateExporter(IEnumerable<IExcelValueConverter> valueConverters) =>
        _configuredExporterFactory(valueConverters ?? Array.Empty<IExcelValueConverter>());
    /// <inheritdoc />
    public IExcelImporter CreateImporter() => _importerFactory();
    /// <inheritdoc />
    public IExcelImporter CreateImporter(IEnumerable<IExcelValueConverter> valueConverters) =>
        _configuredImporterFactory(valueConverters ?? Array.Empty<IExcelValueConverter>());
}

/// <summary>
/// 封装只读 Provider 的导入和分批导入服务。
/// </summary>
public sealed class ProviderImportContractDriver : IProviderImportContractDriver
{
    /// <summary>
    /// 创建只读 Provider 导入器的工厂。
    /// </summary>
    private readonly Func<IExcelImporter> _importerFactory;
    /// <summary>
    /// 创建只读 Provider 分批导入器的工厂。
    /// </summary>
    private readonly Func<IExcelBatchImporter> _batchImporterFactory;

    /// <summary>
    /// 初始化一个 <see cref="ProviderImportContractDriver"/> 类型的实例。
    /// </summary>
    /// <param name="name">Provider 稳定名称。</param>
    /// <param name="importerFactory">工作簿导入器工厂。</param>
    /// <param name="batchImporterFactory">分批导入器工厂。</param>
    public ProviderImportContractDriver(string name, Func<IExcelImporter> importerFactory,
        Func<IExcelBatchImporter> batchImporterFactory)
    {
        Name = name ?? throw new ArgumentNullException(nameof(name));
        _importerFactory = importerFactory ?? throw new ArgumentNullException(nameof(importerFactory));
        _batchImporterFactory = batchImporterFactory ?? throw new ArgumentNullException(nameof(batchImporterFactory));
        var importer = _importerFactory();
        if (importer is not IExcelProviderCapabilities capabilities)
            throw new InvalidOperationException($"{name} importer does not expose provider capabilities.");
        DeclaredCapabilities = capabilities;
    }

    /// <inheritdoc />
    public string Name { get; }
    /// <inheritdoc />
    public IExcelProviderCapabilities DeclaredCapabilities { get; }
    /// <inheritdoc />
    public IExcelImporter CreateImporter() => _importerFactory();
    /// <inheritdoc />
    public IExcelBatchImporter CreateBatchImporter() => _batchImporterFactory();
}
