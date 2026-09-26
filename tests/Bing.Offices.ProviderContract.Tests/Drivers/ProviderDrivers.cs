using System.Collections.Generic;
using Bing.Offices.ClosedXml.Exports;
using Bing.Offices.ClosedXml.Imports;
using Bing.Offices.ExcelDataReader;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.ProviderContract.Tests.Drivers;

/// <summary>
/// 所有 Provider 合同驱动器的单一枚举入口。
/// </summary>
public static class ProviderDrivers
{
    /// <summary>
    /// 获取 NPOI、MiniExcel、ClosedXML 三个驱动器。
    /// </summary>
    public static IReadOnlyList<IProviderContractDriver> All { get; } = new IProviderContractDriver[]
    {
        new ProviderContractDriver("NPOI", () => new NpoiExcelExporter(), () => new NpoiExcelImporter(),
            converters => new NpoiExcelExporter(converters),
            converters => new NpoiExcelImporter(valueConverters: converters)),
        new ProviderContractDriver("MiniExcel", () => new MiniExcelExcelExporter(), () => new MiniExcelExcelImporter(),
            converters => new MiniExcelExcelExporter(converters),
            converters => new MiniExcelExcelImporter(valueConverters: converters)),
        new ProviderContractDriver("ClosedXML", () => new ClosedXmlExcelExporter(), () => new ClosedXmlExcelImporter(),
            converters => new ClosedXmlExcelExporter(converters),
            converters => new ClosedXmlExcelImporter(valueConverters: converters))
    };

    /// <summary>
    /// 获取只读 Provider 合同驱动器集合。
    /// </summary>
    public static IReadOnlyList<IProviderImportContractDriver> ImportOnly { get; } =
        new IProviderImportContractDriver[]
        {
            new ProviderImportContractDriver("ExcelDataReader",
                () => new ExcelDataReaderExcelImporter(),
                () => new ExcelDataReaderExcelImporter())
        };

    /// <summary>
    /// 按名称获取驱动器。
    /// </summary>
    /// <param name="name">区分大小写的 Provider 名称。</param>
    /// <returns>名称匹配的合同驱动器。</returns>
    public static IProviderContractDriver Get(string name)
    {
        foreach (var driver in All)
            if (string.Equals(driver.Name, name, System.StringComparison.Ordinal))
                return driver;
        throw new System.ArgumentException($"Unknown provider contract driver: {name}", nameof(name));
    }

    /// <summary>
    /// 按名称获取只读 Provider 驱动器。
    /// </summary>
    /// <param name="name">区分大小写的只读 Provider 名称。</param>
    /// <returns>名称匹配的只读合同驱动器。</returns>
    public static IProviderImportContractDriver GetImportOnly(string name)
    {
        foreach (var driver in ImportOnly)
            if (string.Equals(driver.Name, name, System.StringComparison.Ordinal))
                return driver;
        throw new System.ArgumentException($"Unknown import-only provider contract driver: {name}", nameof(name));
    }
}
