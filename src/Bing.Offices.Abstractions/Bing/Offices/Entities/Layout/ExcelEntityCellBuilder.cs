using System.Collections;
using System.ComponentModel;
using System.Linq.Expressions;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

/// <summary>
/// 固定单元格的映射配置构建器。
/// </summary>
/// <typeparam name="TValue">固定单元格绑定的属性类型。</typeparam>
public sealed class ExcelEntityCellBuilder<TValue>
{
    /// <summary>
    /// 获取固定单元格使用的请求级映射配置快照。
    /// </summary>
    internal ExcelMappingConfiguration MappingConfiguration { get; private set; }

    /// <summary>
    /// 获取固定单元格使用的规范化映射文档快照。
    /// </summary>
    internal ExcelMappingDocument MappingDocument { get; private set; }

    /// <summary>
    /// 设置固定单元格使用的请求级映射配置。
    /// </summary>
    /// <param name="configuration">映射配置。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityCellBuilder<TValue> Mapping(ExcelMappingConfiguration configuration)
    {
        MappingConfiguration = configuration == null ? null :
            MappingConfigurationCloner.Clone(configuration, MappingSourceKind.Request);
        return this;
    }

    /// <summary>
    /// 设置固定单元格使用的规范化映射文档。
    /// </summary>
    /// <param name="document">映射文档。</param>
    /// <returns>当前构建器。</returns>
    public ExcelEntityCellBuilder<TValue> Mapping(ExcelMappingDocument document)
    {
        MappingDocument = document == null ? null : MappingDocumentCloner.Clone(document);
        return this;
    }
}
