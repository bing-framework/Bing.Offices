#if BOM_MIGRATION
namespace Bing.Offices.Exports
{
    /// <summary>
    /// 用于验证打包分析器的旧版导出服务替身。
    /// </summary>
    public interface IExcelExportService { }

    /// <summary>
    /// 用于验证打包分析器的旧版选项替身。
    /// </summary>
    /// <typeparam name="T">数据项类型。</typeparam>
    public sealed class ExportOptions<T> { }
}

namespace Bing.Offices.Attributes
{
    /// <summary>
    /// 用于验证打包分析器的旧版动态列标记替身。
    /// </summary>
    public sealed class HasDynamicColumnAttribute : System.Attribute { }
}

namespace Bing.Offices.Analyzers.Consumer
{
    /// <summary>
    /// 验证旧 API 引用产生迁移提示。
    /// </summary>
    [Bing.Offices.Attributes.HasDynamicColumn]
    public sealed class LegacyOrder
    {
        /// <summary>
        /// 获取或设置旧版导出服务。
        /// </summary>
        public Bing.Offices.Exports.IExcelExportService Service { get; set; }

        /// <summary>
        /// 获取或设置旧版导出选项。
        /// </summary>
        public Bing.Offices.Exports.ExportOptions<LegacyOrder> Options { get; set; }
    }
}
#endif
