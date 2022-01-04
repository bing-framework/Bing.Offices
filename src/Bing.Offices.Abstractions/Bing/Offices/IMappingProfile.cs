using Bing.Offices.Configurations;

namespace Bing.Offices
{
    /// <summary>
    /// 映射配置文件
    /// </summary>
    public interface IMappingProfile
    {
    }

    /// <summary>
    /// 映射配置文件
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    public interface IMappingProfile<T> : IMappingProfile
    {
        /// <summary>
        /// 配置
        /// </summary>
        /// <param name="configuration">Excel 配置</param>
        void Configure(IExcelConfiguration<T> configuration);
    }
}
