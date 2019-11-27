using Bing.Offices.Excel.Configurations;
using Bing.Offices.Excel.Mappings;
using Bing.Offices.Mappings.Tests.Models;

namespace Bing.Offices.Mappings.Tests.Mappings.Exports
{
    /// <summary>
    /// 用户导出映射
    /// </summary>
    public class UserExportMap : ExcelExportMapBase<User>
    {
        /// <summary>
        /// 映射导出配置
        /// </summary>
        /// <param name="configuration">配置</param>
        protected override void MapExport(IExcelConfiguration<User> configuration)
        {
            configuration.HasTitle("测试标题");
            configuration.Property(x => x.Id).HasColumnTitle("用户标识");
            configuration.Property(x => x.Name).HasColumnTitle("姓名");
            configuration.Property(x => x.Age).HasColumnTitle("年龄");
        }
    }
}
