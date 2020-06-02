using Bing.Offices.Conversions;
using Bing.Offices.Exports;
using Bing.Offices.Imports;
using Bing.Offices.Npoi.Conversions;
using Bing.Offices.Npoi.Exports;
using Bing.Offices.Npoi.Imports;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;

namespace Bing.Offices.Npoi.Extensions
{
    /// <summary>
    /// 服务扩展
    /// </summary>
    public static partial class Extensions
    {
        /// <summary>
        /// 注册Npoi操作
        /// </summary>
        /// <param name="services">服务集合</param>
        public static void AddNpoi(this IServiceCollection services)
        {
            services.TryAddTransient<IExcelImportService, ExcelImportService>();
            services.TryAddTransient<IExcelExportService, ExcelExportService>();
            services.TryAddTransient<IExcelImportProvider, ExcelImportProvider>();
            services.TryAddTransient<IExcelExportProvider, ExcelExportProvider>();
            services.TryAddTransient<ICellValueConverter, CellValueConverter>();
        }
    }
}
