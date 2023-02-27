using Bing.Offices.Decorators;

namespace Bing.Offices.Exports;

/// <summary>
/// Excel导出提供程序
/// </summary>
public interface IExcelExportProvider
{
    /// <summary>
    /// 导出
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    /// <param name="options">导出选项配置</param>
    byte[] Export<T>(IExportOptions<T> options) where T : class, new();

    /// <summary>
    /// 处理表头
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    /// <param name="workbookBytes">工作簿流</param>
    /// <param name="options">导出选项配置</param>
    /// <param name="context">装饰器上下文</param>
    byte[] HandleHeader<T>(byte[] workbookBytes, IExportOptions<T> options, IDecoratorContext context)
        where T : class, new();

    /// <summary>
    /// 合并单元格
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    /// <param name="workbookBytes">工作簿流</param>
    /// <param name="options">导出选项配置</param>
    /// <param name="context">装饰器上下文</param>
    byte[] MergeCells<T>(byte[] workbookBytes, IExportOptions<T> options, IDecoratorContext context)
        where T : class, new();

    /// <summary>
    /// 自动换行
    /// </summary>
    /// <typeparam name="T">实体类型</typeparam>
    /// <param name="workbookBytes">工作簿流</param>
    /// <param name="options">导出选项配置</param>
    /// <param name="context">装饰器上下文</param>
    byte[] WarpText<T>(byte[] workbookBytes, IExportOptions<T> options, IDecoratorContext context)
        where T : class, new();
}