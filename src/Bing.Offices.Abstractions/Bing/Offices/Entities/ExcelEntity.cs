using System.Collections;
using System.ComponentModel;
using System.Linq.Expressions;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

/// <summary>
/// 创建单个聚合对象的 Excel 布局。
/// </summary>
public static class ExcelEntity
{
    /// <summary>
    /// 创建并校验一个不可变实体布局。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="configure">布局配置委托。</param>
    /// <returns>不可变布局。</returns>
    public static ExcelEntityLayout<TEntity> Layout<TEntity>(
        Action<ExcelEntityLayoutBuilder<TEntity>> configure)
        where TEntity : class, new()
    {
        if (configure == null)
            throw new ArgumentNullException(nameof(configure));
        var builder = new ExcelEntityLayoutBuilder<TEntity>();
        configure(builder);
        return builder.Build();
    }

    /// <summary>
    /// 根据实体属性上的固定单元格特性创建布局。
    /// </summary>
    /// <typeparam name="TEntity">聚合对象类型。</typeparam>
    /// <param name="configure">在特性布局基础上继续配置的委托。</param>
    /// <returns>不可变布局。</returns>
    /// <remarks>
    /// 特性只扫描公开实例属性；列表区域、合并区域和映射覆盖仍通过委托配置。
    /// </remarks>
    public static ExcelEntityLayout<TEntity> LayoutFromAttributes<TEntity>(
        Action<ExcelEntityLayoutBuilder<TEntity>> configure = null)
        where TEntity : class, new()
    {
        var builder = new ExcelEntityLayoutBuilder<TEntity>();
        builder.CellsFromAttributes();
        configure?.Invoke(builder);
        return builder.Build();
    }
}
