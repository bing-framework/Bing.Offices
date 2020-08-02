using System;
using System.Linq.Expressions;

namespace Bing.Offices.FluentApi
{
    /// <summary>
    /// Excel配置
    /// </summary>
    public interface IExcelFluent
    {
        ISheetFluent Sheet();
    }

    /// <summary>
    /// Excel配置
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    public interface IExcelFluent<TEntity> : IExcelFluent
    {
        /// <summary>
        /// 属性配置
        /// </summary>
        /// <typeparam name="TProperty">属性类型</typeparam>
        /// <param name="propertyExpression">属性表达式</param>
        IPropertyFluent<TEntity, TProperty> Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression);

        /// <summary>
        /// 属性配置
        /// </summary>
        /// <typeparam name="TProperty">属性类型</typeparam>
        /// <param name="propertyName">属性表达式</param>
        IPropertyFluent<TEntity, TProperty> Property<TProperty>(string propertyName);
    }
}
