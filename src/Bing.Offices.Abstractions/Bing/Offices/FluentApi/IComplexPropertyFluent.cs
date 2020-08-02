using System;
using System.Linq.Expressions;

namespace Bing.Offices.FluentApi
{
    /// <summary>
    /// 复杂属性配置
    /// </summary>
    public interface IComplexPropertyFluent
    {
    }

    /// <summary>
    /// 复杂属性配置
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    public interface IComplexPropertyFluent<TEntity> : IComplexPropertyFluent
    {
        IPropertyFluent<TEntity, TProperty>
            Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression);
    }
}
