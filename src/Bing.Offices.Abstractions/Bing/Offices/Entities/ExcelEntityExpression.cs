using System.Collections;
using System.ComponentModel;
using System.Linq.Expressions;
using Bing.Offices.Configurations;
using Bing.Offices.Exports;
using Bing.Offices.Imports;

namespace Bing.Offices.Entities;

/// <summary>
/// 负责解析实体属性表达式并创建已编译访问委托。
/// </summary>
internal static class ExcelEntityExpression
{
    /// <summary>
    /// 从直接属性访问表达式中取得属性元数据。
    /// </summary>
    /// <param name="expression">待解析的表达式主体。</param>
    /// <returns>表达式直接访问的属性元数据。</returns>
    internal static System.Reflection.PropertyInfo GetProperty(Expression expression)
    {
        if (expression is UnaryExpression unary && unary.NodeType == ExpressionType.Convert)
            expression = unary.Operand;
        if (!(expression is MemberExpression member) || !(member.Member is System.Reflection.PropertyInfo property)
            || member.Expression == null || member.Expression.NodeType != ExpressionType.Parameter)
            throw new ArgumentException("实体布局属性必须是直接属性表达式。", nameof(expression));
        return property;
    }

    /// <summary>
    /// 为可写属性创建接受对象参数的写入委托。
    /// </summary>
    /// <typeparam name="TEntity">声明目标属性的实体类型。</typeparam>
    /// <param name="property">待写入的属性元数据。</param>
    /// <returns>已编译的写入委托；属性不可写时返回 <see langword="null" />。</returns>
    internal static Action<object, object> CreateObjectSetter<TEntity>(System.Reflection.PropertyInfo property)
    {
        if (!property.CanWrite)
            return null;
        var target = Expression.Parameter(typeof(object), "target");
        var value = Expression.Parameter(typeof(object), "value");
        var assignment = Expression.Assign(Expression.Property(Expression.Convert(target, typeof(TEntity)), property),
            Expression.Convert(value, property.PropertyType));
        return Expression.Lambda<Action<object, object>>(assignment, target, value).Compile();
    }

    /// <summary>
    /// 为指定值类型的可写属性创建对象写入委托。
    /// </summary>
    /// <typeparam name="TEntity">声明目标属性的实体类型。</typeparam>
    /// <typeparam name="TValue">属性值类型。</typeparam>
    /// <param name="property">待写入的属性元数据。</param>
    /// <returns>已编译的写入委托；属性不可写时返回 <see langword="null" />。</returns>
    internal static Action<object, object> CreateSetter<TEntity, TValue>(System.Reflection.PropertyInfo property)
        => CreateObjectSetter<TEntity>(property);
}
