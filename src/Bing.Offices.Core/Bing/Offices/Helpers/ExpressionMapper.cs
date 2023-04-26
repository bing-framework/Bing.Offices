using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using Bing.Collections;
using Bing.Offices.Attributes;
using Bing.Offices.Exceptions;
using Bing.Offices.Metadata.Excels;
using static System.Collections.Hashtable;
using Convert = System.Convert;
using Enum = System.Enum;

namespace Bing.Offices.Helpers;

/// <summary>
/// 表达式树映射
/// </summary>
internal static class ExpressionMapper
{
    /// <summary>
    /// 哈希表
    /// </summary>
    private static readonly Hashtable _table = Synchronized(new Hashtable(1024));

    /// <summary>
    /// 将单元行快速转换为指定类型
    /// </summary>
    /// <typeparam name="T">类型</typeparam>
    /// <param name="row">单元行</param>
    /// <param name="func">转换函数</param>
    public static T FastConvert<T>(IRow row, Func<IList<ICell>, T> func) => func.Invoke(row.Cells);

    /// <summary>
    /// 获取转换函数
    /// </summary>
    /// <typeparam name="T">类型</typeparam>
    /// <param name="key">键</param>
    /// <param name="props">属性信息集合</param>
    public static Func<IList<ICell>, T> GetFunc<T>(string key, IEnumerable<PropertyInfo> props)
    {
        if (_table.ContainsKey(key))
            return (Func<IList<ICell>, T>)_table[key];
        var memberBindingList = new List<MemberBinding>();
        // 获取 FirstOrDefault 方法 TResult FirstOrDefault[TSource,TResult](System.Collections.Generic.IEnumerable`1[TSource], System.Func`2[TSource,bool])
        // 参考地址：https://stackoverflow.com/questions/74310069/makegenericmethod-for-max-is-not-working-in-net-6
        var firstOrDefaultMethod = typeof(Enumerable).GetMethods()
            .Single(m => m.Name == nameof(Enumerable.FirstOrDefault)
                         && m.GetParameters().Length == 2
                         && m.GetParameters()[1].ParameterType == typeof(Func<,>).MakeGenericType(m.GetGenericArguments()[0], typeof(bool)))
            .MakeGenericMethod(new[] { typeof(ICell) });

        var cellsParam = Expression.Parameter(typeof(IList<ICell>), "Cells");
        foreach (var prop in props)
        {
            // lambda表达式：PropertyName = prop.Name
            Expression<Func<ICell, bool>> propertyEqualExpr = c => c.PropertyName == prop.Name;

            // 字典需要额外处理
            if (prop.PropertyType == typeof(IDictionary<string, object>) &&
                prop.GetCustomAttribute<DynamicColumnAttribute>() != null)
            {
                memberBindingList.Add(GetDictionaryBinding(prop, cellsParam));
                continue;
            }

            memberBindingList.Add(GetMapperBinding(prop, firstOrDefaultMethod, cellsParam, propertyEqualExpr));
        }

        var memberInitExpression = Expression.MemberInit(Expression.New(typeof(T)), memberBindingList.ToArray());
        var blockExpr = Expression.Block(memberInitExpression);
        var lambda = Expression.Lambda<Func<IList<ICell>, T>>(blockExpr, new[] { cellsParam });
        var func = lambda.Compile();
        _table[key] = func;
        return func;
    }

    /// <summary>
    /// 获取字典绑定表达式
    /// </summary>
    /// <param name="prop">属性信息</param>
    /// <param name="cellsParam">单元格列表参数</param>
    private static MemberBinding GetDictionaryBinding(PropertyInfo prop, ParameterExpression cellsParam)
    {
        // 调用ConvertToDictionary方法
        var convertToDictionaryMethod =
            typeof(ExpressionMapper).GetMethods().First(m => m.Name == nameof(ConvertToDictionary));
        var convertToDictionaryExpr = Expression.Call(convertToDictionaryMethod, cellsParam);
        return Expression.Bind(prop, convertToDictionaryExpr);
    }

    /// <summary>
    /// 获取映射绑定表达式
    /// </summary>
    /// <param name="prop">属性信息</param>
    /// <param name="firstOrDefaultMethod">FirstOrDefault方法</param>
    /// <param name="cellsParam">单元格列表参数</param>
    /// <param name="propertyEqualExpr">属性相等表达式</param>
    private static MemberBinding GetMapperBinding(PropertyInfo prop, MethodInfo firstOrDefaultMethod, ParameterExpression cellsParam, Expression<Func<ICell, bool>> propertyEqualExpr)
    {
        // 调用ChangeType方法
        var changeTypeMethod = typeof(ExpressionMapper).GetMethods()
            .First(m => m.Name == nameof(ChangeType));
        // FirstOrDefault方法
        var firstOrDefaultMethodExpr = Expression.Call(firstOrDefaultMethod, cellsParam, propertyEqualExpr);
        // 获取单元格值表达，Cells.SingleOrDefault(c => c.PropertyName == prop.Name).Value
        var cellValueExpr =
            Expression.Condition(Expression.Equal(firstOrDefaultMethodExpr, Expression.Constant(null)),
                Expression.Constant(null),
                Expression.Property(firstOrDefaultMethodExpr, typeof(ICell), "Value"));
        // 当前属性类型常量
        var propTypeConst = Expression.Constant(prop.PropertyType);
        // 变更类型
        var changeTypeExpr =
            Expression.Call(changeTypeMethod, cellValueExpr, propTypeConst, firstOrDefaultMethodExpr);
        Expression expr = Expression.Convert(changeTypeExpr, prop.PropertyType);
        return Expression.Bind(prop, expr);
    }

    /// <summary>
    /// 变更类型
    /// </summary>
    /// <param name="value">值</param>
    /// <param name="type">类型</param>
    /// <param name="cell">单元格</param>
    public static object ChangeType(object value, Type type, ICell cell)
    {
        try
        {
            if (value == null && type.IsGenericType)
                return Activator.CreateInstance(type);
            if (value == null || string.IsNullOrWhiteSpace(value.ToString()))
                return null;
            if (type == value.GetType())
                return value;
            if (type.IsEnum)
            {
                if (value is string)
                    return Enum.Parse(type, value as string);
                else
                    return Enum.ToObject(type, value);
            }
            if (!type.IsInterface && type.IsGenericType)
            {
                Type innerType = type.GetGenericArguments()[0];
                object innerValue = ChangeType(value, innerType, cell);
                return Activator.CreateInstance(type, new object[] { innerValue });
            }
            if (value is string && type == typeof(Guid))
                return new Guid(value as string);
            if (value is string && type == typeof(Version))
                return new Version(value as string);
            if (!(value is IConvertible))
                return value;
            return Convert.ChangeType(value, type);
        }
        catch (Exception ex)
        {
            throw new OfficeDataConvertException($"值转换失败。输入值为: {value}, 目标类型为: {type.FullName}", ex)
            {
                PrimitiveType = value.GetType(),
                TargetType = type,
                Value = value,
                RowIndex = cell.RowIndex + 1,
                ColumnIndex = cell.ColumnIndex + 1,
                Name = cell.Name,
            };
        }
    }

    /// <summary>
    /// 转换为字典
    /// </summary>
    /// <param name="cells">单元格集合</param>
    public static IDictionary<string, object> ConvertToDictionary(IEnumerable<ICell> cells)
    {
        var dynamicCells = cells.Where(x => x.IsDynamic);
        var dictionary = new Dictionary<string, object>();
        dynamicCells.ForEach(cell => { dictionary[cell.Name] = cell.Value; });
        return dictionary;
    }
}
