using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using Bing.Offices.Abstractions.Metadata.Excels;
using Convert = System.Convert;
using Enum = System.Enum;

namespace Bing.Offices.Helpers
{
    /// <summary>
    /// 表达式树映射
    /// </summary>
    internal static class ExpressionMapper
    {
        /// <summary>
        /// 哈希表
        /// </summary>
        private static Hashtable Table = Hashtable.Synchronized(new Hashtable(1024));

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
            if (Table.ContainsKey(key))
                return (Func<IList<ICell>, T>) Table[key];
            var memberBindingList = new List<MemberBinding>();
            // 获取FirstOrDefault方法
            var firstOrDefaultMethod = typeof(Enumerable).GetMethods()
                .Single(m => m.Name == "FirstOrDefault" && m.GetParameters().Length == 2)
                .MakeGenericMethod(new[] {typeof(ICell)});
            var cellsParam = Expression.Parameter(typeof(IList<ICell>), "Cells");
            foreach (var prop in props)
            {
                // lambda表达式：PropertyName = prop.Name
                Expression<Func<ICell, bool>> propertyEqualExpr = c => c.PropertyName == prop.Name;
                // 调用ChangeType方法
                var changeTypeMethod = typeof(ExpressionMapper).GetMethods()
                    .First(m => m.Name == nameof(ChangeType));
                // FirstOrDefault方法
                var firstOrDefaultMethodExpr = Expression.Call(firstOrDefaultMethod, cellsParam, propertyEqualExpr);
                // 当前属性类型常量
                var propTypeConst = Expression.Constant(prop.PropertyType);
                // 获取单元格值表达，Cells.SingleOrDefault(c=>c.PropertyName == prop.Name).Value
                var cellValueExpr = Expression.Property(firstOrDefaultMethodExpr, typeof(ICell), "Value");
                // 变更类型
                var changeTypeExpr = Expression.Call(changeTypeMethod,
                    Expression.Convert(cellValueExpr, typeof(string)), propTypeConst);
                Expression expr = Expression.Convert(changeTypeExpr, prop.PropertyType);
                memberBindingList.Add(Expression.Bind(prop, expr));
            }

            var memberInitExpression = Expression.MemberInit(Expression.New(typeof(T)), memberBindingList.ToArray());
            var blockExpr = Expression.Block(memberInitExpression);
            var lambda = Expression.Lambda<Func<IList<ICell>, T>>(blockExpr, new[] {cellsParam});
            var func = lambda.Compile();
            Table[key] = func;
            return func;
        }

        /// <summary>
        /// 变更类型
        /// </summary>
        /// <param name="value">值</param>
        /// <param name="type">类型</param>
        public static object ChangeType(string value, Type type)
        {
            object obj = null;
            var nullableType = Nullable.GetUnderlyingType(type);
            if (nullableType != null)
            {
                if (value == null)
                    obj = null;
            }
            else if (typeof(Enum).IsAssignableFrom(type))
            {
                obj = Enum.Parse(type, value);
            }
            else
            {
                obj = Convert.ChangeType(value, type);
            }
            return obj;
        }
    }
}
