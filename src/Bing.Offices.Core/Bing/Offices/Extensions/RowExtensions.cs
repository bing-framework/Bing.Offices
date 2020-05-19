using System;
using System.Collections.Generic;
using System.Linq;
using Bing.Offices.Helpers;
using Bing.Extensions;
using Bing.Offices.Metadata.Excels;

namespace Bing.Offices.Extensions
{
    /// <summary>
    /// 单元行(<see cref="IRow"/>) 扩展
    /// </summary>
    public static class RowExtensions
    {
        /// <summary>
        /// 设置行校验结果
        /// </summary>
        /// <param name="row">单元行</param>
        /// <param name="valid">是否校验通过</param>
        /// <param name="cell">单元格</param>
        /// <param name="errorMsg">错误消息</param>
        public static void Valid(this IRow row, bool valid, ICell cell, string errorMsg)
        {
            if (valid)
                return;
            row.ErrorMsg += $"{cell.Name}{errorMsg};";
        }

        /// <summary>
        /// 转换
        /// </summary>
        /// <typeparam name="T">对象类型</typeparam>
        /// <param name="rows">单元行集合</param>
        public static IEnumerable<T> Convert<T>(this IEnumerable<IRow> rows) => rows?.ConvertByExpressTree<T>();

        /// <summary>
        /// 通过表达式树，将单元行集合快速转换为指定类型集合
        /// </summary>
        /// <typeparam name="T">对象类型</typeparam>
        /// <param name="rows">单元行集合</param>
        private static IEnumerable<T> ConvertByExpressTree<T>(this IEnumerable<IRow> rows)
        {
            if (rows == null || !rows.Any())
                return null;
            Func<IList<ICell>, T> func = GetFunc<T>(rows.ToList()[0]);
            var list = new List<T>();
            rows.ToList().ForEach(row =>
            {
                var item = ExpressionMapper.FastConvert(row, func);
                list.Add(item);
            });
            return list;
        }

        /// <summary>
        /// 获取转换函数
        /// </summary>
        /// <typeparam name="T">类型</typeparam>
        /// <param name="row">单元行</param>
        private static Func<IList<ICell>, T> GetFunc<T>(IRow row)
        {
            var propertyNames = string.Empty;
            row.Cells.ForEach(cell => propertyNames += cell.PropertyName + "_");
            var key = typeof(T).FullName + "_" + propertyNames.Trim('_');
            var props = typeof(T).GetProperties().Where(x => x.CanWrite && x.CanRead && !x.HasIgnore());
            var func = ExpressionMapper.GetFunc<T>(key, props);
            return func;
        }
    }
}
