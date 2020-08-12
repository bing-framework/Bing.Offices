using System;
using System.Collections.Generic;
using System.Reflection;
using Bing.Offices.Attributes;
using Bing.Offices.FluentApi;
using Bing.Offices.Metadata;

namespace Bing.Offices
{
    /// <summary>
    /// 内部帮助类
    /// </summary>
    internal static class InternalHelper
    {
        /// <summary>
        /// 获取Excel映射设置
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        public static ExcelFluent<TEntity> GetExcelFluentMapping<TEntity>() =>
            (ExcelFluent<TEntity>) InternalCache.TypeFluentCache.GetOrAdd(typeof(TEntity), GetExcelFluent<TEntity>);

        /// <summary>
        /// 获取Excel设置
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        /// <param name="type">类型</param>
        private static IExcelFluent GetExcelFluent<TEntity>(Type type)
        {
            var excelFluent = new ExcelFluent<TEntity>();
            // 初始化 工作表 特性
            foreach (var sheetAttribute in type.GetCustomAttributes<SheetAttribute>())
            {
                if (sheetAttribute.Index >= 0)
                    excelFluent.GetSheet(sheetAttribute.Index).Metadata = sheetAttribute.SheetMetadata;
            }

            var dict = new Dictionary<PropertyInfo, PropertyFluent>();
            var propertyInfos = InternalCache.TypePropertyCache.GetOrAdd(type, t => t.GetProperties());
            foreach (var propertyInfo in propertyInfos)
            {
                var column = propertyInfo.GetCustomAttribute<ColumnAttribute>() ?? new ColumnAttribute();
                if (string.IsNullOrWhiteSpace(column.Title))
                    column.Title = propertyInfo.Name;
                var propertyFluentType = typeof(PropertyFluent<,>).MakeGenericType(type, propertyInfo.PropertyType);
                var propertyMetadata = excelFluent.Metadata.GetPropertyMetadata(propertyInfo);
                var propertyFluent = Activator.CreateInstance(propertyFluentType,
                    new object[] {propertyMetadata});
                InitPropertyMetadata(propertyMetadata, column);
                dict.Add(propertyInfo, (PropertyFluent) propertyFluent);
            }

            excelFluent.PropertyFluents = dict;

            return excelFluent;
        }

        /// <summary>
        /// 初始化属性元数据
        /// </summary>
        /// <param name="metadata">属性元数据</param>
        /// <param name="attribute">列属性</param>
        private static void InitPropertyMetadata(PropertyMetadata metadata, ColumnAttribute attribute)
        {
            metadata.Column.Title = attribute.Title;
            metadata.Column.ColumnIndex = attribute.Index;
            metadata.Column.Formatter = attribute.Formatter;
            metadata.IsIgnored = attribute.IsIgnored;
            metadata.Column.Style.Width = attribute.Width;
        }
    }
}
