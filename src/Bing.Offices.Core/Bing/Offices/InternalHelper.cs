using System;
using System.Collections.Generic;
using System.Reflection;
using Bing.Offices.Attributes;
using Bing.Offices.Metadata;

namespace Bing.Offices
{
    /// <summary>
    /// 内部帮助类
    /// </summary>
    internal static class InternalHelper
    {
        /// <summary>
        /// 获取类元数据
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        public static ClassMetadata<TEntity> GetClassMetadataMapping<TEntity>() =>
            (ClassMetadata<TEntity>) InternalCache.TypeClassMetadataCache.GetOrAdd(typeof(TEntity), GetClassMetadata<TEntity>);

        /// <summary>
        /// 获取类元数据
        /// </summary>
        /// <typeparam name="TEntity">实体类型</typeparam>
        /// <param name="type">类型</param>
        private static IClassMetadata GetClassMetadata<TEntity>(Type type)
        {
            var classMetadata = new ClassMetadata<TEntity>();
            // 初始化 工作表 特性
            foreach (var sheetAttribute in type.GetCustomAttributes<SheetAttribute>())
            {
                if (sheetAttribute.Index >= 0)
                    classMetadata.SheetMetadataDict[sheetAttribute.Index] = sheetAttribute.SheetMetadata;
            }

            var dict = new Dictionary<PropertyInfo, PropertyMetadata>();
            var propertyInfos = InternalCache.TypePropertyCache.GetOrAdd(type, t => t.GetProperties());
            foreach (var propertyInfo in propertyInfos)
            {
                var column = propertyInfo.GetCustomAttribute<ColumnAttribute>() ?? new ColumnAttribute();
                if (string.IsNullOrWhiteSpace(column.Title))
                    column.Title = propertyInfo.Name;

            }

            return classMetadata;
        }
    }
}
