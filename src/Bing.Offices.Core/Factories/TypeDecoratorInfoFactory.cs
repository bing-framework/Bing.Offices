using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Bing.Offices.Abstractions.Attributes;
using Bing.Offices.Abstractions.Decorators;
using Bing.Offices.Attributes;
using Bing.Extensions;

namespace Bing.Offices.Factories
{
    /// <summary>
    /// 类型装饰器信息工厂
    /// </summary>
    public static class TypeDecoratorInfoFactory
    {
        /// <summary>
        /// 类型装饰器字典
        /// </summary>
        private static readonly IDictionary<Type, TypeDecoratorInfo> TypeDecoratorDict = new ConcurrentDictionary<Type, TypeDecoratorInfo>();

        /// <summary>
        /// 创建类型装饰器信息实例
        /// </summary>
        /// <param name="type">类型</param>
        public static TypeDecoratorInfo CreateInstance(Type type)
        {
            if (type == null)
                throw new ArgumentNullException(nameof(type));
            if (TypeDecoratorDict.ContainsKey(type))
                return TypeDecoratorDict[type];
            var typeDecoratorInfo = new TypeDecoratorInfo();
            // 全局装饰器特性
            typeDecoratorInfo.TypeDecorators.AddRange(type.GetCustomAttributes<DecoratorAttributeBase>());
            // 列装饰器特性
            var props = type.GetProperties().Where(x => x.IsDefined(typeof(ColumnNameAttribute))).ToList();
            for (int i = 0; i < props.Count; i++)
            {
                typeDecoratorInfo.PropertyDecoratorInfos.Add(new PropertyDecoratorInfo()
                {
                    ColumnIndex = i,
                    Decorators = props[i].GetCustomAttributes<DecoratorAttributeBase>().ToList()
                });
            }
            TypeDecoratorDict[type] = typeDecoratorInfo;
            return typeDecoratorInfo;
        }
    }
}
