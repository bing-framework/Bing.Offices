using System.Collections.Generic;
using System.Linq;
using Bing.Offices.Attributes;
using Bing.Offices.Decorators;

namespace Bing.Offices.Extensions
{
    /// <summary>
    /// 类型装饰器信息(<see cref="TypeDecoratorInfo"/>) 扩展
    /// </summary>
    public static class TypeDecoratorInfoExtensions
    {
        /// <summary>
        /// 获取装饰器特性
        /// </summary>
        /// <typeparam name="T">装饰器特性类型</typeparam>
        /// <param name="typeDecoratorInfo">类型装饰器信息</param>
        public static T GetDecoratorAttribute<T>(this TypeDecoratorInfo typeDecoratorInfo) where T : DecoratorAttributeBase
        {
            var attribute = typeDecoratorInfo.TypeDecorators.SingleOrDefault(x => x.GetType() == typeof(T));
            return (T) attribute;
        }

        /// <summary>
        /// 获取装饰器特性集合
        /// </summary>
        /// <typeparam name="T">装饰器特性类型</typeparam>
        /// <param name="typeDecoratorInfo">类型装饰器信息</param>
        public static List<T> GetDecoratorAttributes<T>(this TypeDecoratorInfo typeDecoratorInfo)
            where T : DecoratorAttributeBase
        {
            var attributes = typeDecoratorInfo.TypeDecorators.Where(x => x.GetType() == typeof(T));
            return attributes.Cast<T>().ToList();
        }
    }
}
