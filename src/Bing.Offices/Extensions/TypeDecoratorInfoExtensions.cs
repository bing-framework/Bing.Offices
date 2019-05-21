using System.Linq;
using Bing.Offices.Attributes.Decorators;
using Bing.Offices.Decorators;

namespace Bing.Offices.Extensions
{
    /// <summary>
    /// 类型装饰器信息扩展
    /// </summary>
    public static class TypeDecoratorInfoExtensions
    {
        /// <summary>
        /// 获取装饰器特性
        /// </summary>
        /// <typeparam name="T">装饰器特性类型</typeparam>
        /// <param name="typeDecoratorInfo">类型装饰器信息</param>
        /// <returns></returns>
        public static T GetDecorate<T>(this TypeDecoratorInfo typeDecoratorInfo) where T : DecorateBaseAttribute
        {
            var attribute = typeDecoratorInfo?.TypeDecorates.SingleOrDefault(x => x.GetType() == typeof(T));
            return attribute as T;
        }
    }
}
