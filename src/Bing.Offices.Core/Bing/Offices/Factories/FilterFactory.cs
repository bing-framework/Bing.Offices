using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Bing.Offices.Attributes;
using Bing.Offices.Filters;
using Bing.Extensions;

namespace Bing.Offices.Factories
{
    /// <summary>
    /// 过滤器工厂
    /// </summary>
    internal static class FilterFactory
    {
        /// <summary>
        /// 过滤器字典
        /// </summary>
        private static readonly IDictionary<Type, IFilter> FilterDict = new ConcurrentDictionary<Type, IFilter>();

        /// <summary>
        /// 过滤器集合字典
        /// </summary>
        private static readonly IDictionary<Type, IList<IFilter>> FiltersDict = new ConcurrentDictionary<Type, IList<IFilter>>();

        /// <summary>
        /// 创建过滤器实例
        /// </summary>
        /// <param name="type">绑定过滤器的特性类型s</param>
        public static IFilter CreateInstance(Type type)
        {
            if (FilterDict.ContainsKey(type))
                return FilterDict[type];
            var filterType = Assembly.GetAssembly(type).GetTypes().ToList()
                ?.Where(t => typeof(IFilter).IsAssignableFrom(t))?.FirstOrDefault(t =>
                    t.IsDefined(typeof(BindFilterAttribute)) &&
                    t.GetCustomAttribute<BindFilterAttribute>()?.FilterType == type);
            if (filterType == null)
                throw new ArgumentNullException(nameof(filterType), "找不到指定过滤器类型");
            var filter = Activator.CreateInstance(filterType) as IFilter;
            FilterDict[type] = filter;
            return filter;
        }

        /// <summary>
        /// 创建过滤器实例集合
        /// </summary>
        /// <typeparam name="T">实体类型</typeparam>
        public static IList<IFilter> CreateInstances<T>()
        {
            var templateType = typeof(T);
            if (FiltersDict.ContainsKey(templateType))
                return FiltersDict[templateType];
            var filters = new List<IFilter>();
            var filterAttributes = new List<FilterAttributeBase>();
            var typeFilterInfo = TypeFilterInfoFactory.CreateInstance(templateType);
            typeFilterInfo.PropertyFilterInfos.ForEach(x => x.Filters.ForEach(y => filterAttributes.Add(y)));
            filterAttributes.Distinct(new FilterAttributeComparer()).ToList().ForEach(x =>
            {
                var filter = CreateInstance(x.GetType());
                if (filter != null)
                    filters.Add(filter);
            });
            FiltersDict[templateType] = filters;
            return filters;
        }
    }
}
