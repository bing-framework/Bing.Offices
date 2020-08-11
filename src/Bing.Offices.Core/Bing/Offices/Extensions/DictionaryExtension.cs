using System;
using System.Collections.Generic;

namespace Bing.Offices.Extensions
{
    /// <summary>
    /// 字典(<see cref="IDictionary{TKey,TValue}"/>) 扩展
    /// </summary>
    public static class DictionaryExtension
    {
        /// <summary>
        /// 添加或更新
        /// </summary>
        /// <typeparam name="TKey">键类型</typeparam>
        /// <typeparam name="TValue">值类型</typeparam>
        /// <param name="this">字典</param>
        /// <param name="key">键</param>
        /// <param name="value">值</param>
        public static TValue AddOrUpdate<TKey, TValue>(this IDictionary<TKey, TValue> @this, TKey key, TValue value)
        {
            if (!@this.ContainsKey(key))
                @this.Add(new KeyValuePair<TKey, TValue>(key, value));
            else
                @this[key] = value;
            return @this[key];
        }

        /// <summary>
        /// 添加或更新
        /// </summary>
        /// <typeparam name="TKey">键类型</typeparam>
        /// <typeparam name="TValue">值类型</typeparam>
        /// <param name="this">字典</param>
        /// <param name="key">键</param>
        /// <param name="addValue">添加值</param>
        /// <param name="updateValueFactory">更新值函数</param>
        public static TValue AddOrUpdate<TKey, TValue>(this IDictionary<TKey, TValue> @this, TKey key, TValue addValue, Func<TKey, TValue, TValue> updateValueFactory)
        {
            if (!@this.ContainsKey(key))
                @this.Add(new KeyValuePair<TKey, TValue>(key, addValue));
            else
                @this[key] = updateValueFactory(key, @this[key]);
            return @this[key];
        }

        /// <summary>
        /// 添加或更新
        /// </summary>
        /// <typeparam name="TKey">键类型</typeparam>
        /// <typeparam name="TValue">值类型</typeparam>
        /// <param name="this">字典</param>
        /// <param name="key">键</param>
        /// <param name="addValueFactory">添加值函数</param>
        /// <param name="updateValueFactory">更新值函数</param>
        public static TValue AddOrUpdate<TKey, TValue>(this IDictionary<TKey, TValue> @this, TKey key, Func<TKey, TValue> addValueFactory, Func<TKey, TValue, TValue> updateValueFactory)
        {
            if (!@this.ContainsKey(key))
                @this.Add(new KeyValuePair<TKey, TValue>(key, addValueFactory(key)));
            else
                @this[key] = updateValueFactory(key, @this[key]);
            return @this[key];
        }
    }
}
