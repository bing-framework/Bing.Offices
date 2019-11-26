using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;

namespace Bing.Offices.Core.Extension
{
    /// <summary>
    /// 类型扩展
    /// </summary>
    public static class TypeExtensions
    {
        /// <summary>
        /// 获取显示名称
        /// </summary>
        /// <param name="customAttributeProvider">自定义特性提供程序</param>
        /// <param name="inherit">是否继承</param>
        public static string GetDisplayName(this ICustomAttributeProvider customAttributeProvider, bool inherit = false)
        {
            string displayName = null;
            var displayAttribute = customAttributeProvider.GetAttribute<DisplayAttribute>();
            if (displayAttribute != null)
            {
                displayName = displayAttribute.Name;
            }
            else
            {
                var displayNameAttribute = customAttributeProvider.GetAttribute<DisplayNameAttribute>();
                if (displayNameAttribute != null)
                    displayName = displayNameAttribute.DisplayName;
            }
            return displayName;
        }

        /// <summary>
        /// 获取类型描述
        /// </summary>
        /// <param name="customAttributeProvider">自定义特性提供程序</param>
        /// <param name="inherit">是否继承</param>
        public static string GetDescription(this ICustomAttributeProvider customAttributeProvider, bool inherit = false)
        {
            var desc = string.Empty;
            var descAttribute = customAttributeProvider.GetAttribute<DescriptionAttribute>();
            if (descAttribute != null)
                desc = descAttribute.Description;
            return desc;
        }

        /// <summary>
        /// 获取类型描述或显示名称
        /// </summary>
        /// <param name="customAttributeProvider">自定义特性提供程序</param>
        /// <param name="inherit">是否继承</param>
        public static string GetTypeDisplayOrDescription(this ICustomAttributeProvider customAttributeProvider,
            bool inherit = false)
        {
            var display = customAttributeProvider.GetDescription(inherit);
            if (string.IsNullOrWhiteSpace(display))
                display = customAttributeProvider.GetDisplayName(inherit);
            return display ?? string.Empty;
        }

        /// <summary>
        /// 获取特性
        /// </summary>
        /// <typeparam name="T">特性类型</typeparam>
        /// <param name="customAttributeProvider">自定义特性提供程序</param>
        /// <param name="inherit">是否继承</param>
        public static T GetAttribute<T>(this ICustomAttributeProvider customAttributeProvider, bool inherit = false)
            where T : Attribute =>
            customAttributeProvider
                .GetCustomAttributes(typeof(T), inherit)
                .OfType<T>()
                .FirstOrDefault();

        /// <summary>
        /// 获取特性列表
        /// </summary>
        /// <typeparam name="T">特性类型</typeparam>
        /// <param name="customAttributeProvider">自定义特性提供程序</param>
        /// <param name="inherit">是否继承</param>
        public static IEnumerable<T> GetAttributes<T>(this ICustomAttributeProvider customAttributeProvider,
            bool inherit = false) where T : Attribute =>
            customAttributeProvider
                .GetCustomAttributes(typeof(T), inherit)
                .OfType<T>();

        /// <summary>
        /// 检查指定类型成员中是否存在指定的特性
        /// </summary>
        /// <typeparam name="T">特性类型</typeparam>
        /// <param name="customAttributeProvider">自定义特性提供程序</param>
        /// <param name="inherit">是否继承</param>
        public static bool ExistsAttribute<T>(this ICustomAttributeProvider customAttributeProvider,
            bool inherit = false) where T : Attribute => customAttributeProvider
            .GetCustomAttributes(typeof(T), inherit)
            .Any(m => m is T);

        /// <summary>
        /// 是否必填
        /// </summary>
        /// <param name="propertyInfo">属性信息</param>
        public static bool IsRequired(this PropertyInfo propertyInfo)
        {
            if (propertyInfo.GetAttribute<RequiredAttribute>(true) != null)
                return true;
            //Boolean、Byte、SByte、Int16、UInt16、Int32、UInt32、Int64、UInt64、Char、Double、Single
            if (propertyInfo.PropertyType.IsPrimitive)
                return true;
            switch (propertyInfo.PropertyType.Name)
            {
                case "DateTime":
                case "Decimal":
                    return true;
            }
            return false;
        }

        /// <summary>
        /// 获取当前程序集中应用此特性的类
        /// </summary>
        /// <typeparam name="TAttribute">特性类型</typeparam>
        /// <param name="assembly">程序集</param>
        /// <param name="inherit">是否继承</param>
        public static IEnumerable<Type> GetTypesWith<TAttribute>(this Assembly assembly, bool inherit)
            where TAttribute : Attribute
        {
            var attributeType = typeof(TAttribute);
            foreach (var type in assembly.GetTypes().Where(type => type.GetCustomAttributes(attributeType, true).Length > 0))
                yield return type;
        }

        /// <summary>
        /// 获取枚举定义列表
        /// </summary>
        /// <param name="type">枚举类型</param>
        public static List<(string Name, int Value, string DisplayName, string Description)> GetEnumDefinitionList(
            this Type type)
        {
            var list = new List<(string Name, int Value, string DisplayName, string Description)>();
            var attributeType = type;
            if (!attributeType.IsEnum)
                return null;
            var names = Enum.GetNames(attributeType);
            var values = Enum.GetValues(attributeType);
            var index = 0;
            foreach (var value in values)
            {
                var name = names[index];
                var field = value.GetType().GetField(values.ToString());
                var displayName = field.GetDisplayName();
                var desc = field.GetDescription();
                (string Name, int Value, string DisplayName, string Description) item = (name, Convert.ToInt32(value),
                    string.IsNullOrWhiteSpace(displayName) ? null : displayName,
                    string.IsNullOrWhiteSpace(desc) ? null : desc);
                list.Add(item);
                index++;
            }

            return list;
        }

    }
}
