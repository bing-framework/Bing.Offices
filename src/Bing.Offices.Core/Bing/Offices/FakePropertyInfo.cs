using System;
using System.Globalization;
using System.Reflection;
using AspectCore.Extensions.Reflection;

namespace Bing.Offices
{
    /// <summary>
    /// 虚假属性信息
    /// </summary>
    internal sealed class FakePropertyInfo : PropertyInfo
    {
        /// <summary>
        /// 值
        /// </summary>
        private readonly object _value;

        /// <summary>
        /// 获取值函数
        /// </summary>
        private readonly Func<object> _getValueFunc;

        /// <summary>
        /// 初始化一个<see cref="FakePropertyInfo"/>类型的实例
        /// </summary>
        /// <param name="entityType">实体类型</param>
        /// <param name="propertyType">属性类型</param>
        /// <param name="propertyName">属性名称</param>
        public FakePropertyInfo(Type entityType, Type propertyType, string propertyName)
        {
            DeclaringType = entityType;
            ReflectedType = entityType;
            PropertyType = propertyType;
            Name = propertyName;
            _value = propertyType.GetDefaultValue();
            _getValueFunc = () => _value;
            Attributes = PropertyAttributes.None;
        }

        /// <summary>
        /// 类的声明类型
        /// </summary>
        public override Type DeclaringType { get; }

        /// <summary>
        /// 名称
        /// </summary>
        public override string Name { get; }

        /// <summary>
        /// 反射类型
        /// </summary>
        public override Type ReflectedType { get; }

        /// <summary>
        /// 能否读取
        /// </summary>
        public override bool CanRead => false;

        /// <summary>
        /// 能否写入
        /// </summary>
        public override bool CanWrite => false;

        /// <summary>
        /// 属性类型
        /// </summary>
        public override Type PropertyType { get; }

        /// <summary>
        /// 属性特性
        /// </summary>
        public override PropertyAttributes Attributes { get; }

        /// <summary>
        /// 获取get方法
        /// </summary>
        public override MethodInfo GetGetMethod(bool nonPublic) => _getValueFunc.Method;

        /// <summary>
        /// 获取set方法
        /// </summary>
        public override MethodInfo GetSetMethod(bool nonPublic) => null;

        /// <summary>
        /// 输出字符串
        /// </summary>
        public override string ToString() => $"{PropertyType.Name}, {Name}";

        /// <summary>
        /// 获取自定义特性数组
        /// </summary>
        public override object[] GetCustomAttributes(bool inherit) => throw new NotSupportedException();

        /// <summary>
        /// 获取自定义特性数组
        /// </summary>
        public override object[] GetCustomAttributes(Type attributeType, bool inherit) => throw new NotSupportedException();

        /// <summary>
        /// 获取索引参数
        /// </summary>
        public override ParameterInfo[] GetIndexParameters() => throw new NotSupportedException();

        /// <summary>
        /// 是否定义指定特性类型
        /// </summary>
        public override bool IsDefined(Type attributeType, bool inherit) => throw new NotSupportedException();

        /// <summary>
        /// 获取访问器
        /// </summary>
        public override MethodInfo[] GetAccessors(bool nonPublic) => throw new NotSupportedException();

        /// <summary>
        /// 获取值
        /// </summary>
        public override object GetValue(object obj, BindingFlags invokeAttr, Binder binder, object[] index, CultureInfo culture) => _value;

        /// <summary>
        /// 设置值
        /// </summary>
        public override void SetValue(object obj, object value, BindingFlags invokeAttr, Binder binder, object[] index, CultureInfo culture) => throw new NotSupportedException();
    }
}
