using System.Globalization;
using System.Reflection;
using AspectCore.Extensions.Reflection;

namespace Bing.Offices.Internals;

/// <summary>
/// 伪属性信息
/// </summary>
internal sealed class FakePropertyInfo : PropertyInfo
{
    /// <summary>
    /// 获取值函数
    /// </summary>
    private readonly Func<object> _getValueFunc;

    /// <summary>
    /// 值
    /// </summary>
    private readonly object _value;

    /// <summary>
    /// 初始化一个<see cref="FakePropertyInfo"/>类型的实例
    /// </summary>
    /// <param name="entityType">实体类型</param>
    /// <param name="propertyType">属性类型</param>
    /// <param name="propertyName">属性名</param>
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

    /// <inheritdoc />
    public override Type DeclaringType { get; }

    /// <inheritdoc />
    public override string Name { get; }

    /// <inheritdoc />
    public override Type ReflectedType { get; }

    /// <inheritdoc />
    public override bool CanRead => false;

    /// <inheritdoc />
    public override bool CanWrite => false;

    /// <inheritdoc />
    public override Type PropertyType { get; }

    /// <inheritdoc />
    public override PropertyAttributes Attributes { get; }

    /// <inheritdoc />
    public override MethodInfo GetGetMethod(bool nonPublic) => _getValueFunc.Method;

    /// <inheritdoc />
    public override MethodInfo GetSetMethod(bool nonPublic) => null;

    /// <inheritdoc />
    public override string ToString() => $"{PropertyType.Name}, {Name}";

    /// <inheritdoc />
    public override object[] GetCustomAttributes(bool inherit) => throw new NotSupportedException();

    /// <inheritdoc />
    public override object[] GetCustomAttributes(Type attributeType, bool inherit) => throw new NotSupportedException();

    /// <inheritdoc />
    public override bool IsDefined(Type attributeType, bool inherit) => throw new NotSupportedException();

    /// <inheritdoc />
    public override MethodInfo[] GetAccessors(bool nonPublic) => throw new NotSupportedException();

    /// <inheritdoc />
    public override ParameterInfo[] GetIndexParameters() => throw new NotSupportedException();

    /// <inheritdoc />
    public override object GetValue(object obj, BindingFlags invokeAttr, Binder binder, object[] index, CultureInfo culture) => _value;

    /// <inheritdoc />
    public override void SetValue(object obj, object value, BindingFlags invokeAttr, Binder binder, object[] index, CultureInfo culture) => throw new NotSupportedException();

}
