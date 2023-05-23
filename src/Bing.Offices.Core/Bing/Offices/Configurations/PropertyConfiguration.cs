using System.Reflection;
using Bing.Offices.Exports;
using Bing.Offices.Internals;
using Bing.Offices.Settings;

namespace Bing.Offices.Configurations;

/// <summary>
/// 属性配置
/// </summary>
internal class PropertyConfiguration : IPropertyConfiguration
{
    /// <summary>
    /// 列属性设置
    /// </summary>
    public ExportColumnPropertySetting ColumnPropertySetting { get; }

    /// <summary>
    /// 初始化一个<see cref="PropertyConfiguration"/>类型的实例
    /// </summary>
    /// <param name="property">属性信息</param>
    public PropertyConfiguration(PropertyInfo property)
    {
        ColumnPropertySetting = new ExportColumnPropertySetting(property);
    }

    /// <summary>
    /// 构建
    /// </summary>
    public void Buid()
    {
        // 对于伪属性信息，无法获取其中的特性
        if (ColumnPropertySetting.PropertyInfo is FakePropertyInfo)
            return;
        ColumnPropertySetting.InitByAttribute();
    }
}

/// <summary>
/// 属性配置
/// </summary>
/// <typeparam name="TEntity">实体类型</typeparam>
/// <typeparam name="TProperty">属性类型</typeparam>
internal sealed class PropertyConfiguration<TEntity, TProperty> : PropertyConfiguration, IPropertyConfiguration<TEntity, TProperty>
{
    /// <summary>
    /// 属性信息
    /// </summary>
    private readonly PropertyInfo _propertyInfo;

    /// <summary>
    /// 初始化一个<see cref="PropertyConfiguration"/>类型的实例
    /// </summary>
    /// <param name="property">属性信息</param>
    public PropertyConfiguration(PropertyInfo property) : base(property)
    {
        _propertyInfo = property;
    }

    /// <summary>
    /// 设置列索引
    /// </summary>
    /// <param name="index">索引</param>
    public IPropertyConfiguration<TEntity, TProperty> HasColumnIndex(int index)
    {
        if (index >= 0)
            ColumnPropertySetting.ColumnIndex = index;
        return this;
    }

    /// <summary>
    /// 设置列宽
    /// </summary>
    /// <param name="width">宽度</param>
    public IPropertyConfiguration<TEntity, TProperty> HasColumnWidth(int width)
    {
        ColumnPropertySetting.ColumnWidth = width;
        return this;
    }

    /// <summary>
    /// 设置列标题
    /// </summary>
    /// <param name="title">标题</param>
    public IPropertyConfiguration<TEntity, TProperty> HasColumnTitle(string title)
    {
        ColumnPropertySetting.Name = title ?? throw new ArgumentNullException(nameof(title));
        return this;
    }

    /// <summary>
    /// 设置列格式化
    /// </summary>
    /// <param name="formatter">格式化文本</param>
    public IPropertyConfiguration<TEntity, TProperty> HasColumnFormatter(string formatter)
    {
        ColumnPropertySetting.StringFormat = formatter;
        return this;
    }

    /// <summary>
    /// 设置表头样式
    /// </summary>
    /// <param name="styleSetup">样式配置</param>
    public IPropertyConfiguration<TEntity, TProperty> HasHeaderStyle(Action<IBaseStyle> styleSetup)
    {
        throw new NotImplementedException();
    }

    /// <summary>
    /// 设置内容样式
    /// </summary>
    /// <param name="styleSetup">样式配置</param>
    public IPropertyConfiguration<TEntity, TProperty> HasBodyStyle(Action<IBaseStyle> styleSetup)
    {
        throw new NotImplementedException();
    }

    /// <summary>
    /// 设置表头样式
    /// </summary>
    /// <param name="style">样式</param>
    public IPropertyConfiguration<TEntity, TProperty> HasHeaderStyle(IBaseStyle style)
    {
        ColumnPropertySetting.HeaderStyle = style;
        return this;
    }

    /// <summary>
    /// 设置内容样式
    /// </summary>
    /// <param name="style">样式配置</param>
    public IPropertyConfiguration<TEntity, TProperty> HasBodyStyle(IBaseStyle style)
    {
        ColumnPropertySetting.ColumnStyle = style;
        return this;
    }

    /// <summary>
    /// 设置默认值
    /// </summary>
    /// <param name="value">默认值</param>
    public IPropertyConfiguration<TEntity, TProperty> HasDefaultValue(object value)
    {
        ColumnPropertySetting.DefaultValue = value;
        return this;
    }

    /// <summary>
    /// 设置忽略属性
    /// </summary>
    /// <param name="ignored">是否忽略。默认：true</param>
    public IPropertyConfiguration<TEntity, TProperty> Ignored(bool ignored = true)
    {
        ColumnPropertySetting.Ignored = ignored;
        return this;
    }

    /// <summary>
    /// 设置列输入格式化
    /// </summary>
    /// <param name="formatterFunc">格式化函数</param>
    public IPropertyConfiguration<TEntity, TProperty> HasColumnInputFormatter(Func<string, TProperty> formatterFunc)
    {
        InternalCache.ColumnInputFormatterFunCache.AddOrUpdate(_propertyInfo, formatterFunc, (_, _) => formatterFunc);
        return this;
    }

    /// <summary>
    /// 设置输出格式化
    /// </summary>
    /// <param name="formatterFunc">格式化函数</param>
    public IPropertyConfiguration<TEntity, TProperty> HasOutputFormatter(Func<TEntity, TProperty, object> formatterFunc)
    {
        InternalCache.OutputFormatterFuncCache.AddOrUpdate(_propertyInfo, formatterFunc, (_, _) => formatterFunc);
        return this;
    }

    /// <summary>
    /// 设置输入格式化
    /// </summary>
    /// <param name="formatterFunc">格式化函数</param>
    public IPropertyConfiguration<TEntity, TProperty> HasInputFormatter(Func<TEntity, TProperty, TProperty> formatterFunc)
    {
        InternalCache.InputFormatterFuncCache.AddOrUpdate(_propertyInfo, formatterFunc, (_, _) => formatterFunc);
        return this;
    }
}
