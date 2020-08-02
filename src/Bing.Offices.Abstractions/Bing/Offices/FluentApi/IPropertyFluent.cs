namespace Bing.Offices.FluentApi
{
    /// <summary>
    /// 属性配置
    /// </summary>
    public interface IPropertyFluent
    {
    }

    /// <summary>
    /// 属性配置
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <typeparam name="TProperty">属性类型</typeparam>
    public interface IPropertyFluent<out TEntity, TProperty> : IPropertyFluent
    {

    }
}
