using System.Linq.Expressions;
using System.Reflection;
using Bing.Expressions;
using Bing.Offices.Internals;
using Bing.Offices.Settings;
using Bing.Reflection;

namespace Bing.Offices.Configurations;

/// <summary>
/// Excel 配置
/// </summary>
internal abstract class ExcelConfiguration : IExcelConfiguration
{
    /// <summary>
    /// 属性配置字典
    /// </summary>
    public IDictionary<PropertyInfo, PropertyConfiguration> PropertyConfigurationDictionary { get; } = new Dictionary<PropertyInfo, PropertyConfiguration>();

    /// <summary>
    /// Excel 设置
    /// </summary>
    public ExcelSetting ExcelSetting { get; } = ExcelSetting.Default;

    /// <summary>
    /// 全局工作表设置
    /// </summary>
    public GlobalSheetSetting GlobalSheetSetting { get; } = new();

    /// <summary>
    /// 工作表设置
    /// </summary>
    public IDictionary<int, SheetSetting> SheetSettings { get; } = new Dictionary<int, SheetSetting> { { 0, new SheetSetting() } };

    /// <summary>
    /// 冻结设置
    /// </summary>
    public IList<FreezeSetting> FreezeSettings { get; } = new List<FreezeSetting>();

    /// <summary>
    /// 配置Excel设置
    /// </summary>
    /// <param name="configAction">配置操作</param>
    public IExcelConfiguration HasExcelSetting(Action<ExcelSetting> configAction)
    {
        configAction?.Invoke(ExcelSetting);
        return this;
    }

    /// <summary>
    /// 配置全局工作表设置
    /// </summary>
    /// <param name="setupAction">配置操作</param>
    public IExcelConfiguration HasGlobalSheetSetting(Action<GlobalSheetSetting> setupAction)
    {
        setupAction?.Invoke(GlobalSheetSetting);
        return this;
    }

    /// <summary>
    /// 配置工作表设置
    /// </summary>
    /// <param name="configAction">配置操作</param>
    /// <param name="sheetIndex">工作表索引</param>
    public IExcelConfiguration HasSheetSetting(Action<SheetSetting> configAction, int sheetIndex = 0)
    {
        if (configAction == null)
            throw new ArgumentNullException(nameof(configAction));
        if (sheetIndex >= 0)
        {
            if (!SheetSettings.TryGetValue(sheetIndex, out var sheetSetting))
            {
                sheetSetting = new SheetSetting();
                SheetSettings[sheetIndex] = sheetSetting;
            }
            configAction.Invoke(sheetSetting);
        }

        return this;
    }

    /// <summary>
    /// 设置冻结窗格
    /// </summary>
    /// <param name="colSplit">冻结的列数</param>
    /// <param name="rowSplit">冻结的行数</param>
    public IExcelConfiguration HasFreezePane(int colSplit, int rowSplit)
    {
        FreezeSettings.Add(new FreezeSetting(colSplit, rowSplit));
        return this;
    }

    /// <summary>
    /// 设置冻结窗格
    /// </summary>
    /// <param name="colSplit">冻结的列数</param>
    /// <param name="rowSplit">冻结的行数</param>
    /// <param name="leftMostColumn">冻结后第一列的列号</param>
    /// <param name="topRow">冻结后第一行的行号</param>
    public IExcelConfiguration HasFreezePane(int colSplit, int rowSplit, int leftMostColumn, int topRow)
    {
        FreezeSettings.Add(new FreezeSetting(colSplit, rowSplit, leftMostColumn, topRow));
        return this;
    }
}


/// <summary>
/// Excel 配置
/// </summary>
/// <typeparam name="TEntity">实体类型</typeparam>
internal sealed class ExcelConfiguration<TEntity> : ExcelConfiguration, IExcelConfiguration<TEntity>
{
    /// <summary>
    /// 实体类型
    /// </summary>
    public Type EntityType => typeof(TEntity);

    /// <summary>
    /// 数据校验函数
    /// </summary>
    internal Func<TEntity, bool> DataValidationFunc { get; private set; }

    /// <summary>
    /// 设置数据校验
    /// </summary>
    /// <param name="dataValidateFunc">数据校验函数</param>
    public IExcelConfiguration<TEntity> WithDataValidation(Func<TEntity, bool> dataValidateFunc)
    {
        DataValidationFunc = dataValidateFunc;
        return this;
    }

    /// <summary>
    /// 设置属性配置
    /// </summary>
    /// <typeparam name="TProperty">属性类型</typeparam>
    /// <param name="propertyExpression">属性表达式</param>
    public IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression)
    {
        var memberInfo = Lambdas.GetMember(propertyExpression);
        var property = memberInfo as PropertyInfo;
        if (property == null || PropertyConfigurationDictionary.ContainsKey(property))
        {
            property = TypeReflections.TypeCacheManager.GetTypeProperties(EntityType)
                .FirstOrDefault(p => p.Name == memberInfo.Name);
            if (null == property)
                throw new InvalidOperationException($"this property [{memberInfo.Name}] does not exists!");
        }
        return (IPropertyConfiguration<TEntity, TProperty>)PropertyConfigurationDictionary[property];
    }

    /// <summary>
    /// 设置属性配置
    /// </summary>
    /// <typeparam name="TProperty">属性类型</typeparam>
    /// <param name="propertyName">属性名</param>
    public IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(string propertyName)
    {
        var property = PropertyConfigurationDictionary.Keys.FirstOrDefault(p => p.Name == propertyName);
        if (property != null)
            return (IPropertyConfiguration<TEntity, TProperty>)PropertyConfigurationDictionary[property];
        var propertyType = typeof(TProperty);
        property = new FakePropertyInfo(EntityType, propertyType, propertyName);
        var propertyConfigurationType = typeof(PropertyConfiguration<,>).MakeGenericType(EntityType, propertyType);
        var propertyConfiguration = (PropertyConfiguration)Activator.CreateInstance(propertyConfigurationType, property);
        PropertyConfigurationDictionary[property] = propertyConfiguration;
        return (IPropertyConfiguration<TEntity, TProperty>)PropertyConfigurationDictionary[property];
    }
}
