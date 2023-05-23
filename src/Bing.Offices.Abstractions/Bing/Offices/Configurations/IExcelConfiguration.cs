using System.Linq.Expressions;
using Bing.Offices.Settings;

namespace Bing.Offices.Configurations;

/// <summary>
/// Excel 配置
/// </summary>
public interface IExcelConfiguration
{
    /// <summary>
    /// 配置Excel设置
    /// </summary>
    /// <param name="configAction">配置操作</param>
    IExcelConfiguration HasExcelSetting(Action<ExcelSetting> configAction);

    /// <summary>
    /// 配置全局工作表设置
    /// </summary>
    /// <param name="setupAction">配置操作</param>
    IExcelConfiguration HasGlobalSheetSetting(Action<GlobalSheetSetting> setupAction);

    /// <summary>
    /// 配置工作表设置
    /// </summary>
    /// <param name="configAction">配置操作</param>
    /// <param name="sheetIndex">工作表索引</param>
    IExcelConfiguration HasSheetSetting(Action<SheetSetting> configAction, int sheetIndex = 0);

    /// <summary>
    /// 设置冻结窗格
    /// </summary>
    /// <param name="colSplit">冻结的列数</param>
    /// <param name="rowSplit">冻结的行数</param>
    IExcelConfiguration HasFreezePane(int colSplit, int rowSplit);

    /// <summary>
    /// 设置冻结窗格
    /// </summary>
    /// <param name="colSplit">冻结的列数</param>
    /// <param name="rowSplit">冻结的行数</param>
    /// <param name="leftMostColumn">冻结后第一列的列号</param>
    /// <param name="topRow">冻结后第一行的行号</param>
    IExcelConfiguration HasFreezePane(int colSplit, int rowSplit, int leftMostColumn, int topRow);
}

/// <summary>
/// Excel 配置
/// </summary>
/// <typeparam name="TEntity">实体类型</typeparam>
public interface IExcelConfiguration<TEntity> : IExcelConfiguration
{
    /// <summary>
    /// 设置数据校验
    /// </summary>
    /// <param name="dataValidateFunc">数据校验函数</param>
    IExcelConfiguration<TEntity> WithDataValidation(Func<TEntity, bool> dataValidateFunc);

    /// <summary>
    /// 设置属性配置
    /// </summary>
    /// <typeparam name="TProperty">属性类型</typeparam>
    /// <param name="propertyExpression">属性表达式</param>
    IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression);

    /// <summary>
    /// 设置属性配置
    /// </summary>
    /// <typeparam name="TProperty">属性类型</typeparam>
    /// <param name="propertyName">属性名</param>
    IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(string propertyName);
}
