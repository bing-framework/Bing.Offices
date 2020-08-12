using System;
using System.Linq.Expressions;
using Bing.Offices.Metadata;

namespace Bing.Offices.FluentApi
{
    /// <summary>
    /// Excel配置
    /// </summary>
    public interface IExcelFluent
    {
        /// <summary>
        /// 设置工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="startRowIndex">起始行索引</param>
        /// <param name="enableAutoColumnWidth">是否启用自动列宽</param>
        IExcelFluent HasSheetConfiguration(int sheetIndex, string sheetName, int startRowIndex, bool enableAutoColumnWidth);

        /// <summary>
        /// 设置工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="action">操作</param>
        IExcelFluent HasSheet(int sheetIndex, Action<ISheetFluent> action);

        /// <summary>
        /// 获取工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        ISheetFluent Sheet(int sheetIndex);

        /// <summary>
        /// 设置类模式
        /// </summary>
        /// <param name="mode">类模式</param>
        IExcelFluent HasModel(ClassMode mode = ClassMode.List);

        #region ExcelSettings FluentAPI

        /// <summary>
        /// 设置Excel作者
        /// </summary>
        /// <param name="author">作者</param>
        IExcelFluent HasAuthor(string author);

        /// <summary>
        /// 设置Excel标题
        /// </summary>
        /// <param name="title">标题</param>
        IExcelFluent HasTitle(string title);

        /// <summary>
        /// 设置Excel描述
        /// </summary>
        /// <param name="description">描述</param>
        IExcelFluent HasDescription(string description);

        /// <summary>
        /// 设置Excel主题
        /// </summary>
        /// <param name="subject">主题</param>
        IExcelFluent HasSubject(string subject);

        /// <summary>
        /// 设置Excel公司
        /// </summary>
        /// <param name="company">公司</param>
        IExcelFluent HasCompany(string company);

        /// <summary>
        /// 设置Excel类别
        /// </summary>
        /// <param name="category">类别</param>
        IExcelFluent HasCategory(string category);

        #endregion
    }

    /// <summary>
    /// Excel配置
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    public interface IExcelFluent<TEntity> : IExcelFluent
    {
        /// <summary>
        /// 属性配置
        /// </summary>
        /// <typeparam name="TProperty">属性类型</typeparam>
        /// <param name="propertyExpression">属性表达式</param>
        IPropertyFluent<TEntity, TProperty> Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression);

        /// <summary>
        /// 属性配置
        /// </summary>
        /// <typeparam name="TProperty">属性类型</typeparam>
        /// <param name="propertyName">属性表达式</param>
        IPropertyFluent<TEntity, TProperty> Property<TProperty>(string propertyName);
    }
}
