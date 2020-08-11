using System;
using System.Linq.Expressions;

namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 定义类元数据
    /// </summary>
    public interface IClassMetadata
    {
        /// <summary>
        /// 设置工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="startRowIndex">起始行索引</param>
        /// <param name="enableAutoColumnWidth">是否启用自动列宽</param>
        IClassMetadata HasSheetConfiguration(int sheetIndex, string sheetName, int startRowIndex, bool enableAutoColumnWidth);

        /// <summary>
        /// 获取工作表元数据
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        ISheetMetadata Sheet(int sheetIndex);

        /// <summary>
        /// 设置类模式
        /// </summary>
        /// <param name="mode">类模式</param>
        IClassMetadata HasMode(ClassMode mode = ClassMode.List);

        #region ExcelSettings FluentAPI

        /// <summary>
        /// 设置Excel作者
        /// </summary>
        /// <param name="author">作者</param>
        IClassMetadata HasAuthor(string author);

        /// <summary>
        /// 设置Excel标题
        /// </summary>
        /// <param name="title">标题</param>
        IClassMetadata HasTitle(string title);

        /// <summary>
        /// 设置Excel描述
        /// </summary>
        /// <param name="description">描述</param>
        IClassMetadata HasDescription(string description);

        /// <summary>
        /// 设置Excel主题
        /// </summary>
        /// <param name="subject">主题</param>
        IClassMetadata HasSubject(string subject);

        /// <summary>
        /// 设置Excel公司
        /// </summary>
        /// <param name="company">公司</param>
        IClassMetadata HasCompany(string company);

        /// <summary>
        /// 设置Excel类别
        /// </summary>
        /// <param name="category">类别</param>
        IClassMetadata HasCategory(string category);

        #endregion
    }

    /// <summary>
    /// 定义类元数据
    /// </summary>
    /// <typeparam name="TEntity">对象类型</typeparam>
    public interface IClassMetadata<TEntity> : IClassMetadata
    {
        /// <summary>
        /// 获取属性元数据
        /// </summary>
        /// <typeparam name="TProperty">属性类型</typeparam>
        /// <param name="propertyExpression">属性表达式</param>
        IPropertyMetadata<TEntity, TProperty> Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression);

        /// <summary>
        /// 获取属性元数据
        /// </summary>
        /// <typeparam name="TProperty">属性类型</typeparam>
        /// <param name="propertyName">属性名称</param>
        IPropertyMetadata<TEntity, TProperty> Property<TProperty>(string propertyName);
    }
}
