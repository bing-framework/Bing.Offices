using System;
using System.Linq.Expressions;
using Bing.Offices.Excel.Settings;

namespace Bing.Offices.Excel.Configurations
{
    /// <summary>
    /// Excel配置
    /// </summary>
    public interface IExcelConfiguration
    {
        /// <summary>
        /// Excel设置
        /// </summary>
        ExcelSetting ExcelSetting { get; }

        /// <summary>
        /// 设置工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="sheetName">工作表名称</param>
        IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName);

        /// <summary>
        /// 设置工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="enableAutoColumnWidth">启用自动列宽</param>
        IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName, bool enableAutoColumnWidth);

        /// <summary>
        /// 设置工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="startRowIndex">起始行索引</param>
        IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName, int startRowIndex);

        /// <summary>
        /// 设置工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="startRowIndex">起始行索引</param>
        /// <param name="enableAutoColumnWidth">启用自动列宽</param>
        IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName, int startRowIndex, bool enableAutoColumnWidth);

        /// <summary>
        /// 设置冻结面板窗格。创建一个拆分(冻结窗格)，任何现有的冻结窗格或拆分窗格都将被覆盖。
        /// </summary>
        /// <param name="colSplit">水平分割位置</param>
        /// <param name="rowSplit">垂直分割位置</param>
        IExcelConfiguration HasFreezePane(int colSplit, int rowSplit);

        /// <summary>
        /// 设置冻结面板窗格。创建一个拆分(冻结窗格)，任何现有的冻结窗格或拆分窗格都将被覆盖。
        /// </summary>
        /// <param name="colSplit">水平分割位置</param>
        /// <param name="rowSplit">垂直分割位置</param>
        /// <param name="leftMostColumn">顶列索引</param>
        /// <param name="topRow">顶行索引</param>
        IExcelConfiguration HasFreezePane(int colSplit, int rowSplit, int leftMostColumn, int topRow);

        /// <summary>
        /// 设置过滤器
        /// </summary>
        /// <param name="firstColumn">首列索引</param>
        IExcelConfiguration HasFilter(int firstColumn);

        /// <summary>
        /// 设置过滤器
        /// </summary>
        /// <param name="firstColumn">首列索引</param>
        /// <param name="lastColumn">最后一列索引</param>
        IExcelConfiguration HasFilter(int firstColumn, int? lastColumn);

        /// <summary>
        /// 设置作者
        /// </summary>
        /// <param name="author">作者</param>
        IExcelConfiguration HasAuthor(string author);

        /// <summary>
        /// 设置标题
        /// </summary>
        /// <param name="title">标题</param>
        IExcelConfiguration HasTitle(string title);

        /// <summary>
        /// 设置描述
        /// </summary>
        /// <param name="description">描述</param>
        IExcelConfiguration HasDescription(string description);

        /// <summary>
        /// 设置主题
        /// </summary>
        /// <param name="subject">主题</param>
        IExcelConfiguration HasSubject(string subject);

        /// <summary>
        /// 设置公司
        /// </summary>
        /// <param name="company">公司</param>
        IExcelConfiguration HasCompany(string company);

        /// <summary>
        /// 设置类别
        /// </summary>
        /// <param name="category">类别</param>
        IExcelConfiguration HasCategory(string category);
    }

    /// <summary>
    /// Excel配置
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    public interface IExcelConfiguration<TEntity> : IExcelConfiguration
    {
        /// <summary>
        /// 设置属性
        /// </summary>
        /// <typeparam name="TProperty">属性类型</typeparam>
        /// <param name="propertyExpression">属性表达式</param>
        IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression);
    }
}
