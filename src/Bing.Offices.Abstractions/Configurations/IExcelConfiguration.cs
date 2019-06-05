using System;
using System.Linq.Expressions;
using Bing.Offices.Abstractions.Settings;

namespace Bing.Offices.Abstractions.Configurations
{
    /// <summary>
    /// Excel 配置
    /// </summary>
    public interface IExcelConfiguration
    {
        /// <summary>
        /// Excel 文档属性设置
        /// </summary>
        ExcelSetting ExcelSetting { get; }

        #region ExcelSetting(Excel文档属性设置)

        /// <summary>
        /// 设置作者
        /// </summary>
        /// <param name="author">作者</param>
        IExcelConfiguration HasAuthor(string author);

        /// <summary>
        /// 设置公司
        /// </summary>
        /// <param name="company">公司</param>
        IExcelConfiguration HasCompany(string company);

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
        /// 设置目录
        /// </summary>
        /// <param name="category">目录</param>
        IExcelConfiguration HasCategory(string category);

        #endregion

        #region FreezePane(冻结窗格)

        /// <summary>
        /// 设置冻结区域
        /// </summary>
        /// <param name="colSplit">冻结单元格列号</param>
        /// <param name="rowSplit">冻结单元格行号</param>
        IExcelConfiguration HasFreezePane(int colSplit, int rowSplit);

        /// <summary>
        /// 设置冻结区域
        /// </summary>
        /// <param name="colSplit">冻结单元格列号</param>
        /// <param name="rowSplit">冻结单元格行号</param>
        /// <param name="leftMostColumn">顶列索引</param>
        /// <param name="topRow">顶行索引</param>
        IExcelConfiguration HasFreezePane(int colSplit, int rowSplit, int leftMostColumn, int topRow);

        #endregion

        #region Filter(筛选器)

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

        #endregion

        #region Sheet(工作表设置)

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
        /// <param name="startRowIndex">起始行索引</param>
        IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName, int startRowIndex);

        #endregion
    }

    /// <summary>
    /// Excel 配置
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    public interface IExcelConfiguration<TEntity> : IExcelConfiguration
    {
        /// <summary>
        /// 设置属性
        /// </summary>
        /// <typeparam name="TProperty">属性类型</typeparam>
        /// <param name="propertyExpression">属性表达式</param>
        /// <returns></returns>
        IPropertyConfiguration Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression);
    }
}
