using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;
using Bing.Offices.Excel.Settings;

namespace Bing.Offices.Excel.Configurations
{
    /// <summary>
    /// Excel配置
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    internal class ExcelConfiguration<TEntity> : IExcelConfiguration<TEntity>
    {
        /// <summary>
        /// 实体类型
        /// </summary>
        private readonly Type _entityType = typeof(TEntity);

        /// <summary>
        /// 属性配置字典
        /// </summary>
        public IDictionary<PropertyInfo, PropertyConfiguration> PropertyConfigurationDictionary { get; internal set; }

        /// <summary>
        /// Excel设置
        /// </summary>
        public ExcelSetting ExcelSetting { get; }

        /// <summary>
        /// 冻结窗格设置
        /// </summary>
        internal IList<FreezeSetting> FreezeSettings { get; set; }

        /// <summary>
        /// 过滤器设置
        /// </summary>
        internal FilterSetting FilterSetting { get; set; }

        /// <summary>
        /// 工作表设置
        /// </summary>
        internal IDictionary<int, SheetSetting> SheetSettings { get; set; }

        /// <summary>
        /// 初始化一个<see cref="ExcelConfiguration{TEntity}"/>类型的实例
        /// </summary>
        public ExcelConfiguration() : this(null) { }

        /// <summary>
        /// 初始化一个<see cref="ExcelConfiguration{TEntity}"/>类型的实例
        /// </summary>
        /// <param name="setting">Excel设置</param>
        public ExcelConfiguration(ExcelSetting setting)
        {
            PropertyConfigurationDictionary = new Dictionary<PropertyInfo, PropertyConfiguration>();
            ExcelSetting = (setting ?? new ExcelSetting());
            SheetSettings = new Dictionary<int, SheetSetting>(4)
            {
                {0, new SheetSetting()}
            };
            FreezeSettings = new List<FreezeSetting>();
        }


        /// <summary>
        /// 设置属性
        /// </summary>
        /// <typeparam name="TProperty">属性类型</typeparam>
        /// <param name="propertyExpression">属性表达式</param>
        public IPropertyConfiguration<TEntity, TProperty> Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 设置工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="sheetName">工作表名称</param>
        public IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 设置工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="enableAutoColumnWidth">启用自动列宽</param>
        public IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName, bool enableAutoColumnWidth)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 设置工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="startRowIndex">起始行索引</param>
        public IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName, int startRowIndex)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 设置工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="startRowIndex">起始行索引</param>
        /// <param name="enableAutoColumnWidth">启用自动列宽</param>
        public IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName, int startRowIndex,
            bool enableAutoColumnWidth)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 设置冻结面板窗格。创建一个拆分(冻结窗格)，任何现有的冻结窗格或拆分窗格都将被覆盖。
        /// </summary>
        /// <param name="colSplit">水平分割位置</param>
        /// <param name="rowSplit">垂直分割位置</param>
        public IExcelConfiguration HasFreezePane(int colSplit, int rowSplit)
        {
            FreezeSettings.Add(new FreezeSetting(colSplit, rowSplit));
            return this;
        }

        /// <summary>
        /// 设置冻结面板窗格。创建一个拆分(冻结窗格)，任何现有的冻结窗格或拆分窗格都将被覆盖。
        /// </summary>
        /// <param name="colSplit">水平分割位置</param>
        /// <param name="rowSplit">垂直分割位置</param>
        /// <param name="leftMostColumn">顶列索引</param>
        /// <param name="topRow">顶行索引</param>
        public IExcelConfiguration HasFreezePane(int colSplit, int rowSplit, int leftMostColumn, int topRow)
        {
            FreezeSettings.Add(new FreezeSetting(colSplit, rowSplit, leftMostColumn, topRow));
            return this;
        }

        /// <summary>
        /// 设置过滤器
        /// </summary>
        /// <param name="firstColumn">首列索引</param>
        public IExcelConfiguration HasFilter(int firstColumn) => HasFilter(firstColumn, null);

        /// <summary>
        /// 设置过滤器
        /// </summary>
        /// <param name="firstColumn">首列索引</param>
        /// <param name="lastColumn">最后一列索引</param>
        public IExcelConfiguration HasFilter(int firstColumn, int? lastColumn)
        {
            FilterSetting = new FilterSetting(firstColumn, lastColumn);
            return this;
        }

        /// <summary>
        /// 设置作者
        /// </summary>
        /// <param name="author">作者</param>
        public IExcelConfiguration HasAuthor(string author)
        {
            ExcelSetting.Author = author;
            return this;
        }

        /// <summary>
        /// 设置标题
        /// </summary>
        /// <param name="title">标题</param>
        public IExcelConfiguration HasTitle(string title)
        {
            ExcelSetting.Title = title;
            return this;
        }

        /// <summary>
        /// 设置描述
        /// </summary>
        /// <param name="description">描述</param>
        public IExcelConfiguration HasDescription(string description)
        {
            ExcelSetting.Description = description;
            return this;
        }

        /// <summary>
        /// 设置主题
        /// </summary>
        /// <param name="subject">主题</param>
        public IExcelConfiguration HasSubject(string subject)
        {
            ExcelSetting.Subject = subject;
            return this;
        }

        /// <summary>
        /// 设置公司
        /// </summary>
        /// <param name="company">公司</param>
        public IExcelConfiguration HasCompany(string company)
        {
            ExcelSetting.Company = company;
            return this;
        }

        /// <summary>
        /// 设置类别
        /// </summary>
        /// <param name="category">类别</param>
        public IExcelConfiguration HasCategory(string category)
        {
            ExcelSetting.Category = category;
            return this;
        }
    }
}
