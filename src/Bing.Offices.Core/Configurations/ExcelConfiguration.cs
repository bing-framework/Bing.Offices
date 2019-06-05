using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using Bing.Offices.Abstractions.Configurations;
using Bing.Offices.Abstractions.Settings;
using Bing.Offices.Extensions;
using Bing.Offices.Settings;

namespace Bing.Offices.Configurations
{
    /// <summary>
    /// Excel 配置
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    internal class ExcelConfiguration<TEntity> : IExcelConfiguration<TEntity>
    {
        /// <summary>
        /// 实体类型
        /// </summary>
        private readonly Type _entityType = typeof(TEntity);

        /// <summary>
        /// Excel 文档属性设置
        /// </summary>
        private readonly ExcelSetting _excelSetting;

        /// <summary>
        /// 工作表设置
        /// </summary>
        private readonly IList<SheetSetting> _sheetSettings;

        /// <summary>
        /// 属性配置字典
        /// </summary>
        public IDictionary<PropertyInfo, IPropertyConfiguration> PropertyConfigurationsDictionary { get; internal set; }

        /// <summary>
        /// Excel 文档属性设置
        /// </summary>
        public IExcelSetting ExcelSetting => _excelSetting;

        /// <summary>
        /// 冻结设置
        /// </summary>
        public IList<IFreezeSetting> FreezeSettings { get; internal set; }

        /// <summary>
        /// 过滤器设置
        /// </summary>
        public IFilterSetting FilterSetting { get; internal set; }

        /// <summary>
        /// 工作表设置
        /// </summary>
        public IList<ISheetSetting> SheetSettings => _sheetSettings.Select(x => (ISheetSetting) x).ToList();

        /// <summary>
        /// 初始化一个<see cref="ExcelConfiguration{TEntity}"/>类型的实例
        /// </summary>
        public ExcelConfiguration() : this(null) { }

        /// <summary>
        /// 初始化一个<see cref="ExcelConfiguration{TEntity}"/>类型的实例
        /// </summary>
        /// <param name="setting">Excel 文档属性设置</param>
        public ExcelConfiguration(ExcelSetting setting)
        {
            PropertyConfigurationsDictionary = new Dictionary<PropertyInfo, IPropertyConfiguration>();
            _excelSetting = (setting ?? DefaultSettings.DefaultExcelSetting) ?? new ExcelSetting();
            _sheetSettings = new List<SheetSetting>(InternalConst.MaxSheetNum/16)
            {
                new SheetSetting()
            };
            FreezeSettings = new List<IFreezeSetting>();
        }

        #region ExcelSetting(Excel文档属性设置)

        /// <summary>
        /// 设置作者
        /// </summary>
        /// <param name="author">作者</param>
        public IExcelConfiguration HasAuthor(string author)
        {
            _excelSetting.Author = author;
            return this;
        }

        /// <summary>
        /// 设置公司
        /// </summary>
        /// <param name="company">公司</param>
        public IExcelConfiguration HasCompany(string company)
        {
            _excelSetting.Company = company;
            return this;
        }

        /// <summary>
        /// 设置标题
        /// </summary>
        /// <param name="title">标题</param>
        public IExcelConfiguration HasTitle(string title)
        {
            _excelSetting.Title = title;
            return this;
        }

        /// <summary>
        /// 设置描述
        /// </summary>
        /// <param name="description">描述</param>
        public IExcelConfiguration HasDescription(string description)
        {
            _excelSetting.Description = description;
            return this;
        }

        /// <summary>
        /// 设置主题
        /// </summary>
        /// <param name="subject">主题</param>
        public IExcelConfiguration HasSubject(string subject)
        {
            _excelSetting.Subject = subject;
            return this;
        }

        /// <summary>
        /// 设置目录
        /// </summary>
        /// <param name="category">目录</param>
        public IExcelConfiguration HasCategory(string category)
        {
            _excelSetting.Category = category;
            return this;
        }

        #endregion

        #region FreezePane(冻结窗格)

        /// <summary>
        /// 设置冻结区域
        /// </summary>
        /// <param name="colSplit">冻结单元格列号</param>
        /// <param name="rowSplit">冻结单元格行号</param>
        public IExcelConfiguration HasFreezePane(int colSplit, int rowSplit)
        {
            FreezeSettings.Add(new FreezeSetting(colSplit, rowSplit));
            return this;
        }

        /// <summary>
        /// 设置冻结区域
        /// </summary>
        /// <param name="colSplit">冻结单元格列号</param>
        /// <param name="rowSplit">冻结单元格行号</param>
        /// <param name="leftMostColumn">顶列索引</param>
        /// <param name="topRow">顶行索引</param>
        public IExcelConfiguration HasFreezePane(int colSplit, int rowSplit, int leftMostColumn, int topRow)
        {
            FreezeSettings.Add(new FreezeSetting(colSplit, rowSplit, leftMostColumn, topRow));
            return this;
        }

        #endregion

        #region Filter(筛选器)

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

        #endregion

        #region Sheet(工作表设置)

        /// <summary>
        /// 设置工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="sheetName">工作表名称</param>
        public IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName) =>
            HasSheetConfiguration(sheetIndex, sheetName, 1);

        /// <summary>
        /// 设置工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="startRowIndex">起始行索引</param>
        public IExcelConfiguration HasSheetConfiguration(int sheetIndex, string sheetName, int startRowIndex)
        {
            var sheetSetting = _sheetSettings.FirstOrDefault(_ => _.Index == sheetIndex);
            if (sheetSetting == null)
            {
                _sheetSettings.Add(new SheetSetting()
                {
                    Index = sheetIndex,
                    Name = sheetName,
                    StartRowIndex = startRowIndex
                });
            }
            else
            {
                sheetSetting.Name = sheetName;
                sheetSetting.StartRowIndex = startRowIndex;
            }

            return this;
        }

        #endregion

        /// <summary>
        /// 设置属性
        /// </summary>
        /// <typeparam name="TProperty">属性类型</typeparam>
        /// <param name="propertyExpression">属性表达式</param>
        /// <returns></returns>
        public IPropertyConfiguration Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression) =>
            PropertyConfigurationsDictionary[_entityType.GetProperty(propertyExpression.GetMemberInfo().Name)];
    }
}
