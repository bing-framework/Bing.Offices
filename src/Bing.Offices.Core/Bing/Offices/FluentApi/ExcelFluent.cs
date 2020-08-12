using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using Bing.Offices.Extensions;
using Bing.Offices.Metadata;
using Bing.Offices.Settings;

namespace Bing.Offices.FluentApi
{
    /// <summary>
    /// Excel配置
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    internal sealed class ExcelFluent<TEntity> : IExcelFluent<TEntity>
    {
        /// <summary>
        /// 类元数据
        /// </summary>
        public ClassMetadata Metadata { get; set; }

        /// <summary>
        /// 工作表设置 字典
        /// </summary>
        public IDictionary<int, SheetFluent> SheetFluents { get; set; }

        /// <summary>
        /// 属性设置 字典
        /// </summary>
        public IDictionary<PropertyInfo, PropertyFluent> PropertyFluents { get; set; }

        /// <summary>
        /// 初始化一个<see cref="ExcelFluent{TEntity}"/>类型的实例
        /// </summary>
        public ExcelFluent() : this(null) { }

        /// <summary>
        /// 初始化一个<see cref="ExcelFluent{TEntity}"/>类型的实例
        /// </summary>
        /// <param name="setting">Excel文档属性设置</param>
        public ExcelFluent(ExcelSetting setting)
        {
            Metadata = new ClassMetadata(typeof(TEntity), setting);
            SheetFluents = new Dictionary<int, SheetFluent>
            {
                {0, new SheetFluent(Metadata.GetSheetMetadata())}
            };
            PropertyFluents = new Dictionary<PropertyInfo, PropertyFluent>();
        }

        /// <summary>
        /// 设置类模式
        /// </summary>
        /// <param name="mode">类模式</param>
        public IExcelFluent HasModel(ClassMode mode = ClassMode.List)
        {
            Metadata.Mode = mode;
            return this;
        }

        #region ExcelSettings FluentAPI

        /// <summary>
        /// 设置Excel作者
        /// </summary>
        /// <param name="author">作者</param>
        public IExcelFluent HasAuthor(string author)
        {
            Metadata.ExcelSetting.Author = author;
            return this;
        }

        /// <summary>
        /// 设置Excel标题
        /// </summary>
        /// <param name="title">标题</param>
        public IExcelFluent HasTitle(string title)
        {
            Metadata.ExcelSetting.Title = title;
            return this;
        }

        /// <summary>
        /// 设置Excel描述
        /// </summary>
        /// <param name="description">描述</param>
        public IExcelFluent HasDescription(string description)
        {
            Metadata.ExcelSetting.Description = description;
            return this;
        }

        /// <summary>
        /// 设置Excel主题
        /// </summary>
        /// <param name="subject">主题</param>
        public IExcelFluent HasSubject(string subject)
        {
            Metadata.ExcelSetting.Subject = subject;
            return this;
        }

        /// <summary>
        /// 设置Excel公司
        /// </summary>
        /// <param name="company">公司</param>
        public IExcelFluent HasCompany(string company)
        {
            Metadata.ExcelSetting.Company = company;
            return this;
        }

        /// <summary>
        /// 设置Excel类别
        /// </summary>
        /// <param name="category">类别</param>
        public IExcelFluent HasCategory(string category)
        {
            Metadata.ExcelSetting.Category = category;
            return this;
        }

        #endregion

        #region Property(属性)

        /// <summary>
        /// 属性配置
        /// </summary>
        /// <typeparam name="TProperty">属性类型</typeparam>
        /// <param name="propertyExpression">属性表达式</param>
        public IPropertyFluent<TEntity, TProperty> Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression)
        {
            var memberInfo = propertyExpression.GetMemberInfo();
            var property = memberInfo as PropertyInfo;
            if (property == null || !PropertyFluents.ContainsKey(property))
            {
                property = InternalCache.TypePropertyCache.GetOrAdd(Metadata.ClassType, t => t.GetProperties()).FirstOrDefault(x => x.Name == memberInfo.Name);
                if (null == property)
                    throw new InvalidOperationException($"the property [{memberInfo.Name}] does not exists.");
            }

            return PropertyFluents[property] as IPropertyFluent<TEntity, TProperty>;
        }

        /// <summary>
        /// 属性配置
        /// </summary>
        /// <typeparam name="TProperty">属性类型</typeparam>
        /// <param name="propertyName">属性表达式</param>
        public IPropertyFluent<TEntity, TProperty> Property<TProperty>(string propertyName)
        {
            var property = PropertyFluents.Keys.FirstOrDefault(x => x.Name == propertyName);
            if (property != null)
                return PropertyFluents[property] as IPropertyFluent<TEntity, TProperty>;

            var propertyType = typeof(TProperty);
            property = new FakePropertyInfo(Metadata.ClassType, propertyType, propertyName);
            var propertyFluentType = typeof(PropertyFluent<,>).MakeGenericType(Metadata.ClassType, propertyType);
            var propertyFluent = (PropertyFluent) Activator.CreateInstance(propertyFluentType,
                new object[] {Metadata.GetPropertyMetadata(property)});
            PropertyFluents[property] = propertyFluent;
            return propertyFluent as IPropertyFluent<TEntity, TProperty>;
        }

        #endregion

        #region Sheet(工作表)

        /// <summary>
        /// 设置工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="startRowIndex">起始行索引</param>
        /// <param name="enableAutoColumnWidth">是否启用自动列宽</param>
        public IExcelFluent HasSheetConfiguration(int sheetIndex, string sheetName, int startRowIndex,
            bool enableAutoColumnWidth)
        {
            if (sheetIndex >= 0)
            {
                if (SheetFluents.TryGetValue(sheetIndex, out var sheetFluent))
                {
                    sheetFluent.Metadata.Name = sheetName;
                    sheetFluent.Metadata.StartRowIndex = startRowIndex;
                    sheetFluent.Metadata.AutoColumnWidthEnabled = enableAutoColumnWidth;
                }
                else
                {
                    SheetFluents[sheetIndex] = new SheetFluent(Metadata.GetSheetMetadata(sheetIndex, () =>
                        new SheetMetadata
                        {
                            Index = sheetIndex,
                            Name = sheetName,
                            StartRowIndex = startRowIndex,
                            AutoColumnWidthEnabled = enableAutoColumnWidth
                        }));
                }
            }

            return this;
        }

        /// <summary>
        /// 设置工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="action">操作</param>
        public IExcelFluent HasSheet(int sheetIndex, Action<ISheetFluent> action)
        {
            if (sheetIndex < 0)
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), $"工作表索引不能小于0");
            if (!SheetFluents.ContainsKey(sheetIndex))
                SheetFluents[sheetIndex] = new SheetFluent(Metadata.GetSheetMetadata(sheetIndex));
            action.Invoke(SheetFluents[sheetIndex]);
            return this;
        }

        /// <summary>
        /// 获取工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        public ISheetFluent Sheet(int sheetIndex) => GetSheet(sheetIndex);

        /// <summary>
        /// 获取工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        internal SheetFluent GetSheet(int sheetIndex)
        {
            if (sheetIndex < 0)
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), $"工作表索引不能小于0");
            if (SheetFluents.TryGetValue(sheetIndex, out var sheetMetadata))
                return sheetMetadata;
            SheetFluents[sheetIndex] = new SheetFluent(Metadata.GetSheetMetadata(sheetIndex));
            return SheetFluents[sheetIndex];
        }

        #endregion
    }
}
