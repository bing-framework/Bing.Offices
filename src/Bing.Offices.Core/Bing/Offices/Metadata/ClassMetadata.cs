using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using Bing.Offices.Extensions;
using Bing.Offices.Settings;

namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 类元数据
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    internal sealed class ClassMetadata<TEntity> : IClassMetadata<TEntity>
    {
        /// <summary>
        /// 类类型
        /// </summary>
        public Type ClassType => typeof(TEntity);

        /// <summary>
        /// 类模式
        /// </summary>
        public ClassMode Mode { get; set; }

        /// <summary>
        /// 属性配置字典
        /// </summary>
        public IDictionary<PropertyInfo, PropertyMetadata> PropertyConfigurationDictionary { get; internal set; }

        /// <summary>
        /// Excel 文档属性设置
        /// </summary>
        public ExcelSetting ExcelSetting { get; }

        /// <summary>
        /// 工作表元数据 字典
        /// </summary>
        internal IDictionary<int, SheetMetadata> SheetMetadataDict { get; set; }

        /// <summary>
        /// 初始化一个<see cref="ClassMetadata{TEntity}"/>类型的实例
        /// </summary>
        public ClassMetadata() : this(null) { }

        /// <summary>
        /// 初始化一个<see cref="ClassMetadata{TEntity}"/>类型的实例
        /// </summary>
        /// <param name="setting">Excel 文档属性设置</param>
        public ClassMetadata(ExcelSetting setting)
        {
            PropertyConfigurationDictionary = new Dictionary<PropertyInfo, PropertyMetadata>();
            ExcelSetting = (setting ?? ExcelSetting.Default) ?? new ExcelSetting();
            SheetMetadataDict = new Dictionary<int, SheetMetadata>
            {
                {0, new SheetMetadata()}
            };
        }

        /// <summary>
        /// 设置类模式
        /// </summary>
        /// <param name="mode">类模式</param>
        public IClassMetadata HasMode(ClassMode mode = ClassMode.List)
        {
            Mode = mode;
            return this;
        }

        #region ExcelSettings FluentAPI

        /// <summary>
        /// 设置Excel作者
        /// </summary>
        /// <param name="author">作者</param>
        public IClassMetadata HasAuthor(string author)
        {
            ExcelSetting.Author = author;
            return this;
        }

        /// <summary>
        /// 设置Excel标题
        /// </summary>
        /// <param name="title">标题</param>
        public IClassMetadata HasTitle(string title)
        {
            ExcelSetting.Title = title;
            return this;
        }

        /// <summary>
        /// 设置Excel描述
        /// </summary>
        /// <param name="description">描述</param>
        public IClassMetadata HasDescription(string description)
        {
            ExcelSetting.Description = description;
            return this;
        }

        /// <summary>
        /// 设置Excel主题
        /// </summary>
        /// <param name="subject">主题</param>
        public IClassMetadata HasSubject(string subject)
        {
            ExcelSetting.Subject = subject;
            return this;
        }

        /// <summary>
        /// 设置Excel公司
        /// </summary>
        /// <param name="company">公司</param>
        public IClassMetadata HasCompany(string company)
        {
            ExcelSetting.Company = company;
            return this;
        }

        /// <summary>
        /// 设置Excel类别
        /// </summary>
        /// <param name="category">类别</param>
        public IClassMetadata HasCategory(string category)
        {
            ExcelSetting.Category = category;
            return this;
        }

        #endregion

        #region Property(属性)

        /// <summary>
        /// 获取属性元数据
        /// </summary>
        /// <typeparam name="TProperty">属性类型</typeparam>
        /// <param name="propertyExpression">属性表达式</param>
        public IPropertyMetadata<TEntity, TProperty> Property<TProperty>(Expression<Func<TEntity, TProperty>> propertyExpression)
        {
            var memberInfo = propertyExpression.GetMemberInfo();
            var property = memberInfo as PropertyInfo;
            if (property == null || PropertyConfigurationDictionary.ContainsKey(property))
            {
                property = InternalCache.TypePropertyCache.GetOrAdd(ClassType, t => t.GetProperties()).FirstOrDefault(x => x.Name == memberInfo.Name);
                if (null == property)
                    throw new InvalidOperationException($"the property [{memberInfo.Name}] does not exists.");
            }

            return PropertyConfigurationDictionary[property] as IPropertyMetadata<TEntity, TProperty>;
        }

        /// <summary>
        /// 获取属性元数据
        /// </summary>
        /// <typeparam name="TProperty">属性类型</typeparam>
        /// <param name="propertyName">属性名称</param>
        public IPropertyMetadata<TEntity, TProperty> Property<TProperty>(string propertyName)
        {
            var property = PropertyConfigurationDictionary.Keys.FirstOrDefault(x => x.Name == propertyName);
            if (property != null)
                return PropertyConfigurationDictionary[property] as IPropertyMetadata<TEntity, TProperty>;

            var propertyType = typeof(TProperty);
            property = new FakePropertyInfo(ClassType, propertyType, propertyName);
            var propertyMetadataType = typeof(PropertyMetadata<,>).MakeGenericType(ClassType, propertyType);
            var propertyMetadata = (PropertyMetadata)Activator.CreateInstance(propertyMetadataType, new object[] { property });
            PropertyConfigurationDictionary[property] = propertyMetadata;
            return propertyMetadata as IPropertyMetadata<TEntity, TProperty>;
        }

        #endregion

        #region Sheet(工作表)

        /// <summary>
        /// 获取工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        public ISheetMetadata Sheet(int sheetIndex)
        {
            if (sheetIndex < 0)
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), $"工作表索引不能小于0");
            if (SheetMetadataDict.TryGetValue(sheetIndex, out var sheetMetadata))
                return sheetMetadata;
            SheetMetadataDict[sheetIndex] = new SheetMetadata
            {
                Index = sheetIndex
            };
            return SheetMetadataDict[sheetIndex];
        }

        /// <summary>
        /// 工作表配置
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="sheetName">工作表名称</param>
        /// <param name="startRowIndex">起始行索引</param>
        /// <param name="enableAutoColumnWidth">是否启用自动列宽</param>
        public IClassMetadata HasSheetConfiguration(int sheetIndex, string sheetName, int startRowIndex, bool enableAutoColumnWidth)
        {
            if (sheetIndex >= 0)
            {
                if (SheetMetadataDict.TryGetValue(sheetIndex, out var sheetMetadata))
                {
                    sheetMetadata.Name = sheetName;
                    sheetMetadata.StartRowIndex = startRowIndex;
                    sheetMetadata.AutoColumnWidthEnabled = enableAutoColumnWidth;
                }
                else
                {
                    SheetMetadataDict[sheetIndex] = new SheetMetadata
                    {
                        Index = sheetIndex,
                        Name = sheetName,
                        StartRowIndex = startRowIndex,
                        AutoColumnWidthEnabled = enableAutoColumnWidth
                    };
                }
            }
            return this;
        }

        #endregion
    }
}
