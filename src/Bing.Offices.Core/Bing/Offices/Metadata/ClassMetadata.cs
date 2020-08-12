using System;
using System.Collections.Generic;
using System.Reflection;
using Bing.Offices.Settings;

namespace Bing.Offices.Metadata
{
    /// <summary>
    /// 类元数据
    /// </summary>
    internal sealed class ClassMetadata
    {
        /// <summary>
        /// 类类型
        /// </summary>
        public Type ClassType { get; }

        /// <summary>
        /// 类模式
        /// </summary>
        public ClassMode Mode { get; set; }

        /// <summary>
        /// Excel 文档属性设置
        /// </summary>
        public ExcelSetting ExcelSetting { get; }

        /// <summary>
        /// 属性元数据 字典
        /// </summary>
        public IDictionary<PropertyInfo, PropertyMetadata> PropertyMetadataDict { get; internal set; }

        /// <summary>
        /// 工作表元数据 字典
        /// </summary>
        public IDictionary<int, SheetMetadata> SheetMetadataDict { get; internal set; }

        /// <summary>
        /// 初始化一个<see cref="ClassMetadata"/>类型的实例
        /// </summary>
        /// <param name="classType">类类型</param>
        public ClassMetadata(Type classType) : this(classType, null) { }

        /// <summary>
        /// 初始化一个<see cref="ClassMetadata"/>类型的实例
        /// </summary>
        /// <param name="classType">类类型</param>
        /// <param name="setting">Excel文档属性设置</param>
        public ClassMetadata(Type classType, ExcelSetting setting)
        {
            ClassType = classType;
            ExcelSetting = (setting ?? ExcelSetting.Default) ?? new ExcelSetting();
            PropertyMetadataDict = new Dictionary<PropertyInfo, PropertyMetadata>();
            SheetMetadataDict = new Dictionary<int, SheetMetadata>
            {
                {0, new SheetMetadata()}
            };
            PropertyMetadataDict = new Dictionary<PropertyInfo, PropertyMetadata>();
        }

        /// <summary>
        /// 获取工作表元数据
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        public SheetMetadata GetSheetMetadata(int sheetIndex = 0)
        {
            if (!SheetMetadataDict.ContainsKey(sheetIndex))
                SheetMetadataDict[sheetIndex] = new SheetMetadata
                {
                    Index = sheetIndex
                };
            return SheetMetadataDict[sheetIndex];
        }

        /// <summary>
        /// 获取工作表元数据
        /// </summary>
        /// <param name="sheetIndex">工作表索引</param>
        /// <param name="func">获取函数</param>
        public SheetMetadata GetSheetMetadata(int sheetIndex, Func<SheetMetadata> func)
        {
            if (!SheetMetadataDict.ContainsKey(sheetIndex))
                SheetMetadataDict[sheetIndex] = func();
            return SheetMetadataDict[sheetIndex];
        }

        /// <summary>
        /// 获取属性元数据
        /// </summary>
        /// <param name="propertyInfo">属性信息</param>
        public PropertyMetadata GetPropertyMetadata(PropertyInfo propertyInfo)
        {
            if (!PropertyMetadataDict.ContainsKey(propertyInfo))
                PropertyMetadataDict[propertyInfo] = new PropertyMetadata(propertyInfo);
            return PropertyMetadataDict[propertyInfo];
        }
    }
}
