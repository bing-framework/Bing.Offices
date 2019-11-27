using System;

namespace Bing.Offices.Excel.Settings
{
    /// <summary>
    /// 属性设置
    /// </summary>
    internal class PropertySetting
    {
        /// <summary>
        /// 列索引
        /// </summary>
        public int ColumnIndex { get; set; } = -1;

        /// <summary>
        /// 列宽
        /// </summary>
        public int ColumnWidth { get; set; }

        /// <summary>
        /// 列标题
        /// </summary>
        public string ColumnTitle { get; set; }

        /// <summary>
        /// 列格式化
        /// </summary>
        public string ColumnFormatter { get; set; }

        /// <summary>
        /// 是否忽略该属性
        /// </summary>
        public bool IsIgnored { get; set; }

        /// <summary>
        /// 是否动态列
        /// </summary>
        public bool IsDynamicColumn { get; set; }

        /// <summary>
        /// 默认值
        /// </summary>
        public object DefaultValue { get; set; }
    }

    /// <summary>
    /// 属性设置
    /// </summary>
    /// <typeparam name="TEntity">实体类型</typeparam>
    /// <typeparam name="TProperty">属性类型</typeparam>
    internal class PropertySetting<TEntity, TProperty> : PropertySetting
    {
        /// <summary>
        /// 输入格式化函数
        /// </summary>
        public Func<TEntity, TProperty, TProperty> InputFormatterFunc { get; set; }

        /// <summary>
        /// 输出格式化函数
        /// </summary>
        public Func<TEntity, TProperty, object> OutputFormatterFunc { get; set; }
    }
}
