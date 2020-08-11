using System;
using Bing.Offices.Metadata;

namespace Bing.Offices.Attributes
{
    /// <summary>
    /// 列 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public sealed class ColumnAttribute : Attribute
    {
        /// <summary>
        /// 索引
        /// </summary>
        public int Index
        {
            get => PropertyMetadata.Column.ColumnIndex;
            set => PropertyMetadata.Column.ColumnIndex = value;
        }

        /// <summary>
        /// 标题
        /// </summary>
        public string Title
        {
            get => PropertyMetadata.Column.Title;
            set => PropertyMetadata.Column.Title = value;
        }

        /// <summary>
        /// 格式化
        /// </summary>
        public string Formatter
        {
            get => PropertyMetadata.Column.Formatter;
            set => PropertyMetadata.Column.Formatter = value;
        }

        /// <summary>
        /// 是否忽略
        /// </summary>
        public bool IsIgnored
        {
            get => PropertyMetadata.IsIgnored;
            set => PropertyMetadata.IsIgnored = value;
        }

        /// <summary>
        /// 列宽
        /// </summary>
        public int Width
        {
            get => PropertyMetadata.Column.Style.Width;
            set => PropertyMetadata.Column.Style.SetWidth(value);
        }

        /// <summary>
        /// 属性元数据
        /// </summary>
        internal PropertyMetadata PropertyMetadata { get; }

        /// <summary>
        /// 初始化一个<see cref="ColumnAttribute"/>类型的实例
        /// </summary>
        public ColumnAttribute() => PropertyMetadata = new PropertyMetadata();

        /// <summary>
        /// 初始化一个<see cref="ColumnAttribute"/>类型的实例
        /// </summary>
        /// <param name="index">索引</param>
        public ColumnAttribute(int index) => PropertyMetadata = new PropertyMetadata
        {
            Column = new ColumnMetadata
            {
                ColumnIndex = index
            }
        };

        /// <summary>
        /// 初始化一个<see cref="ColumnAttribute"/>类型的实例
        /// </summary>
        /// <param name="title">标题</param>
        public ColumnAttribute(string title) => PropertyMetadata = new PropertyMetadata
        {
            Column = new ColumnMetadata
            {
                Title = title
            }
        };
    }
}
