using System;
using Bing.Offices.Metadata;

namespace Bing.Offices.Attributes
{
    /// <summary>
    /// 工作表 特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = true)]
    public sealed class SheetAttribute : Attribute
    {
        /// <summary>
        /// 索引
        /// </summary>
        public int Index
        {
            get => SheetMetadata.Index;
            set => SheetMetadata.Index = value;
        }

        /// <summary>
        /// 名称
        /// </summary>
        public string Name
        {
            get => SheetMetadata.Name;
            set => SheetMetadata.Name = value;
        }

        /// <summary>
        /// 标题
        /// </summary>
        public string Title
        {
            get => SheetMetadata.Title;
            set => SheetMetadata.Title = value;
        }

        /// <summary>
        /// 起始行索引
        /// </summary>
        public int StartRowIndex
        {
            get => SheetMetadata.StartRowIndex;
            set => SheetMetadata.StartRowIndex = value;
        }

        /// <summary>
        /// 标题行索引
        /// </summary>
        public int HeaderRowIndex => SheetMetadata.HeaderRowIndex;

        /// <summary>
        /// 工作表元数据
        /// </summary>
        internal SheetMetadata SheetMetadata { get; }

        /// <summary>
        /// 初始化一个<see cref="SheetAttribute"/>类型的实例
        /// </summary>
        public SheetAttribute() => SheetMetadata = new SheetMetadata();
    }
}
