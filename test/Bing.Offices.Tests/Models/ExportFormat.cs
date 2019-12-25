using System;
using System.ComponentModel.DataAnnotations.Schema;
using Bing.Offices.Attributes;

namespace Bing.Offices.Tests.Models
{
    /// <summary>
    /// 导出格式化
    /// </summary>
    public class ExportFormat
    {
        /// <summary>
        /// 系统标识
        /// </summary>
        [ColumnName("系统标识")]
        public string Id { get; set; }

        /// <summary>
        /// 名称
        /// </summary>
        [ColumnName("名称")]
        public string Name { get; set; }

        /// <summary>
        /// 索引值
        /// </summary>
        [ColumnName("索引值")]
        public int Index { get; set; }

        /// <summary>
        /// 忽略属性
        /// </summary>
        [ExcelIgnore]
        [ColumnName("忽略属性")]
        public string IgnoreProperty { get; set; }

        /// <summary>
        /// 忽略映射属性
        /// </summary>
        [NotMapped]
        [ColumnName("忽略映射属性")]
        public string NotMappedProperty { get; set; }

        /// <summary>
        /// 金钱
        /// </summary>
        [DataFormat("0.00")]
        [ColumnName("金钱")]
        public decimal Money { get; set; }

        [DataFormat("yyyy年MM月dd日 hh:mm:ss")]
        [ColumnName("创建时间")]
        public DateTime CreateTime { get; set; }
    }
}
