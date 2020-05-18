using System.Collections.Generic;

namespace Bing.Offices.Imports
{
    /// <summary>
    /// 导入列头信息
    /// </summary>
    public class ImportHeaderInfo
    {
        /// <summary>
        /// 是否必填
        /// </summary>
        public bool IsRequired { get; set; }

        /// <summary>
        /// 列名称
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        /// 列属性
        /// </summary>
        public IImportHeader Header { get; set; }

        /// <summary>
        /// 图片属性
        /// </summary>
        public IImportImageField ImageField { get; set; }

        /// <summary>
        /// 映射值字典
        /// </summary>
        public Dictionary<string, dynamic> MappingValues { get; set; } = new Dictionary<string, dynamic>();

        /// <summary>
        /// 是否存在
        /// </summary>
        public bool IsExist { get; set; }
    }
}
