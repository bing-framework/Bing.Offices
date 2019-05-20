using System.Collections.Generic;
using Bing.Offices.Enums;

namespace Bing.Offices.Models.Excels
{
    /// <summary>
    /// 导入选项
    /// </summary>
    public class ImportOptions
    {
        /// <summary>
        /// 文件地址
        /// </summary>
        public string FileUrl { get; set; }

        /// <summary>
        /// 工作表索引
        /// </summary>
        public int SheetIndex { get; set; } = 0;

        /// <summary>
        /// 表头行索引
        /// </summary>
        public int HeaderRowIndex { get; set; } = 0;        

        /// <summary>
        /// 数据行起始索引
        /// </summary>
        public int DataRowStartIndex { get; set; } = 1;

        /// <summary>
        /// 映射字典
        /// </summary>
        public Dictionary<string,string> MappingDictionary { get; set; }

        /// <summary>
        /// 校验模式
        /// </summary>
        public ValidateMode ValidateMode { get; set; } = ValidateMode.Stop;
    }
}
